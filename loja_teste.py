#!/usr/bin/env python3
"""
Sistema de Gestão de Estoque e Vendas - ADR INFO
Sistema Completo Consolidado
Desenvolvido com PyQt6

Para executar o sistema:
python loja_teste.py
"""

import sys
import os
import sqlite3
import tempfile
import subprocess
import threading
import json
import asyncio
from datetime import datetime
from typing import List, Optional, Dict
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTabWidget, QLabel,
                             QTableWidget, QTableWidgetItem, QLineEdit, 
                             QFormLayout, QMessageBox, QSpinBox, QDoubleSpinBox,
                             QHeaderView, QDialog, QDialogButtonBox, QDateEdit,
                             QComboBox, QRadioButton, QButtonGroup, QCheckBox)
from PyQt6.QtCore import Qt, QDate, QObject, pyqtSignal
from PyQt6.QtGui import QFont, QIcon

# Importação do Telethon
try:
    from telethon import TelegramClient
    from telethon.errors import SessionPasswordNeededError, PhoneCodeInvalidError
    TELETHON_AVAILABLE = True
except ImportError:
    TELETHON_AVAILABLE = False

# Bibliotecas para Windows (impressão)
if sys.platform == "win32":
    try:
        import win32print
        import win32api
        WIN32_AVAILABLE = True
    except ImportError:
        WIN32_AVAILABLE = False
else:
    WIN32_AVAILABLE = False

# Importação da FPDF
try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

# ========================= DATABASE INITIALIZATION =========================
def init_db():
    """Inicializa o banco de dados com todas as tabelas necessárias"""
    conn = sqlite3.connect('stock_manager.db')
    cursor = conn.cursor()

    # Tabela de produtos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            purchase_price REAL NOT NULL,
            sale_price REAL NOT NULL,
            quantity INTEGER NOT NULL
        )
    ''')

    # Tabela de vendas
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL,
            total_price REAL NOT NULL,
            sale_date TEXT NOT NULL,
            amount_received REAL DEFAULT 0,
            change_amount REAL DEFAULT 0,
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
    ''')

    # Adicionar novas colunas se elas não existirem (para compatibilidade com BDs existentes)
    try:
        cursor.execute('ALTER TABLE sales ADD COLUMN amount_received REAL DEFAULT 0')
    except sqlite3.OperationalError:
        pass  # Coluna já existe

    try:
        cursor.execute('ALTER TABLE sales ADD COLUMN change_amount REAL DEFAULT 0')
    except sqlite3.OperationalError:
        pass  # Coluna já existe
    
    try:
        cursor.execute('ALTER TABLE sales ADD COLUMN payment_type TEXT DEFAULT "Dinheiro"')
    except sqlite3.OperationalError:
        pass  # Coluna já existe

    # Tabela de gastos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            description TEXT NOT NULL,
            amount REAL NOT NULL,
            expense_date TEXT NOT NULL,
            category TEXT DEFAULT 'Geral',
            product_id INTEGER DEFAULT NULL,
            quantity INTEGER DEFAULT NULL,
            expense_type TEXT DEFAULT 'monetary',
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
    ''')

    # Adicionar novas colunas para gastos com produtos se não existirem
    try:
        cursor.execute('ALTER TABLE expenses ADD COLUMN product_id INTEGER DEFAULT NULL')
    except sqlite3.OperationalError:
        pass  # Coluna já existe

    try:
        cursor.execute('ALTER TABLE expenses ADD COLUMN quantity INTEGER DEFAULT NULL')
    except sqlite3.OperationalError:
        pass  # Coluna já existe

    try:
        cursor.execute('ALTER TABLE expenses ADD COLUMN expense_type TEXT DEFAULT "monetary"')
    except sqlite3.OperationalError:
        pass  # Coluna já existe

    # Tabela de serviços
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            description TEXT NOT NULL,
            value REAL NOT NULL,
            service_date TEXT NOT NULL,
            category TEXT DEFAULT 'Geral',
            customer_name TEXT DEFAULT NULL
        )
    ''')

    conn.commit()
    conn.close()

# ========================= UTILITY CLASSES =========================
class InvoiceGenerator:
    """Gerador de notas fiscais em PDF"""
    def __init__(self):
        self.invoices_dir = "invoices"
        if not os.path.exists(self.invoices_dir):
            os.makedirs(self.invoices_dir)
    
    def generate_cart_invoice(self, cart_items, sales, payment_data=None):
        """Gera uma nota fiscal consolidada para múltiplos produtos do carrinho."""
        if not FPDF_AVAILABLE:
            print("⚠️ FPDF não disponível. Nota fiscal não será gerada.")
            return None
            
        # Criar PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        
        # Título
        pdf.cell(0, 10, 'ADR INFO', 0, 1, 'C')
        pdf.ln(10)
        
        # Data e hora (usar da primeira venda)
        pdf.set_font('Arial', '', 12)
        sale_date = sales[0].sale_date if sales else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        pdf.cell(0, 10, f'Data e Hora: {sale_date}', 0, 1)
        pdf.ln(5)
        
        # Cabeçalho da tabela
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(80, 10, 'Produto', 1, 0, 'C')
        pdf.cell(30, 10, 'Quantidade', 1, 0, 'C')
        pdf.cell(30, 10, 'Valor Unit.', 1, 0, 'C')
        pdf.cell(30, 10, 'Subtotal', 1, 1, 'C')
        
        # Dados dos produtos
        pdf.set_font('Arial', '', 10)
        total_original = 0
        
        for cart_item in cart_items:
            pdf.cell(80, 10, cart_item.product.name, 1, 0)
            pdf.cell(30, 10, str(cart_item.quantity), 1, 0, 'C')
            pdf.cell(30, 10, f'R$ {cart_item.product.sale_price:.2f}', 1, 0, 'C')
            pdf.cell(30, 10, f'R$ {cart_item.subtotal:.2f}', 1, 1, 'C')
            total_original += cart_item.subtotal
        
        pdf.ln(10)
        
        # Total geral (com desconto, se houver)
        pdf.set_font('Arial', 'B', 12)
        
        total_final = payment_data.get('total_with_discount', total_original) if payment_data else total_original
        discount_amount = payment_data.get('discount', 0) if payment_data else 0
        
        if discount_amount > 0:
            pdf.cell(0, 10, f'SUBTOTAL: R$ {total_original:.2f}', 0, 1, 'R')
            pdf.cell(0, 10, f'DESCONTO: -R$ {discount_amount:.2f}', 0, 1, 'R')
            pdf.cell(0, 10, f'TOTAL GERAL: R$ {total_final:.2f}', 0, 1, 'R')
        else:
            pdf.cell(0, 10, f'TOTAL GERAL: R$ {total_final:.2f}', 0, 1, 'R')
        
        # Informações de pagamento
        if payment_data:
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, 'INFORMAÇÕES DE PAGAMENTO:', 0, 1)
            
            pdf.set_font('Arial', '', 10)
            pdf.cell(0, 8, f'Valor Recebido: R$ {payment_data["amount_received"]:.2f}', 0, 1)
            
            if payment_data['change'] > 0:
                pdf.cell(0, 8, f'Troco: R$ {payment_data["change"]:.2f}', 0, 1)
            else:
                pdf.cell(0, 8, 'Troco: R$ 0,00', 0, 1)
        
        pdf.ln(10)
        
        # Rodapé
        pdf.set_font('Arial', 'I', 8)
        pdf.cell(0, 5, 'Obrigado pela preferencia!', 0, 1, 'C')
        
        # Salvar arquivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"nota_fiscal_carrinho_{timestamp}.pdf"
        filepath = os.path.join(self.invoices_dir, filename)
        
        pdf.output(filepath)

        # Enviar PDF automaticamente ao Telegram se configurado
        try:
            telegram_manager.send_pdf_if_enabled(filepath)
        except Exception:
            pass

        return filepath

class PrinterManager:
    """Gerenciador de impressoras para integração com nota fiscal"""
    
    def __init__(self):
        self.selected_printer = None
        self.auto_print_enabled = True
        
    def get_available_printers(self) -> List[str]:
        """Retorna lista de impressoras disponíveis no sistema"""
        printers = []
        
        if sys.platform == "win32":
            # Windows - usar win32print se disponível
            if WIN32_AVAILABLE:
                try:
                    printers_info = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
                    for printer in printers_info:
                        printers.append(printer['pName'])
                    
                    # Adicionar impressoras de rede
                    try:
                        network_printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_CONNECTIONS, None, 1)
                        for printer in network_printers:
                            printers.append(printer['pName'])
                    except:
                        pass
                        
                except Exception as e:
                    print(f"Erro ao listar impressoras Windows: {e}")
        
        # Se não conseguiu encontrar impressoras, adicionar Microsoft Print to PDF como fallback
        if not printers and sys.platform == "win32":
            virtual_printers = [
                "Microsoft Print to PDF",
                "Microsoft XPS Document Writer",
            ]
            for vprinter in virtual_printers:
                try:
                    # Tentar abrir a impressora para ver se existe
                    if WIN32_AVAILABLE:
                        handle = win32print.OpenPrinter(vprinter)
                        win32print.ClosePrinter(handle)
                        printers.append(vprinter)
                        break
                except:
                    continue
        
        return printers
    
    def get_default_printer(self) -> Optional[str]:
        """Retorna a impressora padrão do sistema"""
        if sys.platform == "win32" and WIN32_AVAILABLE:
            try:
                return win32print.GetDefaultPrinter()
            except:
                return None
        return None
    
    def set_selected_printer(self, printer_name: str) -> bool:
        """Define a impressora selecionada"""
        available_printers = self.get_available_printers()
        if printer_name in available_printers:
            self.selected_printer = printer_name
            return True
        return False
    
    def set_auto_print(self, enabled: bool):
        """Habilita/desabilita impressão automática"""
        self.auto_print_enabled = enabled
    
    def should_auto_print(self) -> bool:
        """Verifica se deve imprimir automaticamente"""
        return self.auto_print_enabled and self.selected_printer is not None
    
    def get_printer_status(self, printer_name: str = None) -> Dict[str, any]:
        """Retorna o status da impressora"""
        if printer_name is None:
            printer_name = self.selected_printer
        
        status = {
            'name': printer_name,
            'available': False,
            'is_default': False,
            'status_text': 'Não disponível'
        }
        
        if not printer_name:
            return status
        
        # Verificar se está disponível
        available_printers = self.get_available_printers()
        if printer_name in available_printers:
            status['available'] = True
            status['status_text'] = 'Disponível'
        
        # Verificar se é a impressora padrão
        default_printer = self.get_default_printer()
        if printer_name == default_printer:
            status['is_default'] = True
        
        return status
    
    def print_pdf(self, pdf_path: str, printer_name: str = None) -> bool:
        """Imprime um arquivo PDF na impressora especificada"""
        if printer_name is None:
            printer_name = self.selected_printer
        
        if not printer_name:
            print("Nenhuma impressora selecionada")
            return False
        
        if not os.path.exists(pdf_path):
            print(f"Arquivo PDF não encontrado: {pdf_path}")
            return False
        
        try:
            if sys.platform == "win32":
                # Método simples para Windows
                if "Microsoft Print to PDF" in printer_name:
                    # Para Microsoft Print to PDF, abrir o arquivo
                    import subprocess
                    subprocess.run(['start', '', pdf_path], shell=True, check=True)
                    return True
                else:
                    # Para impressoras físicas, tentar comando de impressão direto
                    import subprocess
                    cmd = f'start /min "" "{pdf_path}" /print /d:"{printer_name}"'
                    result = subprocess.run(cmd, shell=True, capture_output=True, timeout=30)
                    return result.returncode == 0
            else:
                print(f"Sistema operacional não suportado para impressão: {sys.platform}")
                return False
                
        except Exception as e:
            print(f"Erro ao imprimir PDF: {e}")
            return False
    
    def print_test_page(self, printer_name: str = None) -> bool:
        """Imprime uma página de teste"""
        if printer_name is None:
            printer_name = self.selected_printer
        
        if not printer_name or not FPDF_AVAILABLE:
            return False
        
        try:
            # Criar PDF de teste temporário
            test_pdf = FPDF()
            test_pdf.add_page()
            
            test_pdf.set_font('Arial', 'B', 16)
            test_pdf.cell(0, 10, 'PAGINA DE TESTE', 0, 1, 'C')
            test_pdf.ln(10)
            
            test_pdf.set_font('Arial', '', 12)
            test_pdf.cell(0, 10, f'Impressora: {printer_name}', 0, 1)
            test_pdf.cell(0, 10, f'Data: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1)
            test_pdf.cell(0, 10, 'Sistema de Gestao ADR INFO', 0, 1)
            test_pdf.ln(10)
            test_pdf.cell(0, 10, 'Se voce pode ver esta pagina, a impressora esta funcionando!', 0, 1)
            
            # Salvar PDF temporário
            temp_path = os.path.join(tempfile.gettempdir(), f'teste_impressao_{int(datetime.now().timestamp())}.pdf')
            test_pdf.output(temp_path)
            
            # Imprimir página de teste
            success = self.print_pdf(temp_path, printer_name)
            return success
            
        except Exception as e:
            print(f"Erro ao criar página de teste: {e}")
            return False
        finally:
            # Limpar arquivo temporário
            try:
                if 'temp_path' in locals():
                    os.unlink(temp_path)
            except:
                pass

# ========================= MODEL CLASSES =========================
class CartItem:
    """Representa um item no carrinho de compras"""
    def __init__(self, product, quantity):
        self.product = product
        self.quantity = quantity
        self.subtotal = product.sale_price * quantity
        self.item_type = "product"  # Tipo do item: "product" ou "service"
    
    def update_quantity(self, new_quantity):
        """Atualiza a quantidade e recalcula o subtotal"""
        if new_quantity <= self.product.quantity:  # Verifica se há estoque suficiente
            self.quantity = new_quantity
            self.subtotal = self.product.sale_price * new_quantity
            return True
        return False

class ServiceCartItem:
    """Representa um serviço no carrinho de compras"""
    def __init__(self, service):
        self.service = service
        self.quantity = 1  # Serviços sempre têm quantidade 1
        self.subtotal = service.value
        self.item_type = "service"  # Tipo do item: "product" ou "service"
        # Criar um produto fictício para compatibilidade com o sistema existente
        self.product = type('obj', (object,), {
            'name': f"SERVICO: {service.description}",
            'sale_price': service.value,
            'id': f"service_{service.id}",
            'quantity': 999  # Serviços não têm limitação de estoque
        })()
    
    def update_quantity(self, new_quantity):
        """Serviços sempre mantêm quantidade 1"""
        return False  # Não permite alterar quantidade de serviços

class Product:
    def __init__(self, name, purchase_price, sale_price, quantity, id=None):
        self.id = id
        self.name = name
        self.purchase_price = purchase_price
        self.sale_price = sale_price
        self.quantity = quantity

    def save(self):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            if self.id:
                cursor.execute("UPDATE products SET name=?, purchase_price=?, sale_price=?, quantity=? WHERE id=?",
                               (self.name, self.purchase_price, self.sale_price, self.quantity, self.id))
            else:
                cursor.execute("INSERT INTO products (name, purchase_price, sale_price, quantity) VALUES (?, ?, ?, ?)",
                               (self.name, self.purchase_price, self.sale_price, self.quantity))
                self.id = cursor.lastrowid
            conn.commit()

    @staticmethod
    def get_all():
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM products")
            products = []
            for row in cursor.fetchall():
                products.append(Product(row[1], row[2], row[3], row[4], row[0]))
            return products

    @staticmethod
    def get_by_id(product_id):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM products WHERE id=?", (product_id,))
            row = cursor.fetchone()
            if row:
                return Product(row[1], row[2], row[3], row[4], row[0])
            return None

    @staticmethod
    def delete(product_id):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM products WHERE id=?", (product_id,))
            conn.commit()
    
    @staticmethod
    def get_total_inventory_value():
        """Calcula o valor total do estoque (custo de aquisição)"""
        products = Product.get_all()
        total_value = sum(product.purchase_price * product.quantity for product in products)
        return total_value
    
    @staticmethod
    def get_daily_profit(date_str):
        """Calcula o lucro líquido das vendas de uma data específica, descontando os gastos"""
        sales = Sale.get_sales_by_date(date_str)
        total_profit = 0
        
        for sale in sales:
            product = Product.get_by_id(sale.product_id)
            if product:
                # Lucro = (preço de venda - preço de compra) * quantidade vendida
                profit_per_unit = product.sale_price - product.purchase_price
                total_profit += profit_per_unit * sale.quantity
        
        # Adicionar receitas de serviços do dia
        daily_services = Service.get_total_services_by_date(date_str)
        total_profit += daily_services
        
        # Descontar gastos do dia
        daily_expenses = Expense.get_total_expenses_by_date(date_str)
        net_profit = total_profit - daily_expenses
        
        return net_profit

class Sale:
    def __init__(self, product_id, quantity, total_price, sale_date=None, id=None, amount_received=0, change_amount=0, payment_type="Dinheiro"):
        self.id = id
        self.product_id = product_id
        self.quantity = quantity
        self.total_price = total_price
        self.sale_date = sale_date if sale_date else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.amount_received = amount_received
        self.change_amount = change_amount
        self.payment_type = payment_type

    def save(self):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            if self.id:
                cursor.execute("UPDATE sales SET product_id=?, quantity=?, total_price=?, sale_date=?, amount_received=?, change_amount=?, payment_type=? WHERE id=?",
                               (self.product_id, self.quantity, self.total_price, self.sale_date, self.amount_received, self.change_amount, self.payment_type, self.id))
            else:
                cursor.execute("INSERT INTO sales (product_id, quantity, total_price, sale_date, amount_received, change_amount, payment_type) VALUES (?, ?, ?, ?, ?, ?, ?)",
                               (self.product_id, self.quantity, self.total_price, self.sale_date, self.amount_received, self.change_amount, self.payment_type))
                self.id = cursor.lastrowid
            conn.commit()

    @staticmethod
    def get_all():
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM sales")
            sales = []
            for row in cursor.fetchall():
                # Verificar se as novas colunas existem (compatibilidade com BD antigo)
                amount_received = row[5] if len(row) > 5 else 0
                change_amount = row[6] if len(row) > 6 else 0
                payment_type = row[7] if len(row) > 7 else "Dinheiro"
                sales.append(Sale(row[1], row[2], row[3], row[4], row[0], amount_received, change_amount, payment_type))
            return sales

    @staticmethod
    def get_by_id(sale_id):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM sales WHERE id=?", (sale_id,))
            row = cursor.fetchone()
            if row:
                # Verificar se as novas colunas existem (compatibilidade com BD antigo)
                amount_received = row[5] if len(row) > 5 else 0
                change_amount = row[6] if len(row) > 6 else 0
                payment_type = row[7] if len(row) > 7 else "Dinheiro"
                return Sale(row[1], row[2], row[3], row[4], row[0], amount_received, change_amount, payment_type)
            return None

    @staticmethod
    def get_sales_by_product_id(product_id):
        conn = sqlite3.connect("stock_manager.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM sales WHERE product_id=?", (product_id,))
        sales = []
        for row in cursor.fetchall():
            # Verificar se as novas colunas existem (compatibilidade com BD antigo)
            amount_received = row[5] if len(row) > 5 else 0
            change_amount = row[6] if len(row) > 6 else 0
            sales.append(Sale(row[1], row[2], row[3], row[4], row[0], amount_received, change_amount))
        conn.close()
        return sales
    
    @staticmethod
    def get_sales_by_date(date_str):
        """Retorna vendas de uma data específica (formato: YYYY-MM-DD)"""
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM sales WHERE DATE(sale_date) = ?", (date_str,))
            sales = []
            for row in cursor.fetchall():
                # Verificar se as novas colunas existem (compatibilidade com BD antigo)
                amount_received = row[5] if len(row) > 5 else 0
                change_amount = row[6] if len(row) > 6 else 0
                payment_type = row[7] if len(row) > 7 else "Dinheiro"
                sales.append(Sale(row[1], row[2], row[3], row[4], row[0], amount_received, change_amount, payment_type))
            return sales
    
    @staticmethod
    def process_cart_sale(cart_items, payment_data=None):
        """
        Processa uma venda com múltiplos itens do carrinho.
        Retorna uma lista de vendas criadas.
        """
        sales = []
        total_sale_amount = sum(item.subtotal for item in cart_items)
        
        # Get total with discount from payment data
        total_with_discount = payment_data.get('total_with_discount', total_sale_amount) if payment_data else total_sale_amount
        
        # Calcular valores proporcionais de pagamento para cada item
        for cart_item in cart_items:
            # Calcular proporção deste item no total
            item_proportion = cart_item.subtotal / total_sale_amount if total_sale_amount > 0 else 0
            
            # Calculate discounted total for this item
            item_discounted_total = total_with_discount * item_proportion
            
            # Calcular valores proporcionais de pagamento
            if payment_data:
                item_amount_received = payment_data['amount_received'] * item_proportion
                item_change = payment_data['change'] * item_proportion
            else:
                item_amount_received = item_discounted_total
                item_change = 0
            
            # Obter tipo de pagamento
            payment_type = payment_data.get('payment_type', 'Dinheiro') if payment_data else 'Dinheiro'
            
            # Verificar se é serviço ou produto
            if hasattr(cart_item, 'item_type') and cart_item.item_type == "service":
                # Para serviços, usar um ID especial e não atualizar estoque
                product_id = f"service_{cart_item.service.id}"
                # Criar venda para o serviço
                sale = Sale(
                    product_id, 
                    cart_item.quantity, 
                    item_discounted_total,
                    amount_received=item_amount_received,
                    change_amount=item_change,
                    payment_type=payment_type
                )
                sale.save()
                # Não atualizar estoque para serviços
            else:
                # Criar venda para produto normal
                sale = Sale(
                    cart_item.product.id, 
                    cart_item.quantity, 
                    item_discounted_total,
                    amount_received=item_amount_received,
                    change_amount=item_change,
                    payment_type=payment_type
                )
                sale.save()
                
                # Atualizar estoque do produto (apenas para produtos)
                if hasattr(cart_item.product, 'id') and not str(cart_item.product.id).startswith('service_'):
                    product = Product.get_by_id(cart_item.product.id)
                    if product:
                        product.quantity -= cart_item.quantity
                        product.save()
            
            sales.append(sale)
        
        # Enviar BD automaticamente ao Telegram se configurado
        try:
            telegram_manager.send_db_if_enabled()
        except Exception:
            pass

        return sales

class Expense:
    def __init__(self, description, amount, expense_date=None, category="Geral", id=None, product_id=None, quantity=None, expense_type="monetary"):
        self.id = id
        self.description = description
        self.amount = amount
        self.expense_date = expense_date if expense_date else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.category = category
        self.product_id = product_id
        self.quantity = quantity
        self.expense_type = expense_type  # "monetary" ou "product"

    def save(self):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            if self.id:
                cursor.execute("UPDATE expenses SET description=?, amount=?, expense_date=?, category=?, product_id=?, quantity=?, expense_type=? WHERE id=?",
                               (self.description, self.amount, self.expense_date, self.category, self.product_id, self.quantity, self.expense_type, self.id))
            else:
                cursor.execute("INSERT INTO expenses (description, amount, expense_date, category, product_id, quantity, expense_type) VALUES (?, ?, ?, ?, ?, ?, ?)",
                               (self.description, self.amount, self.expense_date, self.category, self.product_id, self.quantity, self.expense_type))
                self.id = cursor.lastrowid
                
                # Se for gasto com produto, descontar do estoque DENTRO DA MESMA TRANSAÇÃO
                if self.expense_type == "product" and self.product_id and self.quantity:
                    # Buscar o produto atual
                    cursor.execute("SELECT * FROM products WHERE id=?", (self.product_id,))
                    row = cursor.fetchone()
                    if row and row[4] >= self.quantity:  # row[4] é a quantidade
                        # Atualizar diretamente na mesma transação
                        new_quantity = row[4] - self.quantity
                        cursor.execute("UPDATE products SET quantity=? WHERE id=?", (new_quantity, self.product_id))
                        
            conn.commit()

    @staticmethod
    def get_all():
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM expenses ORDER BY expense_date DESC")
            expenses = []
            for row in cursor.fetchall():
                # Verificar se as novas colunas existem (compatibilidade com BD antigo)
                product_id = row[5] if len(row) > 5 else None
                quantity = row[6] if len(row) > 6 else None
                expense_type = row[7] if len(row) > 7 else "monetary"
                expenses.append(Expense(row[1], row[2], row[3], row[4], row[0], product_id, quantity, expense_type))
            return expenses

    @staticmethod
    def get_by_id(expense_id):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM expenses WHERE id=?", (expense_id,))
            row = cursor.fetchone()
            if row:
                # Verificar se as novas colunas existem (compatibilidade com BD antigo)
                product_id = row[5] if len(row) > 5 else None
                quantity = row[6] if len(row) > 6 else None
                expense_type = row[7] if len(row) > 7 else "monetary"
                return Expense(row[1], row[2], row[3], row[4], row[0], product_id, quantity, expense_type)
            return None

    @staticmethod
    def delete(expense_id):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM expenses WHERE id=?", (expense_id,))
            conn.commit()
    
    @staticmethod
    def get_expenses_by_date(date_str):
        """Retorna gastos de uma data específica (formato: YYYY-MM-DD)"""
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM expenses WHERE DATE(expense_date) = ?", (date_str,))
            expenses = []
            for row in cursor.fetchall():
                # Verificar se as novas colunas existem (compatibilidade com BD antigo)
                product_id = row[5] if len(row) > 5 else None
                quantity = row[6] if len(row) > 6 else None
                expense_type = row[7] if len(row) > 7 else "monetary"
                expenses.append(Expense(row[1], row[2], row[3], row[4], row[0], product_id, quantity, expense_type))
            return expenses
    
    @staticmethod
    def get_total_expenses_by_date(date_str):
        """Retorna o total de gastos de uma data específica"""
        expenses = Expense.get_expenses_by_date(date_str)
        return sum(expense.amount for expense in expenses)

class Service:
    def __init__(self, description, value, service_date=None, category="Geral", customer_name=None, id=None):
        self.id = id
        self.description = description
        self.value = value
        self.service_date = service_date if service_date else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.category = category
        self.customer_name = customer_name

    def save(self):
        conn = sqlite3.connect("stock_manager.db")
        cursor = conn.cursor()
        if self.id:
            cursor.execute("UPDATE services SET description=?, value=?, service_date=?, category=?, customer_name=? WHERE id=?",
                           (self.description, self.value, self.service_date, self.category, self.customer_name, self.id))
        else:
            cursor.execute("INSERT INTO services (description, value, service_date, category, customer_name) VALUES (?, ?, ?, ?, ?)",
                           (self.description, self.value, self.service_date, self.category, self.customer_name))
            self.id = cursor.lastrowid
        conn.commit()
        conn.close()

    @staticmethod
    def get_all():
        conn = sqlite3.connect("stock_manager.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM services ORDER BY service_date DESC")
        services = []
        for row in cursor.fetchall():
            services.append(Service(row[1], row[2], row[3], row[4], row[5], row[0]))
        conn.close()
        return services

    @staticmethod
    def get_by_id(service_id):
        conn = sqlite3.connect("stock_manager.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM services WHERE id=?", (service_id,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return Service(row[1], row[2], row[3], row[4], row[5], row[0])
        return None

    @staticmethod
    def delete(service_id):
        conn = sqlite3.connect("stock_manager.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM services WHERE id=?", (service_id,))
        conn.commit()
        conn.close()
    
    @staticmethod
    def get_services_by_date(date_str):
        """Retorna serviços de uma data específica (formato: YYYY-MM-DD)"""
        conn = sqlite3.connect("stock_manager.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM services WHERE DATE(service_date) = ?", (date_str,))
        services = []
        for row in cursor.fetchall():
            services.append(Service(row[1], row[2], row[3], row[4], row[5], row[0]))
        conn.close()
        return services
    
    @staticmethod
    def get_total_services_by_date(date_str):
        """Retorna o total de receitas de serviços de uma data específica"""
        services = Service.get_services_by_date(date_str)
        return sum(service.value for service in services)

# ========================= DIALOG CLASSES =========================
class ProductDialog(QDialog):
    def __init__(self, product=None, parent=None):
        super().__init__(parent)
        self.product = product
        self.setWindowTitle("📦 Cadastro de Produto" if not product else "✏️ Editar Produto")
        self.setModal(True)
        self.resize(400, 300)
        
        layout = QVBoxLayout()
        
        # Título
        title_label = QLabel("📦 Gestão de Produtos 📦")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #2E8B57; margin: 10px; padding: 10px; background-color: #F0FFF0; border-radius: 5px;")
        layout.addWidget(title_label)
        
        # Formulário
        form_layout = QFormLayout()
        
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Ex: iPhone 13, Película...")
        
        self.purchase_price_edit = QDoubleSpinBox()
        self.purchase_price_edit.setMaximum(999999.99)
        self.purchase_price_edit.setDecimals(2)
        self.purchase_price_edit.setPrefix("R$ ")
        
        self.sale_price_edit = QDoubleSpinBox()
        self.sale_price_edit.setMaximum(999999.99)
        self.sale_price_edit.setDecimals(2)
        self.sale_price_edit.setPrefix("R$ ")
        
        self.quantity_edit = QSpinBox()
        self.quantity_edit.setMaximum(999999)
        self.quantity_edit.setSuffix(" unid.")
        
        if product:
            self.name_edit.setText(product.name)
            self.purchase_price_edit.setValue(product.purchase_price)
            self.sale_price_edit.setValue(product.sale_price)
            self.quantity_edit.setValue(product.quantity)
        
        form_layout.addRow("📝 Nome do Produto:", self.name_edit)
        form_layout.addRow("💰 Preço de Compra:", self.purchase_price_edit)
        form_layout.addRow("💵 Preço de Venda:", self.sale_price_edit)
        form_layout.addRow("📊 Quantidade:", self.quantity_edit)
        
        layout.addLayout(form_layout)
        
        # Botões
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | 
                                     QDialogButtonBox.StandardButton.Cancel)
        button_box.button(QDialogButtonBox.StandardButton.Ok).setText("✅ Salvar")
        button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("❌ Cancelar")
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        layout.addWidget(button_box)
        self.setLayout(layout)
    
    def get_product_data(self):
        return {
            'name': self.name_edit.text(),
            'purchase_price': self.purchase_price_edit.value(),
            'sale_price': self.sale_price_edit.value(),
            'quantity': self.quantity_edit.value()
        }

class AddToCartDialog(QDialog):
    def __init__(self, product, parent=None):
        super().__init__(parent)
        self.product = product
        self.setWindowTitle(f"🛒 Adicionar ao Carrinho - {product.name}")
        self.setModal(True)
        self.resize(400, 250)
        
        layout = QVBoxLayout()
        
        # Título
        title_label = QLabel("🛒 Adicionar ao Carrinho 🛒")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #2E8B57; margin: 10px; padding: 10px; background-color: #F0FFF0; border-radius: 5px;")
        layout.addWidget(title_label)
        
        # Informações do produto
        info_label = QLabel(f"📦 Produto: {product.name}\n💵 Preço: R$ {product.sale_price:.2f}\n📊 Estoque: {product.quantity}")
        info_label.setFont(QFont("Arial", 10))
        info_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border-radius: 5px; margin: 5px 0;")
        layout.addWidget(info_label)
        
        # Quantidade a adicionar
        form_layout = QFormLayout()
        self.quantity_edit = QSpinBox()
        self.quantity_edit.setMinimum(1)
        self.quantity_edit.setMaximum(product.quantity)
        self.quantity_edit.setValue(1)
        self.quantity_edit.setSuffix(" unid.")
        
        form_layout.addRow("🔢 Quantidade:", self.quantity_edit)
        layout.addLayout(form_layout)
        
        # Subtotal
        self.subtotal_label = QLabel(f"💰 Subtotal: R$ {product.sale_price:.2f}")
        self.subtotal_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.subtotal_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.subtotal_label.setStyleSheet("color: #228B22; padding: 10px; border: 2px solid #228B22; border-radius: 5px; background-color: #F0FFF0; margin: 5px;")
        layout.addWidget(self.subtotal_label)
        
        # Atualizar subtotal quando quantidade muda
        self.quantity_edit.valueChanged.connect(self.update_subtotal)
        
        # Botões
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | 
                                     QDialogButtonBox.StandardButton.Cancel)
        button_box.button(QDialogButtonBox.StandardButton.Ok).setText("✅ Adicionar")
        button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("❌ Cancelar")
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        layout.addWidget(button_box)
        self.setLayout(layout)
    
    def update_subtotal(self):
        subtotal = self.quantity_edit.value() * self.product.sale_price
        self.subtotal_label.setText(f"💰 Subtotal: R$ {subtotal:.2f}")
    
    def get_quantity(self):
        return self.quantity_edit.value()

class PaymentDialog(QDialog):
    def __init__(self, total_amount, parent=None):
        super().__init__(parent)
        self.original_total = total_amount
        self.total_amount   = total_amount   # será ajustado pelo desconto
        self.amount_received_manual_changed = False  # Track if user manually changed amount received
        self.setWindowTitle("💳 Pagamento")
        self.setModal(True)
        self.resize(440, 480)

        layout = QVBoxLayout()

        # Título
        title_label = QLabel("💳 Finalizar Pagamento 💳")
        title_label.setFont(QFont("Arial", 15, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color:#2E8B57;margin:8px;padding:8px;background-color:#F0FFF0;border-radius:5px;")
        layout.addWidget(title_label)

        # Total original
        self.original_label = QLabel(f"🛒 Subtotal: R$ {total_amount:.2f}")
        self.original_label.setFont(QFont("Arial", 12))
        self.original_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.original_label.setStyleSheet("color:#555;padding:4px;")
        layout.addWidget(self.original_label)

        # ── Desconto ─────────────────────────────────────────
        disc_group = QHBoxLayout()

        disc_type_label = QLabel("🏷️ Desconto:")
        disc_type_label.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        disc_group.addWidget(disc_type_label)

        self.disc_type_combo = QComboBox()
        self.disc_type_combo.addItems(["R$ Valor fixo", "% Porcentagem"])
        self.disc_type_combo.setStyleSheet(
            "padding:5px;border:2px solid #FF8C00;border-radius:5px;font-weight:bold;color:#FF8C00;"
        )
        disc_group.addWidget(self.disc_type_combo)

        self.disc_spin = QDoubleSpinBox()
        self.disc_spin.setMinimum(0)
        self.disc_spin.setMaximum(total_amount)
        self.disc_spin.setDecimals(2)
        self.disc_spin.setValue(0)
        self.disc_spin.setPrefix("R$ ")
        self.disc_spin.setStyleSheet(
            "padding:5px;border:2px solid #FF8C00;border-radius:5px;font-weight:bold;"
            "background-color:#1a0d00;color:#FF8C00;"
        )
        disc_group.addWidget(self.disc_spin)
        layout.addLayout(disc_group)

        # Total com desconto (destaque)
        self.total_label = QLabel(f"💰 Total a Pagar: R$ {total_amount:.2f}")
        self.total_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        self.total_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.total_label.setStyleSheet(
            "color:#2E8B57;padding:8px;border:2px solid #2E8B57;border-radius:5px;"
            "background-color:#F0FFF0;margin:6px;"
        )
        layout.addWidget(self.total_label)

        # Formulário
        form_layout = QFormLayout()

        self.payment_type_combo = QComboBox()
        self.payment_type_combo.addItems(["💰 Dinheiro", "💳 Cartão", "📱 PIX"])
        self.payment_type_combo.setStyleSheet("padding:8px;border:2px solid #2E8B57;border-radius:5px;font-size:12px;")
        self.payment_type_combo.currentTextChanged.connect(self.on_payment_type_changed)
        form_layout.addRow("🏦 Tipo de Pagamento:", self.payment_type_combo)

        self.amount_received_edit = QDoubleSpinBox()
        self.amount_received_edit.setMaximum(999999.99)
        self.amount_received_edit.setDecimals(2)
        self.amount_received_edit.setValue(total_amount)
        self.amount_received_edit.setPrefix("R$ ")
        form_layout.addRow("💵 Valor Recebido:", self.amount_received_edit)
        layout.addLayout(form_layout)

        # Troco
        self.change_label = QLabel("🔄 Troco: R$ 0,00")
        self.change_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.change_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.change_label.setStyleSheet(
            "color:#4169E1;padding:8px;border:2px solid #4169E1;border-radius:5px;"
            "background-color:#F0F8FF;margin:6px;"
        )
        layout.addWidget(self.change_label)

        self.warning_label = QLabel("")
        self.warning_label.setFont(QFont("Arial", 10))
        self.warning_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.warning_label.setStyleSheet("color:red;font-weight:bold;")
        layout.addWidget(self.warning_label)

        layout.addStretch()

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.ok_button = button_box.button(QDialogButtonBox.StandardButton.Ok)
        self.ok_button.setText("✅ Confirmar Pagamento")
        button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("❌ Cancelar")
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)

        # Conectar sinais
        self.disc_spin.valueChanged.connect(self.on_discount_changed)
        self.disc_type_combo.currentIndexChanged.connect(self.on_disc_type_changed)
        self.amount_received_edit.valueChanged.connect(self.on_amount_received_changed)
        self.amount_received_edit.valueChanged.connect(self.update_change)

        self.update_change()

    def on_amount_received_changed(self):
        """Mark that the user manually changed the amount received."""
        self.amount_received_manual_changed = True

    def on_disc_type_changed(self):
        """Ajusta o limite e prefixo quando troca entre R$ e %."""
        if self.disc_type_combo.currentIndex() == 0:  # R$
            self.disc_spin.setMaximum(self.original_total)
            self.disc_spin.setPrefix("R$ ")
        else:  # %
            self.disc_spin.setMaximum(100.0)
            self.disc_spin.setPrefix("% ")
        self.disc_spin.setValue(0)
        self.on_discount_changed()

    def on_discount_changed(self):
        """Recalcula total_amount com desconto e atualiza UI."""
        val = self.disc_spin.value()
        if self.disc_type_combo.currentIndex() == 0:  # R$
            discount = min(val, self.original_total)
        else:  # %
            discount = self.original_total * val / 100.0

        self.total_amount = max(0.0, self.original_total - discount)

        if discount > 0:
            self.total_label.setText(
                f"💰 Total a Pagar: R$ {self.total_amount:.2f}  "
                f"(🏷️ -R$ {discount:.2f})"
            )
            self.total_label.setStyleSheet(
                "color:#CC0000;padding:8px;border:2px solid #CC0000;border-radius:5px;"
                "background-color:#FFF0F0;margin:6px;font-weight:bold;font-size:13px;"
            )
        else:
            self.total_label.setText(f"💰 Total a Pagar: R$ {self.total_amount:.2f}")
            self.total_label.setStyleSheet(
                "color:#2E8B57;padding:8px;border:2px solid #2E8B57;border-radius:5px;"
                "background-color:#F0FFF0;margin:6px;font-weight:bold;font-size:13px;"
            )

        # Ajustar valor recebido
        payment_type = self.payment_type_combo.currentText()
        # Update amount received if:
        # 1. It's Cartão/PIX, OR
        # 2. It's Dinheiro and the user hasn't manually changed the amount yet
        if (payment_type in ["💳 Cartão", "📱 PIX"]) or (payment_type == "💰 Dinheiro" and not self.amount_received_manual_changed):
            self.amount_received_edit.setValue(self.total_amount)

        self.update_change()

    def update_change(self):
        amount_received = self.amount_received_edit.value()
        change = amount_received - self.total_amount

        if change >= 0:
            self.change_label.setText(f"🔄 Troco: R$ {change:.2f}")
            self.change_label.setStyleSheet(
                "color:#4169E1;padding:8px;border:2px solid #4169E1;border-radius:5px;"
                "background-color:#F0F8FF;margin:6px;"
            )
            self.warning_label.setText("")
            self.ok_button.setEnabled(True)
        else:
            self.change_label.setText(f"⚠️ Valor Insuficiente: R$ {abs(change):.2f}")
            self.change_label.setStyleSheet(
                "color:#DC143C;padding:8px;border:2px solid #DC143C;border-radius:5px;"
                "background-color:#FFE4E1;margin:6px;"
            )
            self.warning_label.setText("⚠️ Valor recebido é insuficiente!")
            self.ok_button.setEnabled(False)

    def on_payment_type_changed(self):
        payment_type = self.payment_type_combo.currentText()
        if payment_type in ["💳 Cartão", "📱 PIX"]:
            self.amount_received_edit.setValue(self.total_amount)
            self.amount_received_edit.setEnabled(False)
        else:
            self.amount_received_edit.setEnabled(True)
            # Reset manual changed flag when switching back to Dinheiro
            self.amount_received_manual_changed = False
        self.update_change()

    def get_payment_data(self):
        amount_received = self.amount_received_edit.value()
        change = amount_received - self.total_amount

        payment_type_text = self.payment_type_combo.currentText()
        if payment_type_text.startswith("💰"):
            payment_type = "Dinheiro"
        elif payment_type_text.startswith("💳"):
            payment_type = "Cartão"
        elif payment_type_text.startswith("📱"):
            payment_type = "PIX"
        else:
            payment_type = "Dinheiro"

        return {
            'amount_received':      amount_received,
            'change':               change if change >= 0 else 0,
            'payment_type':         payment_type,
            'total_with_discount':  self.total_amount,       # total já descontado
            'discount':             self.original_total - self.total_amount,
        }

class ExpenseDialog(QDialog):
    def __init__(self, expense=None, parent=None):
        super().__init__(parent)
        self.expense = expense
        self.setWindowTitle("💸 Cadastro de Gasto" if not expense else "✏️ Editar Gasto")
        self.setModal(True)
        self.resize(500, 450)
        
        layout = QVBoxLayout()
        
        # Título
        title_label = QLabel("💸 Gestão de Gastos 💸")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #DC143C; margin: 10px; padding: 10px; background-color: #FFE4E1; border-radius: 5px;")
        layout.addWidget(title_label)
        
        # Seleção do tipo de gasto
        type_title = QLabel("📝 Tipo de Gasto")
        type_title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        layout.addWidget(type_title)
        
        self.expense_type_group = QButtonGroup()
        
        type_layout = QHBoxLayout()
        self.monetary_radio = QRadioButton("💰 Gasto Monetário")
        self.product_radio = QRadioButton("📦 Produto do Estoque")
        
        self.expense_type_group.addButton(self.monetary_radio, 0)
        self.expense_type_group.addButton(self.product_radio, 1)
        
        self.monetary_radio.setChecked(True)  # Padrão
        
        type_layout.addWidget(self.monetary_radio)
        type_layout.addWidget(self.product_radio)
        layout.addLayout(type_layout)
        
        # Conectar mudanças de tipo
        self.monetary_radio.toggled.connect(self.on_type_changed)
        self.product_radio.toggled.connect(self.on_type_changed)
        
        # Formulário comum
        form_layout = QFormLayout()
        
        self.description_edit = QLineEdit()
        self.description_edit.setPlaceholderText("Ex: Conta de luz, Material de limpeza...")
        self.category_edit = QLineEdit()
        self.category_edit.setText("Geral")
        self.category_edit.setPlaceholderText("Ex: Contas, Manutenção, Material...")
        
        form_layout.addRow("📝 Descrição:", self.description_edit)
        form_layout.addRow("🏷️ Categoria:", self.category_edit)
        
        layout.addLayout(form_layout)
        
        # Seção para gasto monetário
        self.monetary_section = QWidget()
        monetary_layout = QFormLayout()
        
        self.amount_edit = QDoubleSpinBox()
        self.amount_edit.setMaximum(999999.99)
        self.amount_edit.setDecimals(2)
        self.amount_edit.setPrefix("R$ ")
        
        monetary_layout.addRow("💰 Valor:", self.amount_edit)
        self.monetary_section.setLayout(monetary_layout)
        layout.addWidget(self.monetary_section)
        
        # Seção para produto do estoque
        self.product_section = QWidget()
        product_layout = QFormLayout()
        
        self.product_combo = QComboBox()
        self.populate_products()
        
        self.quantity_edit = QSpinBox()
        self.quantity_edit.setMinimum(1)
        self.quantity_edit.setMaximum(999999)
        self.quantity_edit.setValue(1)
        self.quantity_edit.setSuffix(" unid.")
        
        # Label para mostrar valor calculado
        self.calculated_value_label = QLabel("💰 Valor: R$ 0,00")
        self.calculated_value_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        self.calculated_value_label.setStyleSheet("color: #2E8B57; padding: 5px; background-color: #F0FFF0; border-radius: 3px;")
        
        product_layout.addRow("📦 Produto:", self.product_combo)
        product_layout.addRow("🔢 Quantidade:", self.quantity_edit)
        product_layout.addRow("", self.calculated_value_label)
        
        self.product_section.setLayout(product_layout)
        layout.addWidget(self.product_section)
        
        # Conectar mudanças para atualizar valor
        self.product_combo.currentIndexChanged.connect(self.update_calculated_value)
        self.quantity_edit.valueChanged.connect(self.update_calculated_value)
        
        # Inicializar visibilidade das seções
        self.on_type_changed()
        
        # Preencher dados se for edição
        if expense:
            self.description_edit.setText(expense.description)
            self.category_edit.setText(expense.category)
            
            if expense.expense_type == "product" and expense.product_id:
                self.product_radio.setChecked(True)
                # Encontrar e selecionar o produto
                for i in range(self.product_combo.count()):
                    if self.product_combo.itemData(i) == expense.product_id:
                        self.product_combo.setCurrentIndex(i)
                        break
                if expense.quantity:
                    self.quantity_edit.setValue(expense.quantity)
            else:
                self.monetary_radio.setChecked(True)
                self.amount_edit.setValue(expense.amount)
        
        # Botões
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | 
                                     QDialogButtonBox.StandardButton.Cancel)
        button_box.button(QDialogButtonBox.StandardButton.Ok).setText("✅ Salvar")
        button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("❌ Cancelar")
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        layout.addWidget(button_box)
        self.setLayout(layout)
        
        # Atualizar valor inicial
        self.update_calculated_value()
    
    def populate_products(self):
        """Preenche o combo box com produtos disponíveis no estoque"""
        self.product_combo.clear()
        self.product_combo.addItem("Selecione um produto...", None)
        
        products = Product.get_all()
        for product in products:
            if product.quantity > 0:  # Apenas produtos com estoque
                self.product_combo.addItem(f"{product.name} (Estoque: {product.quantity})", product.id)
    
    def on_type_changed(self):
        """Alterna a visibilidade das seções baseado no tipo selecionado"""
        is_monetary = self.monetary_radio.isChecked()
        
        self.monetary_section.setVisible(is_monetary)
        self.product_section.setVisible(not is_monetary)
        
        if not is_monetary:
            self.update_calculated_value()
    
    def update_calculated_value(self):
        """Atualiza o valor calculado para gastos com produtos"""
        if self.product_radio.isChecked():
            product_id = self.product_combo.currentData()
            if product_id:
                product = Product.get_by_id(product_id)
                if product:
                    total_value = product.purchase_price * self.quantity_edit.value()
                    self.calculated_value_label.setText(f"💰 Valor: R$ {total_value:.2f}")
                    return
        
        self.calculated_value_label.setText("💰 Valor: R$ 0,00")
    
    def get_expense_data(self):
        """Retorna os dados do gasto baseado no tipo selecionado"""
        if self.monetary_radio.isChecked():
            return {
                'description': self.description_edit.text(),
                'amount': self.amount_edit.value(),
                'category': self.category_edit.text(),
                'expense_type': 'monetary',
                'product_id': None,
                'quantity': None
            }
        else:
            product_id = self.product_combo.currentData()
            if not product_id:
                return None
            
            product = Product.get_by_id(product_id)
            if not product:
                return None
            
            quantity = self.quantity_edit.value()
            total_value = product.purchase_price * quantity
            
            return {
                'description': self.description_edit.text(),
                'amount': total_value,
                'category': self.category_edit.text(),
                'expense_type': 'product',
                'product_id': product_id,
                'quantity': quantity
            }

class ServiceDialog(QDialog):
    def __init__(self, service=None, parent=None):
        super().__init__(parent)
        self.service = service
        self.setWindowTitle("🔧 Cadastro de Serviço" if not service else "✏️ Editar Serviço")
        self.setModal(True)
        self.resize(450, 400)
        
        layout = QVBoxLayout()
        
        # Título
        title_label = QLabel("🔧 Gestão de Serviços 🔧")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #2E8B57; margin: 10px; padding: 10px; background-color: #F0FFF0; border-radius: 5px;")
        layout.addWidget(title_label)
        
        # Formulário
        form_layout = QFormLayout()
        
        self.description_edit = QLineEdit()
        self.description_edit.setPlaceholderText("Ex: Instalação de software, Manutenção de computador...")
        
        self.value_edit = QDoubleSpinBox()
        self.value_edit.setMaximum(999999.99)
        self.value_edit.setDecimals(2)
        self.value_edit.setPrefix("R$ ")
        
        self.category_edit = QLineEdit()
        self.category_edit.setText("Geral")
        self.category_edit.setPlaceholderText("Ex: Manutenção, Instalação, Consultoria...")
        
        self.customer_name_edit = QLineEdit()
        self.customer_name_edit.setPlaceholderText("Nome do cliente (opcional)")
        
        if service:
            self.description_edit.setText(service.description)
            self.value_edit.setValue(service.value)
            self.category_edit.setText(service.category)
            if service.customer_name:
                self.customer_name_edit.setText(service.customer_name)
        
        form_layout.addRow("📝 Descrição do Serviço:", self.description_edit)
        form_layout.addRow("💰 Valor do Serviço:", self.value_edit)
        form_layout.addRow("🏷️ Categoria:", self.category_edit)
        form_layout.addRow("👤 Nome do Cliente:", self.customer_name_edit)
        
        layout.addLayout(form_layout)
        
        # Informação sobre lucro
        info_label = QLabel("💡 O valor informado será adicionado diretamente ao lucro da loja")
        info_label.setFont(QFont("Arial", 9))
        info_label.setStyleSheet("color: #666; padding: 10px; background-color: #f0f0f0; border-radius: 5px; margin: 10px 0;")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Opção para adicionar ao carrinho
        self.add_to_cart_checkbox = QCheckBox("🛒 Adicionar ao carrinho para gerar nota fiscal")
        self.add_to_cart_checkbox.setChecked(True)  # Marcado por padrão
        self.add_to_cart_checkbox.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        self.add_to_cart_checkbox.setStyleSheet("color: #2E8B57; margin: 10px 0;")
        layout.addWidget(self.add_to_cart_checkbox)
        
        # Botões
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | 
                                     QDialogButtonBox.StandardButton.Cancel)
        button_box.button(QDialogButtonBox.StandardButton.Ok).setText("✅ Salvar")
        button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("❌ Cancelar")
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        layout.addWidget(button_box)
        self.setLayout(layout)
    
    def get_service_data(self):
        return {
            'description': self.description_edit.text(),
            'value': self.value_edit.value(),
            'category': self.category_edit.text(),
            'customer_name': self.customer_name_edit.text() or None,
            'add_to_cart': self.add_to_cart_checkbox.isChecked()
        }

# ========================= TELEGRAM MANAGER =========================
class TelegramManager:
    """
    Gerencia a integração com o Telegram via Telethon (MTProto).
    Roda um loop asyncio em thread separada para não travar a UI.
    """

    CONFIG_FILE = "telegram_config.json"
    SESSION_FILE = "rdantas_session"

    def __init__(self):
        self.api_id = None
        self.api_hash = None
        self.phone = None
        self.auto_send_db = False
        self.auto_send_pdf = False
        self._client = None
        self._loop = None
        self._thread = None
        self._phone_code_hash = None
        self._connected = False
        self._status_callback = None   # fn(msg: str)
        self._load_config()
        self._start_loop()

    # ── Config ────────────────────────────────────────────────
    def _load_config(self):
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, "r") as f:
                    cfg = json.load(f)
                self.api_id   = cfg.get("api_id")
                self.api_hash = cfg.get("api_hash")
                self.phone    = cfg.get("phone")
                self.auto_send_db  = cfg.get("auto_send_db", False)
                self.auto_send_pdf = cfg.get("auto_send_pdf", False)
            except Exception:
                pass

    def save_config(self):
        cfg = {
            "api_id":        self.api_id,
            "api_hash":      self.api_hash,
            "phone":         self.phone,
            "auto_send_db":  self.auto_send_db,
            "auto_send_pdf": self.auto_send_pdf,
        }
        try:
            with open(self.CONFIG_FILE, "w") as f:
                json.dump(cfg, f, indent=2)
        except Exception as e:
            print(f"Erro ao salvar config telegram: {e}")

    # ── Asyncio loop em thread separada ───────────────────────
    def _start_loop(self):
        self._loop = asyncio.new_event_loop()
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()

    def _run_loop(self):
        asyncio.set_event_loop(self._loop)
        self._loop.run_forever()

    def _run_async(self, coro):
        """Agenda uma corrotina no loop de background e retorna o Future."""
        return asyncio.run_coroutine_threadsafe(coro, self._loop)

    # ── Criação do cliente ────────────────────────────────────
    def _get_client(self):
        if self._client is None and self.api_id and self.api_hash:
            self._client = TelegramClient(
                self.SESSION_FILE,
                int(self.api_id),
                self.api_hash,
                loop=self._loop
            )
        return self._client

    def _notify(self, msg: str):
        if self._status_callback:
            self._status_callback(msg)

    def set_status_callback(self, fn):
        self._status_callback = fn

    # ── Verificar sessão salva ────────────────────────────────
    def try_auto_connect(self, on_done=None):
        """Tenta reconectar usando sessão salva (sem precisar de código)."""
        if not TELETHON_AVAILABLE or not self.api_id or not self.api_hash:
            return
        fut = self._run_async(self._auto_connect())
        def _cb(f):
            if on_done:
                on_done(self._connected)
        fut.add_done_callback(_cb)

    async def _auto_connect(self):
        try:
            client = self._get_client()
            await client.connect()
            if await client.is_user_authorized():
                self._connected = True
                self._notify("✅ Conectado ao Telegram!")
            else:
                self._connected = False
                self._notify("🔑 Sessão expirada. Faça login novamente.")
        except Exception as e:
            self._connected = False
            self._notify(f"❌ Erro ao conectar: {e}")

    # ── Login: passo 1 (enviar SMS) ───────────────────────────
    def send_code(self, api_id, api_hash, phone, on_done=None):
        self.api_id   = str(api_id).strip()
        self.api_hash = api_hash.strip()
        self.phone    = phone.strip()
        self._client  = None   # forçar recriação com novos dados
        self.save_config()

        fut = self._run_async(self._send_code())
        def _cb(f):
            exc = f.exception()
            if on_done:
                on_done(exc is None, str(exc) if exc else "")
        fut.add_done_callback(_cb)

    async def _send_code(self):
        client = self._get_client()
        await client.connect()
        result = await client.send_code_request(self.phone)
        self._phone_code_hash = result.phone_code_hash
        self._notify("📩 Código enviado! Verifique o Telegram.")

    # ── Login: passo 2 (confirmar código) ─────────────────────
    def confirm_code(self, code, password=None, on_done=None):
        fut = self._run_async(self._confirm_code(code.strip(), password))
        def _cb(f):
            exc = f.exception()
            if on_done:
                on_done(exc is None, str(exc) if exc else "")
        fut.add_done_callback(_cb)

    async def _confirm_code(self, code, password=None):
        client = self._get_client()
        try:
            await client.sign_in(self.phone, code,
                                  phone_code_hash=self._phone_code_hash)
            self._connected = True
            self._notify("✅ Login realizado com sucesso!")
        except SessionPasswordNeededError:
            if password:
                await client.sign_in(password=password)
                self._connected = True
                self._notify("✅ Login com senha 2FA realizado!")
            else:
                raise Exception("Conta com verificação em duas etapas. Informe a senha.")
        except PhoneCodeInvalidError:
            raise Exception("Código inválido. Tente novamente.")

    # ── Envio de arquivo ──────────────────────────────────────
    def send_file(self, filepath, caption="", on_done=None):
        """Envia um arquivo para Mensagens Salvas (não bloqueia a UI)."""
        if not self._connected:
            self._notify("⚠️ Telegram não conectado.")
            return
        if not os.path.exists(filepath):
            self._notify(f"⚠️ Arquivo não encontrado: {filepath}")
            return
        fut = self._run_async(self._send_file(filepath, caption))
        def _cb(f):
            exc = f.exception()
            if exc:
                self._notify(f"❌ Erro ao enviar: {exc}")
            else:
                self._notify(f"✅ Enviado: {os.path.basename(filepath)}")
            if on_done:
                on_done(exc is None)
        fut.add_done_callback(_cb)

    async def _send_file(self, filepath, caption):
        client = self._get_client()
        if not client.is_connected():
            await client.connect()
        await client.send_file("me", filepath, caption=caption)

    # ── Conveniências ─────────────────────────────────────────
    def send_db_if_enabled(self):
        if self.auto_send_db and self._connected:
            now = datetime.now().strftime("%d/%m/%Y %H:%M")
            self.send_file(
                "stock_manager.db",
                caption=f"📦 Backup BD — ADR INFO\n🕐 {now}"
            )

    def send_pdf_if_enabled(self, pdf_path):
        if self.auto_send_pdf and self._connected:
            now = datetime.now().strftime("%d/%m/%Y %H:%M")
            self.send_file(
                pdf_path,
                caption=f"🧾 Nota Fiscal — ADR INFO\n🕐 {now}"
            )

    def is_connected(self):
        return self._connected


# Instância global — acessível por InvoiceGenerator e Sale
telegram_manager = TelegramManager()

# ========================= MAIN WINDOW =========================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📱 ADR INFO - Sistema de Gestão Completo 📱")
        self.setGeometry(100, 100, 1400, 900)
        
        # Inicializar carrinho de compras
        self.cart_items = []
        
        # Inicializar gerenciadores
        self.printer_manager = PrinterManager()
        self.invoice_generator = InvoiceGenerator()
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        # Tabs
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #f0f0f0;
                color: #00FF00;
                padding: 15px 25px;
                margin-right: 2px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: #4CAF50;
                color: #00FF00;
                font-weight: bold;
            }
            QTabBar::tab:hover {
                background-color: #e0e0e0;
                color: #00FF00;
            }
        """)
        layout.addWidget(self.tabs)
        
        # Criar todas as abas
        self.create_products_tab()
        self.create_sales_tab()
        self.create_services_tab()
        self.create_expenses_tab()
        self.create_cash_tab()
        self.create_telegram_tab()
        
        # Atualizar todos os dados
        self.refresh_all_data()
        
        # Aplicar tema escuro como padrão
        self.apply_dark_theme(True)
    
    # ========================= ABA TELEGRAM =========================

    def _update_telegram_status(self, msg: str):
        """Callback thread-safe para atualizar o status na UI."""
        from PyQt6.QtCore import QMetaObject, Qt
        def _update():
            self.tg_status_label.setText(msg)
            is_ok = msg.startswith("✅")
            color = "#00CC66" if is_ok else ("#FF4444" if msg.startswith("❌") else "#FFD700")
            self.tg_status_label.setStyleSheet(
                f"color:{color};padding:8px;border-radius:6px;"
                f"background-color:#1a1a2e;font-weight:bold;font-size:12px;"
            )
        QMetaObject.invokeMethod(self.tg_status_label, "repaint", Qt.ConnectionType.QueuedConnection)
        # Chamada segura para atualizar texto via thread principal
        try:
            self.tg_status_label.setText(msg)
            is_ok = msg.startswith("✅")
            color = "#00CC66" if is_ok else ("#FF4444" if msg.startswith("❌") else "#FFD700")
            self.tg_status_label.setStyleSheet(
                f"color:{color};padding:8px;border-radius:6px;"
                f"background-color:#1a1a2e;font-weight:bold;font-size:12px;"
            )
        except Exception:
            pass

    def create_telegram_tab(self):
        if not TELETHON_AVAILABLE:
            warn = QWidget()
            wl = QVBoxLayout()
            warn.setLayout(wl)
            lbl = QLabel("⚠️ Telethon não instalado.\nExecute: pip install telethon")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setFont(QFont("Arial", 14))
            lbl.setStyleSheet("color:#FF4444;")
            wl.addWidget(lbl)
            self.tabs.addTab(warn, "📱 Telegram")
            return

        from PyQt6.QtWidgets import QScrollArea
        outer = QWidget()
        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer.setLayout(outer_layout)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(scroll.Shape.NoFrame)
        outer_layout.addWidget(scroll)

        tg_widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        tg_widget.setLayout(layout)
        scroll.setWidget(tg_widget)

        FIELD_STYLE = "QLineEdit{background:#1a1a2e;color:#00D4FF;border:2px solid #00D4FF;border-radius:6px;padding:6px 10px;font-size:12px;font-weight:bold;}"

        # Título
        title = QLabel("📱 Integração Telegram")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color:#00D4FF;margin-bottom:6px;")
        layout.addWidget(title)

        # Status
        self.tg_status_label = QLabel("⏳ Verificando sessão salva...")
        self.tg_status_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.tg_status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tg_status_label.setStyleSheet(
            "color:#FFD700;padding:8px;border-radius:6px;background-color:#1a1a2e;font-weight:bold;"
        )
        self.tg_status_label.setWordWrap(True)
        layout.addWidget(self.tg_status_label)

        # Registrar callback de status no TelegramManager
        telegram_manager.set_status_callback(self._update_telegram_status)

        # ── Seção credenciais ──────────────────────────────────
        cred_title = QLabel("🔑 Credenciais da API (my.telegram.org/apps)")
        cred_title.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        cred_title.setStyleSheet("color:#888;margin-top:6px;")
        layout.addWidget(cred_title)

        form_row1 = QHBoxLayout()
        self.tg_api_id_edit = QLineEdit()
        self.tg_api_id_edit.setPlaceholderText("API ID (ex: 1234567)")
        self.tg_api_id_edit.setStyleSheet(FIELD_STYLE)
        self.tg_api_id_edit.setText(telegram_manager.api_id or "")
        form_row1.addWidget(QLabel("API ID:"))
        form_row1.addWidget(self.tg_api_id_edit)

        self.tg_api_hash_edit = QLineEdit()
        self.tg_api_hash_edit.setPlaceholderText("API Hash")
        self.tg_api_hash_edit.setStyleSheet(FIELD_STYLE)
        self.tg_api_hash_edit.setText(telegram_manager.api_hash or "")
        form_row1.addWidget(QLabel("API Hash:"))
        form_row1.addWidget(self.tg_api_hash_edit)
        layout.addLayout(form_row1)

        # Telefone + botão enviar código
        phone_row = QHBoxLayout()
        self.tg_phone_edit = QLineEdit()
        self.tg_phone_edit.setPlaceholderText("Telefone com DDI (ex: +5511999998888)")
        self.tg_phone_edit.setStyleSheet(FIELD_STYLE)
        self.tg_phone_edit.setText(telegram_manager.phone or "")
        phone_row.addWidget(self.tg_phone_edit, 3)

        send_code_btn = QPushButton("📩 Enviar Código")
        send_code_btn.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        send_code_btn.setStyleSheet(
            "QPushButton{background:#0088CC;color:white;padding:8px 14px;border:none;border-radius:6px;}"
            "QPushButton:hover{background:#006FA8;}"
        )
        send_code_btn.clicked.connect(self._tg_send_code)
        phone_row.addWidget(send_code_btn, 1)
        layout.addLayout(phone_row)

        # Código recebido + botão confirmar
        code_row = QHBoxLayout()
        self.tg_code_edit = QLineEdit()
        self.tg_code_edit.setPlaceholderText("Código recebido no Telegram")
        self.tg_code_edit.setStyleSheet(FIELD_STYLE)
        code_row.addWidget(self.tg_code_edit, 3)

        confirm_btn = QPushButton("✅ Confirmar")
        confirm_btn.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        confirm_btn.setStyleSheet(
            "QPushButton{background:#228B22;color:white;padding:8px 14px;border:none;border-radius:6px;}"
            "QPushButton:hover{background:#1a6b1a;}"
        )
        confirm_btn.clicked.connect(self._tg_confirm_code)
        code_row.addWidget(confirm_btn, 1)
        layout.addLayout(code_row)

        # Senha 2FA (opcional)
        self.tg_password_edit = QLineEdit()
        self.tg_password_edit.setPlaceholderText("Senha 2FA (apenas se sua conta tiver verificação em duas etapas)")
        self.tg_password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.tg_password_edit.setStyleSheet(FIELD_STYLE)
        layout.addWidget(self.tg_password_edit)

        # ── Seção de automação ─────────────────────────────────
        auto_title = QLabel("⚙️ Envio Automático")
        auto_title.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        auto_title.setStyleSheet("color:#888;margin-top:8px;")
        layout.addWidget(auto_title)

        self.tg_auto_db_chk = QCheckBox("📦 Enviar banco de dados automaticamente após cada venda")
        self.tg_auto_db_chk.setFont(QFont("Arial", 11))
        self.tg_auto_db_chk.setStyleSheet("color:#AAD4FF;")
        self.tg_auto_db_chk.setChecked(telegram_manager.auto_send_db)
        self.tg_auto_db_chk.stateChanged.connect(self._tg_toggle_auto_db)
        layout.addWidget(self.tg_auto_db_chk)

        self.tg_auto_pdf_chk = QCheckBox("🧾 Enviar PDF de nota fiscal automaticamente após cada venda")
        self.tg_auto_pdf_chk.setFont(QFont("Arial", 11))
        self.tg_auto_pdf_chk.setStyleSheet("color:#AAD4FF;")
        self.tg_auto_pdf_chk.setChecked(telegram_manager.auto_send_pdf)
        self.tg_auto_pdf_chk.stateChanged.connect(self._tg_toggle_auto_pdf)
        layout.addWidget(self.tg_auto_pdf_chk)

        # Botão envio manual do BD
        send_db_btn = QPushButton("📤 Enviar Banco de Dados Agora")
        send_db_btn.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        send_db_btn.setStyleSheet(
            "QPushButton{background:#B8860B;color:white;padding:10px;border:none;border-radius:7px;margin-top:6px;}"
            "QPushButton:hover{background:#996000;}"
        )
        send_db_btn.clicked.connect(self._tg_send_db_now)
        layout.addWidget(send_db_btn)

        layout.addStretch()
        self.tabs.addTab(outer, "📱 Telegram")

        # Tentar reconectar automaticamente com sessão salva
        telegram_manager.try_auto_connect(on_done=lambda ok: self._update_telegram_status(
            "✅ Conectado ao Telegram!" if ok else "🔑 Faça login para conectar."
        ))

    def _tg_send_code(self):
        api_id  = self.tg_api_id_edit.text().strip()
        api_hash = self.tg_api_hash_edit.text().strip()
        phone   = self.tg_phone_edit.text().strip()
        if not api_id or not api_hash or not phone:
            self._update_telegram_status("⚠️ Preencha API ID, API Hash e Telefone.")
            return
        self._update_telegram_status("⏳ Enviando código...")
        def on_done(ok, err):
            if ok:
                self._update_telegram_status("📩 Código enviado! Digite abaixo.")
            else:
                self._update_telegram_status(f"❌ {err}")
        telegram_manager.send_code(api_id, api_hash, phone, on_done=on_done)

    def _tg_confirm_code(self):
        code = self.tg_code_edit.text().strip()
        password = self.tg_password_edit.text().strip() or None
        if not code:
            self._update_telegram_status("⚠️ Digite o código recebido.")
            return
        self._update_telegram_status("⏳ Verificando código...")
        def on_done(ok, err):
            if ok:
                self._update_telegram_status("✅ Login realizado com sucesso!")
            else:
                self._update_telegram_status(f"❌ {err}")
        telegram_manager.confirm_code(code, password=password, on_done=on_done)

    def _tg_toggle_auto_db(self, state):
        telegram_manager.auto_send_db = bool(state)
        telegram_manager.save_config()

    def _tg_toggle_auto_pdf(self, state):
        telegram_manager.auto_send_pdf = bool(state)
        telegram_manager.save_config()

    def _tg_send_db_now(self):
        if not telegram_manager.is_connected():
            self._update_telegram_status("⚠️ Faça login primeiro.")
            return
        self._update_telegram_status("⏳ Enviando banco de dados...")
        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        telegram_manager.send_file(
            "stock_manager.db",
            caption=f"📦 Backup BD — ADR INFO\n🕐 {now}"
        )

    def create_products_tab(self):
        # Aba de produtos
        products_widget = QWidget()
        layout = QVBoxLayout()
        products_widget.setLayout(layout)
        

        
        # Botões
        buttons_layout = QHBoxLayout()
        add_btn = QPushButton("➕ Adicionar Produto")
        edit_btn = QPushButton("✏️ Editar Produto")
        delete_btn = QPushButton("🗑️ Excluir Produto")
        refresh_btn = QPushButton("🔄 Atualizar")
        
        # Estilizar botões
        button_style = """
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 6px;
                font-weight: bold;
                font-size: 12px;
                margin: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """
        
        add_btn.setStyleSheet(button_style)
        edit_btn.setStyleSheet(button_style.replace("#4CAF50", "#2196F3").replace("#45a049", "#1976D2"))
        delete_btn.setStyleSheet(button_style.replace("#4CAF50", "#f44336").replace("#45a049", "#d32f2f"))
        refresh_btn.setStyleSheet(button_style.replace("#4CAF50", "#FF9800").replace("#45a049", "#F57C00"))
        
        add_btn.clicked.connect(self.add_product)
        edit_btn.clicked.connect(self.edit_product)
        delete_btn.clicked.connect(self.delete_product)
        refresh_btn.clicked.connect(self.refresh_products)
        
        buttons_layout.addWidget(add_btn)
        buttons_layout.addWidget(edit_btn)
        buttons_layout.addWidget(delete_btn)
        buttons_layout.addWidget(refresh_btn)
        buttons_layout.addStretch()
        
        layout.addLayout(buttons_layout)
        
        # Campo de busca
        search_layout = QHBoxLayout()
        search_label = QLabel("🔍 Buscar Produto:")
        search_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        search_label.setStyleSheet("color: #2E8B57; margin: 5px;")
        
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Digite o nome do produto para filtrar...")
        self.search_edit.setStyleSheet("""
            QLineEdit {
                padding: 8px 12px;
                border: 2px solid #4CAF50;
                border-radius: 6px;
                font-size: 12px;
                background-color: white;
                color: #333;
            }
            QLineEdit:focus {
                border-color: #2E8B57;
                background-color: #F0FFF0;
            }
        """)
        self.search_edit.textChanged.connect(self.filter_products)
        
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_edit)
        search_layout.addStretch()
        
        layout.addLayout(search_layout)
        
        # Tabela de produtos
        self.products_table = QTableWidget()
        self.products_table.setColumnCount(5)
        self.products_table.setHorizontalHeaderLabels(["🆔 ID", "📝 Nome", "💰 Preço Compra", "💵 Preço Venda", "📊 Quantidade"])
        self.products_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        # Estilizar tabela
        self.products_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #ddd;
                background-color: white;
                alternate-background-color: #f9f9f9;
                font-size: 11px;
                color: #00FF00;
            }
            QHeaderView::section {
                background-color: #4CAF50;
                color: white;
                padding: 12px;
                border: none;
                font-weight: bold;
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #eee;
                color: #00FF00;
            }
            QTableWidget::item:selected {
                background-color: #E8F5E8;
                color: #00FF00;
            }
        """)
        self.products_table.setAlternatingRowColors(True)
        self.products_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        
        layout.addWidget(self.products_table)
        
        self.tabs.addTab(products_widget, "📦 Produtos")
    
    def create_sales_tab(self):
        # Aba de vendas
        sales_widget = QWidget()
        main_layout = QHBoxLayout()
        sales_widget.setLayout(main_layout)
        
        # Lado esquerdo - Produtos disponíveis
        left_layout = QVBoxLayout()
        left_widget = QWidget()
        left_widget.setLayout(left_layout)
        left_widget.setMaximumWidth(600)
        
        # Título produtos disponíveis
        products_title = QLabel("🛍️ Produtos Disponíveis para Venda")
        products_title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        products_title.setStyleSheet("color: #2E8B57; margin: 10px; padding: 10px; background-color: #F0FFF0; border-radius: 5px;")
        left_layout.addWidget(products_title)
        
        # Campo de busca para produtos disponíveis
        search_layout = QHBoxLayout()
        search_label = QLabel("🔍 Buscar Produto:")
        search_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        search_label.setStyleSheet("color: #2E8B57; margin: 5px;")
        
        self.sales_search_edit = QLineEdit()
        self.sales_search_edit.setPlaceholderText("Digite o nome do produto para filtrar...")
        self.sales_search_edit.setStyleSheet("""
            QLineEdit {
                padding: 8px 12px;
                border: 2px solid #2196F3;
                border-radius: 6px;
                font-size: 12px;
                background-color: white;
                color: #333;
            }
            QLineEdit:focus {
                border-color: #1976D2;
                background-color: #F0F8FF;
            }
        """)
        self.sales_search_edit.textChanged.connect(self.filter_available_products)
        
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.sales_search_edit)
        search_layout.addStretch()
        
        left_layout.addLayout(search_layout)
        
        # Tabela de produtos disponíveis para venda
        self.available_products_table = QTableWidget()
        self.available_products_table.setColumnCount(4)
        self.available_products_table.setHorizontalHeaderLabels(["📝 Nome", "💵 Preço", "📊 Estoque", "🛒 Ação"])
        self.available_products_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.available_products_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #ddd;
                background-color: white;
                alternate-background-color: #f9f9f9;
                font-size: 11px;
                color: #00FF00;
            }
            QHeaderView::section {
                background-color: #2196F3;
                color: white;
                padding: 10px;
                border: none;
                font-weight: bold;
                font-size: 11px;
            }
            QTableWidget::item {
                color: #00FF00;
            }
        """)
        self.available_products_table.setAlternatingRowColors(True)
        left_layout.addWidget(self.available_products_table)
        
        main_layout.addWidget(left_widget)
        
        # Lado direito - Carrinho e histórico
        right_layout = QVBoxLayout()
        right_widget = QWidget()
        right_widget.setLayout(right_layout)
        
        # Seção do carrinho
        cart_title = QLabel("🛒 Carrinho de Compras")
        cart_title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        cart_title.setStyleSheet("color: #FF6B35; margin: 10px; padding: 10px; background-color: #FFF4F0; border-radius: 5px;")
        right_layout.addWidget(cart_title)
        
        # Tabela do carrinho
        self.cart_table = QTableWidget()
        self.cart_table.setColumnCount(5)
        self.cart_table.setHorizontalHeaderLabels(["📦 Produto", "💰 Preço Unit.", "🔢 Qtd", "💵 Subtotal", "❌ Remover"])
        self.cart_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.cart_table.setMaximumHeight(250)
        self.cart_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #ddd;
                background-color: white;
                alternate-background-color: #fff8f0;
                font-size: 10px;
                color: #00FF00;
            }
            QHeaderView::section {
                background-color: #FF6B35;
                color: white;
                padding: 8px;
                border: none;
                font-weight: bold;
                font-size: 10px;
            }
            QTableWidget::item {
                color: #00FF00;
            }
        """)
        self.cart_table.setAlternatingRowColors(True)
        right_layout.addWidget(self.cart_table)
        
        # Total do carrinho
        self.cart_total_label = QLabel("💰 Total: R$ 0,00")
        self.cart_total_label.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        self.cart_total_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.cart_total_label.setStyleSheet("color: #228B22; padding: 10px; border: 2px solid #228B22; border-radius: 5px; background-color: #F0FFF0; margin: 10px;")
        right_layout.addWidget(self.cart_total_label)
        
        # Botões do carrinho
        cart_buttons_layout = QHBoxLayout()
        clear_cart_btn = QPushButton("🗑️ Limpar Carrinho")
        finalize_sale_btn = QPushButton("✅ Finalizar Venda")
        
        clear_cart_btn.setStyleSheet("""
            QPushButton {
                background-color: #DC143C;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 6px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #B22222;
            }
        """)
        
        finalize_sale_btn.setStyleSheet("""
            QPushButton {
                background-color: #228B22;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 6px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #32CD32;
            }
        """)
        
        clear_cart_btn.clicked.connect(self.clear_cart)
        finalize_sale_btn.clicked.connect(self.finalize_sale)
        
        cart_buttons_layout.addWidget(clear_cart_btn)
        cart_buttons_layout.addWidget(finalize_sale_btn)
        right_layout.addLayout(cart_buttons_layout)
        
        # Separador
        separator = QLabel()
        separator.setMinimumHeight(20)
        right_layout.addWidget(separator)
        
        # Histórico de vendas
        history_title = QLabel("📊 Histórico de Vendas")
        history_title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        history_title.setStyleSheet("color: #4169E1; margin: 10px; padding: 10px; background-color: #F0F8FF; border-radius: 5px;")
        right_layout.addWidget(history_title)
        
        # Botão atualizar histórico
        refresh_btn = QPushButton("🔄 Atualizar Histórico")
        refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0000CD;
            }
        """)
        refresh_btn.clicked.connect(self.refresh_sales)
        right_layout.addWidget(refresh_btn)
        
        # Tabela de vendas (histórico)
        self.sales_table = QTableWidget()
        self.sales_table.setColumnCount(6)
        self.sales_table.setHorizontalHeaderLabels(["🆔 ID", "📦 Produto", "🔢 Quantidade", "💰 Total", "🏦 Pagamento", "📅 Data"])
        self.sales_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.sales_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #ddd;
                background-color: white;
                alternate-background-color: #f0f8ff;
                font-size: 10px;
                color: #00FF00;
            }
            QHeaderView::section {
                background-color: #4169E1;
                color: white;
                padding: 8px;
                border: none;
                font-weight: bold;
                font-size: 10px;
            }
            QTableWidget::item {
                color: #00FF00;
            }
        """)
        self.sales_table.setAlternatingRowColors(True)
        right_layout.addWidget(self.sales_table)
        
        main_layout.addWidget(right_widget)
        
        self.tabs.addTab(sales_widget, "🛍️ Vendas")
    
    def create_services_tab(self):
        # Aba de serviços
        services_widget = QWidget()
        layout = QVBoxLayout()
        services_widget.setLayout(layout)
        

        
        # Botões
        buttons_layout = QHBoxLayout()
        add_btn = QPushButton("➕ Adicionar Serviço")
        edit_btn = QPushButton("✏️ Editar Serviço")
        delete_btn = QPushButton("🗑️ Excluir Serviço")
        refresh_btn = QPushButton("🔄 Atualizar")
        
        # Estilizar botões com cores específicas para serviços
        add_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; } QPushButton:hover { background-color: #45a049; }")
        edit_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; } QPushButton:hover { background-color: #1976D2; }")
        delete_btn.setStyleSheet("QPushButton { background-color: #f44336; color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; } QPushButton:hover { background-color: #d32f2f; }")
        refresh_btn.setStyleSheet("QPushButton { background-color: #FF9800; color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; } QPushButton:hover { background-color: #F57C00; }")
        
        add_btn.clicked.connect(self.add_service)
        edit_btn.clicked.connect(self.edit_service)
        delete_btn.clicked.connect(self.delete_service)
        refresh_btn.clicked.connect(self.refresh_services)
        
        buttons_layout.addWidget(add_btn)
        buttons_layout.addWidget(edit_btn)
        buttons_layout.addWidget(delete_btn)
        buttons_layout.addWidget(refresh_btn)
        buttons_layout.addStretch()
        
        layout.addLayout(buttons_layout)
        
        # Seção de resumo
        summary_layout = QHBoxLayout()
        
        # Total de serviços do mês
        self.monthly_services_label = QLabel("📅 Serviços do Mês: R$ 0,00")
        self.monthly_services_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.monthly_services_label.setStyleSheet("color: #2E8B57; padding: 10px; border: 2px solid #2E8B57; border-radius: 5px; background-color: #F0FFF0; margin: 5px;")
        summary_layout.addWidget(self.monthly_services_label)
        
        # Total de serviços geral
        self.total_services_label = QLabel("💰 Total Geral: R$ 0,00")
        self.total_services_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.total_services_label.setStyleSheet("color: #228B22; padding: 10px; border: 2px solid #228B22; border-radius: 5px; background-color: #F5FFF5; margin: 5px;")
        summary_layout.addWidget(self.total_services_label)
        
        layout.addLayout(summary_layout)
        
        # Tabela de serviços
        self.services_table = QTableWidget()
        self.services_table.setColumnCount(6)
        self.services_table.setHorizontalHeaderLabels(["🆔 ID", "📝 Descrição", "💰 Valor", "🏷️ Categoria", "👤 Cliente", "📅 Data"])
        self.services_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.services_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #ddd;
                background-color: white;
                alternate-background-color: #f9f9f9;
                font-size: 11px;
                color: #00FF00;
            }
            QHeaderView::section {
                background-color: #2E8B57;
                color: white;
                padding: 10px;
                border: none;
                font-weight: bold;
                font-size: 11px;
            }
            QTableWidget::item {
                color: #00FF00;
            }
        """)
        self.services_table.setAlternatingRowColors(True)
        layout.addWidget(self.services_table)
        
        self.tabs.addTab(services_widget, "🔧 Serviços")
    
    def create_expenses_tab(self):
        # Aba de gastos
        expenses_widget = QWidget()
        layout = QVBoxLayout()
        expenses_widget.setLayout(layout)
        

        
        # Botões
        buttons_layout = QHBoxLayout()
        add_btn = QPushButton("➕ Adicionar Gasto")
        edit_btn = QPushButton("✏️ Editar Gasto")
        delete_btn = QPushButton("🗑️ Excluir Gasto")
        refresh_btn = QPushButton("🔄 Atualizar")
        
        # Estilizar botões com cores para gastos
        add_btn.setStyleSheet("QPushButton { background-color: #DC143C; color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; } QPushButton:hover { background-color: #B22222; }")
        edit_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; } QPushButton:hover { background-color: #1976D2; }")
        delete_btn.setStyleSheet("QPushButton { background-color: #8B0000; color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; } QPushButton:hover { background-color: #660000; }")
        refresh_btn.setStyleSheet("QPushButton { background-color: #FF9800; color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; } QPushButton:hover { background-color: #F57C00; }")
        
        add_btn.clicked.connect(self.add_expense)
        edit_btn.clicked.connect(self.edit_expense)
        delete_btn.clicked.connect(self.delete_expense)
        refresh_btn.clicked.connect(self.refresh_expenses)
        
        buttons_layout.addWidget(add_btn)
        buttons_layout.addWidget(edit_btn)
        buttons_layout.addWidget(delete_btn)
        buttons_layout.addWidget(refresh_btn)
        buttons_layout.addStretch()
        
        layout.addLayout(buttons_layout)
        
        # Seção de resumo
        summary_layout = QHBoxLayout()
        
        # Total de gastos do mês
        self.monthly_expenses_label = QLabel("📅 Gastos do Mês: R$ 0,00")
        self.monthly_expenses_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.monthly_expenses_label.setStyleSheet("color: #DC143C; padding: 10px; border: 2px solid #DC143C; border-radius: 5px; background-color: #FFE4E1; margin: 5px;")
        summary_layout.addWidget(self.monthly_expenses_label)
        
        # Total de gastos geral
        self.total_expenses_label = QLabel("💰 Total Geral: R$ 0,00")
        self.total_expenses_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.total_expenses_label.setStyleSheet("color: #8B0000; padding: 10px; border: 2px solid #8B0000; border-radius: 5px; background-color: #FFF0F0; margin: 5px;")
        summary_layout.addWidget(self.total_expenses_label)
        
        layout.addLayout(summary_layout)
        
        # Tabela de gastos
        self.expenses_table = QTableWidget()
        self.expenses_table.setColumnCount(5)
        self.expenses_table.setHorizontalHeaderLabels(["🆔 ID", "📝 Descrição", "💰 Valor", "🏷️ Categoria", "📅 Data"])
        self.expenses_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.expenses_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #ddd;
                background-color: white;
                alternate-background-color: #ffe4e1;
                font-size: 11px;
                color: #00FF00;
            }
            QHeaderView::section {
                background-color: #DC143C;
                color: white;
                padding: 10px;
                border: none;
                font-weight: bold;
                font-size: 11px;
            }
            QTableWidget::item {
                color: #00FF00;
            }
        """)
        self.expenses_table.setAlternatingRowColors(True)
        layout.addWidget(self.expenses_table)
        
        self.tabs.addTab(expenses_widget, "💸 Gastos")
    
    def create_cash_tab(self):
        # Aba de caixa — com scroll para telas menores
        from PyQt6.QtWidgets import QScrollArea
        outer_widget = QWidget()
        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_widget.setLayout(outer_layout)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(scroll.Shape.NoFrame)
        outer_layout.addWidget(scroll)
        cash_widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(15, 15, 15, 15)
        cash_widget.setLayout(layout)
        scroll.setWidget(cash_widget)

        # Título
        title_label = QLabel("💰 Resumo Financeiro")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #FFD700; margin-bottom: 2px;")
        layout.addWidget(title_label)

        # ── 3 cards lado a lado ──────────────────────────
        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        # Card: Total em Caixa
        self.cash_label = QLabel("🏦 Total em Caixa\nR$ 0,00")
        self.cash_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.cash_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cash_label.setStyleSheet("color:#4169E1;padding:10px 6px;border:2px solid #4169E1;border-radius:10px;background-color:#0a1628;")
        self.cash_label.setMinimumHeight(65)
        top_row.addWidget(self.cash_label)

        # Card: Total da Semana
        self.week_label = QLabel("📅 Semana (Seg–Sáb)\nR$ 0,00")
        self.week_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.week_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.week_label.setStyleSheet("color:#228B22;padding:10px 6px;border:2px solid #228B22;border-radius:10px;background-color:#0a1e0a;")
        self.week_label.setMinimumHeight(65)
        top_row.addWidget(self.week_label)

        # Card: Total do Mês
        self.month_label = QLabel("🗓️ Total do Mês\nR$ 0,00")
        self.month_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.month_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.month_label.setStyleSheet("color:#B8860B;padding:10px 6px;border:2px solid #B8860B;border-radius:10px;background-color:#1a1500;")
        self.month_label.setMinimumHeight(65)
        top_row.addWidget(self.month_label)

        layout.addLayout(top_row)

        # ── Seção de Lucros ───────────────────────────────────────────────
        profit_title = QLabel("📊 Lucro Líquido por Período")
        profit_title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        profit_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        profit_title.setStyleSheet("color:#2E8B57;padding:5px;background-color:#0a1e0a;border-radius:6px;margin-top:4px;")
        layout.addWidget(profit_title)

        # Linha com os dois blocos de seleção lado a lado
        selectors_row = QHBoxLayout()
        selectors_row.setSpacing(10)

        # ── Bloco DIA ────────────────────────────────────────
        day_block = QWidget()
        day_block_layout = QVBoxLayout()
        day_block_layout.setSpacing(4)
        day_block.setLayout(day_block_layout)

        day_block_title = QLabel("📅 Vendas do Dia")
        day_block_title.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        day_block_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        day_block_title.setStyleSheet("color: #00CED1;")
        day_block_layout.addWidget(day_block_title)

        # Seletor de data
        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setStyleSheet("QDateEdit{padding:5px 8px;border:2px solid #00CED1;border-radius:5px;font-size:11px;font-weight:bold;background-color:#0a1e1e;color:#00CED1;}")
        self.date_edit.dateChanged.connect(self.update_day_profit)
        day_block_layout.addWidget(self.date_edit)

        # Card resultado dia
        self.daily_profit_label = QLabel("💸 Vendas do Dia\nR$ 0,00")
        self.daily_profit_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.daily_profit_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.daily_profit_label.setStyleSheet("color:#00CED1;padding:12px 8px;border:2px solid #00CED1;border-radius:10px;background-color:#051515;")
        self.daily_profit_label.setMinimumHeight(75)
        day_block_layout.addWidget(self.daily_profit_label)

        selectors_row.addWidget(day_block)

        # ── Bloco MÊS ────────────────────────────────────────
        month_block = QWidget()
        month_block_layout = QVBoxLayout()
        month_block_layout.setSpacing(4)
        month_block.setLayout(month_block_layout)

        month_block_title = QLabel("🗓️ Vendas do Mês")
        month_block_title.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        month_block_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        month_block_title.setStyleSheet("color: #FF8C00;")
        month_block_layout.addWidget(month_block_title)

        # Seletores mês + ano
        month_sel_row = QHBoxLayout()
        self.month_combo = QComboBox()
        self.month_combo.addItems([
            "01 - Janeiro", "02 - Fevereiro", "03 - Março",
            "04 - Abril",   "05 - Maio",      "06 - Junho",
            "07 - Julho",   "08 - Agosto",    "09 - Setembro",
            "10 - Outubro", "11 - Novembro",  "12 - Dezembro"
        ])
        self.month_combo.setCurrentIndex(QDate.currentDate().month() - 1)
        self.month_combo.setStyleSheet("QComboBox{padding:5px 8px;border:2px solid #FF8C00;border-radius:5px;font-size:11px;font-weight:bold;background-color:#1a0d00;color:#FF8C00;}")
        self.month_combo.currentIndexChanged.connect(self.update_month_profit)

        self.year_spin = QSpinBox()
        self.year_spin.setRange(2020, 2099)
        self.year_spin.setValue(QDate.currentDate().year())
        self.year_spin.setStyleSheet("QSpinBox{padding:5px 8px;border:2px solid #FF8C00;border-radius:5px;font-size:11px;font-weight:bold;background-color:#1a0d00;color:#FF8C00;min-width:65px;}")
        self.year_spin.valueChanged.connect(self.update_month_profit)

        month_sel_row.addWidget(self.month_combo)
        month_sel_row.addWidget(self.year_spin)
        month_block_layout.addLayout(month_sel_row)

        # Card resultado mês
        self.monthly_profit_label = QLabel("📈 Vendas do Mês\nR$ 0,00")
        self.monthly_profit_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.monthly_profit_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.monthly_profit_label.setStyleSheet("color:#FF8C00;padding:12px 8px;border:2px solid #FF8C00;border-radius:10px;background-color:#1a0800;")
        self.monthly_profit_label.setMinimumHeight(75)
        month_block_layout.addWidget(self.monthly_profit_label)

        selectors_row.addWidget(month_block)
        layout.addLayout(selectors_row)

        # Botão atualizar
        refresh_btn = QPushButton("🔄 Atualizar Valores")
        refresh_btn.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        refresh_btn.setStyleSheet("QPushButton{background-color:#4CAF50;color:white;padding:9px;border:none;border-radius:7px;margin-top:4px;}QPushButton:hover{background-color:#45a049;}")
        refresh_btn.clicked.connect(self.refresh_cash)
        layout.addWidget(refresh_btn)

        layout.addStretch()

        self.tabs.addTab(outer_widget, "💰 Caixa")
    

    
    # ========================= MÉTODOS DE PRODUTOS =========================
    def add_product(self):
        dialog = ProductDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_product_data()
            if data['name']:
                product = Product(data['name'], data['purchase_price'], 
                                data['sale_price'], data['quantity'])
                product.save()
                self.refresh_all_data()
                QMessageBox.information(self, "✅ Sucesso", "Produto adicionado com sucesso!")
            else:
                QMessageBox.warning(self, "⚠️ Erro", "Nome do produto é obrigatório!")
    
    def edit_product(self):
        current_row = self.products_table.currentRow()
        if current_row >= 0:
            product_id = int(self.products_table.item(current_row, 0).text())
            product = Product.get_by_id(product_id)
            if product:
                dialog = ProductDialog(product, parent=self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    data = dialog.get_product_data()
                    if data['name']:
                        product.name = data['name']
                        product.purchase_price = data['purchase_price']
                        product.sale_price = data['sale_price']
                        product.quantity = data['quantity']
                        product.save()
                        self.refresh_all_data()
                        QMessageBox.information(self, "✅ Sucesso", "Produto editado com sucesso!")
                    else:
                        QMessageBox.warning(self, "⚠️ Erro", "Nome do produto é obrigatório!")
        else:
            QMessageBox.warning(self, "⚠️ Erro", "Selecione um produto para editar!")
    
    def delete_product(self):
        current_row = self.products_table.currentRow()
        if current_row >= 0:
            product_id = int(self.products_table.item(current_row, 0).text())
            reply = QMessageBox.question(self, "❓ Confirmar", "Tem certeza que deseja excluir este produto?")
            if reply == QMessageBox.StandardButton.Yes:
                Product.delete(product_id)
                self.refresh_all_data()
                QMessageBox.information(self, "✅ Sucesso", "Produto excluído com sucesso!")
        else:
            QMessageBox.warning(self, "⚠️ Erro", "Selecione um produto para excluir!")
    
    def refresh_products(self):
        products = Product.get_all()
        self.products_table.setRowCount(len(products))
        
        for row, product in enumerate(products):
            self.products_table.setItem(row, 0, QTableWidgetItem(str(product.id)))
            self.products_table.setItem(row, 1, QTableWidgetItem(product.name))
            self.products_table.setItem(row, 2, QTableWidgetItem(f"R$ {product.purchase_price:.2f}"))
            self.products_table.setItem(row, 3, QTableWidgetItem(f"R$ {product.sale_price:.2f}"))
            self.products_table.setItem(row, 4, QTableWidgetItem(str(product.quantity)))
    
    def filter_products(self):
        """Filtra os produtos na tabela baseado no texto de busca"""
        search_text = self.search_edit.text().lower().strip()
        
        # Se o campo de busca estiver vazio, mostrar todos os produtos
        if not search_text:
            self.refresh_products()
            return
        
        # Obter todos os produtos
        all_products = Product.get_all()
        
        # Filtrar produtos que contêm o texto de busca no nome
        filtered_products = [
            product for product in all_products 
            if search_text in product.name.lower()
        ]
        
        # Atualizar a tabela apenas com os produtos filtrados
        self.products_table.setRowCount(len(filtered_products))
        
        for row, product in enumerate(filtered_products):
            self.products_table.setItem(row, 0, QTableWidgetItem(str(product.id)))
            self.products_table.setItem(row, 1, QTableWidgetItem(product.name))
            self.products_table.setItem(row, 2, QTableWidgetItem(f"R$ {product.purchase_price:.2f}"))
            self.products_table.setItem(row, 3, QTableWidgetItem(f"R$ {product.sale_price:.2f}"))
            self.products_table.setItem(row, 4, QTableWidgetItem(str(product.quantity)))

    # ========================= MÉTODOS DE VENDAS =========================
    def refresh_available_products(self):
        """Atualiza a tabela de produtos disponíveis na aba de vendas"""
        products = Product.get_all()
        # Filtrar apenas produtos com estoque
        available_products = [p for p in products if p.quantity > 0]
        
        self.available_products_table.setRowCount(len(available_products))
        
        for row, product in enumerate(available_products):
            self.available_products_table.setItem(row, 0, QTableWidgetItem(product.name))
            self.available_products_table.setItem(row, 1, QTableWidgetItem(f"R$ {product.sale_price:.2f}"))
            self.available_products_table.setItem(row, 2, QTableWidgetItem(str(product.quantity)))
            
            # Botão "Adicionar ao Carrinho"
            add_btn = QPushButton("🛒 Adicionar")
            add_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    padding: 5px 10px;
                    border: none;
                    border-radius: 3px;
                    font-weight: bold;
                    font-size: 10px;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
            """)
            add_btn.clicked.connect(self.create_add_to_cart_handler(product))
            self.available_products_table.setCellWidget(row, 3, add_btn)
    
    def filter_available_products(self):
        """Filtra os produtos disponíveis na tabela baseado no texto de busca"""
        search_text = self.sales_search_edit.text().lower().strip()
        
        # Obter todos os produtos com estoque
        products = Product.get_all()
        available_products = [p for p in products if p.quantity > 0]
        
        # Se o campo de busca estiver vazio, mostrar todos os produtos disponíveis
        if not search_text:
            filtered_products = available_products
        else:
            # Filtrar produtos que contêm o texto de busca no nome
            filtered_products = [
                product for product in available_products 
                if search_text in product.name.lower()
            ]
        
        # Atualizar a tabela apenas com os produtos filtrados
        self.available_products_table.setRowCount(len(filtered_products))
        
        for row, product in enumerate(filtered_products):
            self.available_products_table.setItem(row, 0, QTableWidgetItem(product.name))
            self.available_products_table.setItem(row, 1, QTableWidgetItem(f"R$ {product.sale_price:.2f}"))
            self.available_products_table.setItem(row, 2, QTableWidgetItem(str(product.quantity)))
            
            # Botão "Adicionar ao Carrinho"
            add_btn = QPushButton("🛒 Adicionar")
            add_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    padding: 5px 10px;
                    border: none;
                    border-radius: 3px;
                    font-weight: bold;
                    font-size: 10px;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
            """)
            add_btn.clicked.connect(self.create_add_to_cart_handler(product))
            self.available_products_table.setCellWidget(row, 3, add_btn)
    
    def create_add_to_cart_handler(self, product):
        """Cria um handler para adicionar produto ao carrinho"""
        return lambda: self.add_to_cart(product)
    
    def create_remove_from_cart_handler(self, cart_item):
        """Cria um handler para remover produto do carrinho"""
        return lambda: self.remove_from_cart(cart_item)
    
    def add_to_cart(self, product):
        """Adiciona um produto ao carrinho"""
        if product.quantity <= 0:
            QMessageBox.warning(self, "⚠️ Erro", "Produto sem estoque!")
            return
        
        # Verificar se o produto já está no carrinho
        for cart_item in self.cart_items:
            if hasattr(cart_item.product, 'id') and cart_item.product.id == product.id:
                # Produto já no carrinho, perguntar se quer aumentar quantidade
                dialog = AddToCartDialog(product, parent=self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    quantity = dialog.get_quantity()
                    if cart_item.quantity + quantity <= product.quantity:
                        cart_item.update_quantity(cart_item.quantity + quantity)
                        self.refresh_cart_display()
                        QMessageBox.information(self, "✅ Sucesso", f"Quantidade atualizada no carrinho!")
                    else:
                        QMessageBox.warning(self, "⚠️ Erro", "Quantidade excede o estoque disponível!")
                return
        
        # Produto não está no carrinho, adicionar novo
        dialog = AddToCartDialog(product, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            quantity = dialog.get_quantity()
            cart_item = CartItem(product, quantity)
            self.cart_items.append(cart_item)
            self.refresh_cart_display()
            QMessageBox.information(self, "✅ Sucesso", f"{product.name} adicionado ao carrinho!")
    
    def refresh_cart_display(self):
        """Atualiza a exibição do carrinho"""
        self.cart_table.setRowCount(len(self.cart_items))
        total = 0
        
        for row, cart_item in enumerate(self.cart_items):
            self.cart_table.setItem(row, 0, QTableWidgetItem(cart_item.product.name))
            self.cart_table.setItem(row, 1, QTableWidgetItem(f"R$ {cart_item.product.sale_price:.2f}"))
            self.cart_table.setItem(row, 2, QTableWidgetItem(str(cart_item.quantity)))
            self.cart_table.setItem(row, 3, QTableWidgetItem(f"R$ {cart_item.subtotal:.2f}"))
            
            # Botão remover
            remove_btn = QPushButton("❌")
            remove_btn.setStyleSheet("""
                QPushButton {
                    background-color: #DC143C;
                    color: white;
                    padding: 3px 8px;
                    border: none;
                    border-radius: 3px;
                    font-weight: bold;
                    font-size: 10px;
                }
                QPushButton:hover {
                    background-color: #B22222;
                }
            """)
            remove_btn.clicked.connect(self.create_remove_from_cart_handler(cart_item))
            self.cart_table.setCellWidget(row, 4, remove_btn)
            
            total += cart_item.subtotal
        
        self.cart_total_label.setText(f"💰 Total: R$ {total:.2f}")
    
    def remove_from_cart(self, cart_item):
        """Remove um item do carrinho"""
        if cart_item in self.cart_items:
            self.cart_items.remove(cart_item)
            self.refresh_cart_display()
            QMessageBox.information(self, "✅ Sucesso", f"{cart_item.product.name} removido do carrinho!")
    
    def clear_cart(self):
        """Limpa todo o carrinho"""
        if self.cart_items:
            reply = QMessageBox.question(self, "❓ Confirmar", "Tem certeza que deseja limpar o carrinho?")
            if reply == QMessageBox.StandardButton.Yes:
                self.cart_items.clear()
                self.refresh_cart_display()
                QMessageBox.information(self, "✅ Sucesso", "Carrinho limpo!")
        else:
            QMessageBox.information(self, "ℹ️ Info", "O carrinho já está vazio!")
    
    def finalize_sale(self):
        """Finaliza a venda com todos os itens do carrinho"""
        if not self.cart_items:
            QMessageBox.warning(self, "⚠️ Erro", "O carrinho está vazio!")
            return
        
        # Verificar se todos os produtos ainda têm estoque suficiente
        for cart_item in self.cart_items:
            if hasattr(cart_item, 'item_type') and cart_item.item_type == "service":
                continue  # Pular validação para serviços
            
            if hasattr(cart_item.product, 'id') and str(cart_item.product.id).startswith('service_'):
                continue  # Pular validação para serviços
            
            current_product = Product.get_by_id(cart_item.product.id)
            if current_product and current_product.quantity < cart_item.quantity:
                QMessageBox.warning(self, "⚠️ Erro", 
                                  f"Estoque insuficiente para {cart_item.product.name}!\n"
                                  f"Disponível: {current_product.quantity}, Solicitado: {cart_item.quantity}")
                return
        
        # Calcular total da venda
        total = sum(item.subtotal for item in self.cart_items)
        
        # Abrir diálogo de pagamento
        payment_dialog = PaymentDialog(total, parent=self)
        if payment_dialog.exec() == QDialog.DialogCode.Accepted:
            payment_data = payment_dialog.get_payment_data()
            
            try:
                # Processar venda com informações de pagamento
                sales = Sale.process_cart_sale(self.cart_items, payment_data)
                
                # Gerar nota fiscal consolidada
                invoice_path = self.invoice_generator.generate_cart_invoice(self.cart_items, sales, payment_data)
                
                # Tentar imprimir automaticamente se configurado
                print_status = ""
                if invoice_path and self.printer_manager.should_auto_print():
                    try:
                        if self.printer_manager.print_pdf(invoice_path):
                            print_status = "\n✅ Nota fiscal enviada para impressão!"
                        else:
                            print_status = "\n⚠️ Falha ao imprimir nota fiscal automaticamente."
                    except Exception as e:
                        print_status = f"\n⚠️ Erro na impressão automática: {str(e)}"
                elif self.printer_manager.selected_printer:
                    print_status = "\n💡 Impressão automática desabilitada. Vá em Configurações para habilitar."
                else:
                    print_status = "\n💡 Configure uma impressora em Configurações para impressão automática."
                
                # Limpar carrinho
                self.cart_items.clear()
                
                # Atualizar todas as telas
                self.refresh_all_data()
                
                # Mensagem de sucesso com informações de pagamento
                change_message = ""
                if payment_data['change'] > 0:
                    change_message = f"\n🔄 Troco: R$ {payment_data['change']:.2f}"
                
                invoice_message = f"\n📄 Nota fiscal salva em: {invoice_path}" if invoice_path else "\n⚠️ Não foi possível gerar a nota fiscal."

                discount_val = payment_data.get('discount', 0)
                paid_total   = payment_data.get('total_with_discount', total)
                discount_msg = f"\n🏷️ Desconto: -R$ {discount_val:.2f}" if discount_val > 0 else ""

                QMessageBox.information(self, "✅ Venda Realizada!",
                                      f"Venda realizada com sucesso!\n"
                                      f"🛒 Subtotal: R$ {total:.2f}"
                                      f"{discount_msg}\n"
                                      f"💰 Total Pago: R$ {paid_total:.2f}\n"
                                      f"💵 Valor Recebido: R$ {payment_data['amount_received']:.2f}"
                                      f"{change_message}"
                                      f"{invoice_message}"
                                      f"{print_status}")
                
            except Exception as e:
                QMessageBox.critical(self, "❌ Erro", f"Erro ao processar venda: {str(e)}")
    
    def refresh_sales(self):
        sales = Sale.get_all()
        self.sales_table.setRowCount(len(sales))
        
        for row, sale in enumerate(sales):
            if str(sale.product_id).startswith('service_'):
                # É um serviço
                product_name = "SERVICO"
            else:
                product = Product.get_by_id(sale.product_id)
                product_name = product.name if product else "Produto não encontrado"
            
            # Adicionar emoji ao tipo de pagamento
            payment_display = sale.payment_type
            if sale.payment_type == "Dinheiro":
                payment_display = "💰 Dinheiro"
            elif sale.payment_type == "Cartão":
                payment_display = "💳 Cartão"
            elif sale.payment_type == "PIX":
                payment_display = "📱 PIX"
            
            self.sales_table.setItem(row, 0, QTableWidgetItem(str(sale.id)))
            self.sales_table.setItem(row, 1, QTableWidgetItem(product_name))
            self.sales_table.setItem(row, 2, QTableWidgetItem(str(sale.quantity)))
            self.sales_table.setItem(row, 3, QTableWidgetItem(f"R$ {sale.total_price:.2f}"))
            self.sales_table.setItem(row, 4, QTableWidgetItem(payment_display))
            self.sales_table.setItem(row, 5, QTableWidgetItem(sale.sale_date))

    # ========================= MÉTODOS DE SERVIÇOS =========================
    def add_service(self):
        dialog = ServiceDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_service_data()
            if data['description'] and data['value'] > 0:
                service = Service(
                    data['description'], 
                    data['value'], 
                    category=data['category'],
                    customer_name=data['customer_name']
                )
                service.save()
                
                # Se selecionado, adicionar ao carrinho
                if data['add_to_cart']:
                    service_cart_item = ServiceCartItem(service)
                    self.cart_items.append(service_cart_item)
                    self.refresh_cart_display()
                    
                    # Mudar para aba de vendas para mostrar o carrinho
                    self.tabs.setCurrentIndex(1)  # Índice da aba de vendas
                    
                    QMessageBox.information(self, "✅ Sucesso", 
                                          f"Serviço registrado e adicionado ao carrinho!\n"
                                          f"💰 Valor: R$ {data['value']:.2f}\n"
                                          f"Finalize a venda para gerar a nota fiscal.")
                else:
                    QMessageBox.information(self, "✅ Sucesso", 
                                          f"Serviço registrado com sucesso!\n"
                                          f"💰 Valor: R$ {data['value']:.2f}")
                
                self.refresh_all_data()
            else:
                QMessageBox.warning(self, "⚠️ Erro", "Descrição e valor são obrigatórios!")
    
    def edit_service(self):
        current_row = self.services_table.currentRow()
        if current_row >= 0:
            service_id = int(self.services_table.item(current_row, 0).text())
            service = Service.get_by_id(service_id)
            if service:
                dialog = ServiceDialog(service, parent=self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    data = dialog.get_service_data()
                    if data['description'] and data['value'] > 0:
                        service.description = data['description']
                        service.value = data['value']
                        service.category = data['category']
                        service.customer_name = data['customer_name']
                        service.save()
                        self.refresh_all_data()
                        QMessageBox.information(self, "✅ Sucesso", "Serviço editado com sucesso!")
                    else:
                        QMessageBox.warning(self, "⚠️ Erro", "Descrição e valor são obrigatórios!")
        else:
            QMessageBox.warning(self, "⚠️ Erro", "Selecione um serviço para editar!")
    
    def delete_service(self):
        current_row = self.services_table.currentRow()
        if current_row >= 0:
            service_id = int(self.services_table.item(current_row, 0).text())
            reply = QMessageBox.question(self, "❓ Confirmar", "Tem certeza que deseja excluir este serviço?")
            if reply == QMessageBox.StandardButton.Yes:
                Service.delete(service_id)
                self.refresh_all_data()
                QMessageBox.information(self, "✅ Sucesso", "Serviço excluído com sucesso!")
        else:
            QMessageBox.warning(self, "⚠️ Erro", "Selecione um serviço para excluir!")
    
    def refresh_services(self):
        services = Service.get_all()
        self.services_table.setRowCount(len(services))
        
        # Calcular totais
        current_month = datetime.now().strftime("%Y-%m")
        monthly_total = 0
        total_services = 0
        
        for row, service in enumerate(services):
            self.services_table.setItem(row, 0, QTableWidgetItem(str(service.id)))
            self.services_table.setItem(row, 1, QTableWidgetItem(service.description))
            self.services_table.setItem(row, 2, QTableWidgetItem(f"R$ {service.value:.2f}"))
            self.services_table.setItem(row, 3, QTableWidgetItem(service.category))
            self.services_table.setItem(row, 4, QTableWidgetItem(service.customer_name or "N/A"))
            self.services_table.setItem(row, 5, QTableWidgetItem(service.service_date))
            
            # Somar totais
            total_services += service.value
            if service.service_date.startswith(current_month):
                monthly_total += service.value
        
        # Atualizar labels de resumo
        self.monthly_services_label.setText(f"📅 Serviços do Mês: R$ {monthly_total:.2f}")
        self.total_services_label.setText(f"💰 Total Geral: R$ {total_services:.2f}")

    # ========================= MÉTODOS DE GASTOS =========================
    def add_expense(self):
        dialog = ExpenseDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_expense_data()
            if data and data['description']:
                # Verificar se há estoque suficiente para produtos
                if data['expense_type'] == 'product' and data['product_id']:
                    product = Product.get_by_id(data['product_id'])
                    if not product or product.quantity < data['quantity']:
                        QMessageBox.warning(self, "⚠️ Erro", "Estoque insuficiente para este produto!")
                        return
                
                expense = Expense(
                    data['description'], 
                    data['amount'], 
                    category=data['category'],
                    product_id=data['product_id'],
                    quantity=data['quantity'],
                    expense_type=data['expense_type']
                )
                expense.save()
                self.refresh_all_data()
                
                if data['expense_type'] == 'product':
                    product = Product.get_by_id(data['product_id'])
                    QMessageBox.information(self, "✅ Sucesso", 
                                          f"Gasto adicionado com sucesso!\n"
                                          f"Produto '{product.name}' foi descontado do estoque.")
                else:
                    QMessageBox.information(self, "✅ Sucesso", "Gasto adicionado com sucesso!")
            else:
                QMessageBox.warning(self, "⚠️ Erro", "Dados inválidos ou descrição obrigatória!")
    
    def edit_expense(self):
        current_row = self.expenses_table.currentRow()
        if current_row >= 0:
            expense_id = int(self.expenses_table.item(current_row, 0).text())
            expense = Expense.get_by_id(expense_id)
            if expense:
                dialog = ExpenseDialog(expense, parent=self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    data = dialog.get_expense_data()
                    if data and data['description']:
                        # Lógica simplificada para edição
                        expense.description = data['description']
                        expense.amount = data['amount']
                        expense.category = data['category']
                        expense.save()
                        self.refresh_all_data()
                        QMessageBox.information(self, "✅ Sucesso", "Gasto editado com sucesso!")
                    else:
                        QMessageBox.warning(self, "⚠️ Erro", "Dados inválidos ou descrição obrigatória!")
        else:
            QMessageBox.warning(self, "⚠️ Erro", "Selecione um gasto para editar!")
    
    def delete_expense(self):
        current_row = self.expenses_table.currentRow()
        if current_row >= 0:
            expense_id = int(self.expenses_table.item(current_row, 0).text())
            reply = QMessageBox.question(self, "❓ Confirmar", "Tem certeza que deseja excluir este gasto?")
            if reply == QMessageBox.StandardButton.Yes:
                Expense.delete(expense_id)
                self.refresh_all_data()
                QMessageBox.information(self, "✅ Sucesso", "Gasto excluído com sucesso!")
        else:
            QMessageBox.warning(self, "⚠️ Erro", "Selecione um gasto para excluir!")
    
    def refresh_expenses(self):
        expenses = Expense.get_all()
        self.expenses_table.setRowCount(len(expenses))
        
        # Calcular totais
        current_month = datetime.now().strftime("%Y-%m")
        monthly_total = 0
        total_expenses = 0
        
        for row, expense in enumerate(expenses):
            self.expenses_table.setItem(row, 0, QTableWidgetItem(str(expense.id)))
            
            # Descrição com indicação se é produto
            description = expense.description
            if expense.expense_type == "product" and expense.product_id:
                product = Product.get_by_id(expense.product_id)
                if product:
                    description += f" (Produto: {product.name} - Qtd: {expense.quantity})"
            
            self.expenses_table.setItem(row, 1, QTableWidgetItem(description))
            self.expenses_table.setItem(row, 2, QTableWidgetItem(f"R$ {expense.amount:.2f}"))
            self.expenses_table.setItem(row, 3, QTableWidgetItem(expense.category))
            self.expenses_table.setItem(row, 4, QTableWidgetItem(expense.expense_date))
            
            # Somar totais
            total_expenses += expense.amount
            if expense.expense_date.startswith(current_month):
                monthly_total += expense.amount
        
        # Atualizar labels de resumo
        self.monthly_expenses_label.setText(f"📅 Gastos do Mês: R$ {monthly_total:.2f}")
        self.total_expenses_label.setText(f"💰 Total Geral: R$ {total_expenses:.2f}")

    # ========================= MÉTODOS DE CAIXA =========================
    def refresh_cash(self):
        with sqlite3.connect("stock_manager.db", timeout=30.0) as conn:
            cursor = conn.cursor()

            # --- Total em Caixa: soma de todas as vendas ---
            cursor.execute("SELECT COALESCE(SUM(total_price), 0) FROM sales")
            total_cash = cursor.fetchone()[0]
            self.cash_label.setText(f"🏦 Total em Caixa\nR$ {total_cash:,.2f}")

            # --- Total da Semana: segunda a sábado da semana atual ---
            cursor.execute("""
                SELECT COALESCE(SUM(total_price), 0)
                FROM sales
                WHERE
                    DATE(sale_date) >= DATE(
                        'now',
                        'localtime',
                        '-' || CAST((
                            CASE CAST(strftime('%w', 'now', 'localtime') AS INTEGER)
                                WHEN 0 THEN 6
                                ELSE CAST(strftime('%w', 'now', 'localtime') AS INTEGER) - 1
                            END
                        ) AS TEXT) || ' days'
                    )
                    AND CAST(strftime('%w', DATE(sale_date)) AS INTEGER) BETWEEN 1 AND 6
            """)
            total_week = cursor.fetchone()[0]
            self.week_label.setText(f"📅 Total da Semana (Seg–Sáb)\nR$ {total_week:,.2f}")

            # --- Total do Mês atual ---
            cursor.execute("""
                SELECT COALESCE(SUM(total_price), 0)
                FROM sales
                WHERE strftime('%Y-%m', sale_date) = strftime('%Y-%m', 'now', 'localtime')
            """)
            total_month = cursor.fetchone()[0]
            self.month_label.setText(f"🗓️ Total do Mês\nR$ {total_month:,.2f}")

        # Atualizar cards de lucro
        self.update_day_profit()
        self.update_month_profit()

    def update_daily_profit(self):
        # Mantido para compatibilidade
        self.update_day_profit()

    def update_day_profit(self):
        """Atualiza o card de lucro do dia selecionado"""
        selected_date = self.date_edit.date().toString("yyyy-MM-dd")
        profit = Product.get_daily_profit(selected_date)
        color = "#00CED1" if profit >= 0 else "#FF4444"
        sign = "" if profit >= 0 else "-"
        self.daily_profit_label.setText(
            f"💸 Vendas do Dia ({selected_date})\n{sign}R$ {abs(profit):,.2f}"
        )
        self.daily_profit_label.setStyleSheet(f"""
            color: {color};
            padding: 25px 15px;
            border: 3px solid {color};
            border-radius: 12px;
            background-color: #051515;
            margin: 2px;
        """)

    def update_month_profit(self):
        """Atualiza o card de lucro do mês/ano selecionado"""
        month = self.month_combo.currentIndex() + 1  # 1–12
        year = self.year_spin.value()
        month_str = f"{year}-{month:02d}"

        # Calcular lucro do mês: itera por dia do mês
        import calendar
        total_profit = 0.0
        days_in_month = calendar.monthrange(year, month)[1]
        for day in range(1, days_in_month + 1):
            date_str = f"{year}-{month:02d}-{day:02d}"
            total_profit += Product.get_daily_profit(date_str)

        color = "#FF8C00" if total_profit >= 0 else "#FF4444"
        sign = "" if total_profit >= 0 else "-"
        month_name = self.month_combo.currentText()
        self.monthly_profit_label.setText(
            f"📈 Lucro — {month_name} / {year}\n{sign}R$ {abs(total_profit):,.2f}"
        )
        self.monthly_profit_label.setStyleSheet(f"""
            color: {color};
            padding: 25px 15px;
            border: 3px solid {color};
            border-radius: 12px;
            background-color: #1a0800;
            margin: 2px;
        """)

    # ========================= MÉTODOS DE CONFIGURAÇÕES =========================
    def refresh_printers(self):
        """Atualiza a lista de impressoras disponíveis"""
        self.printer_combo.clear()
        self.printer_combo.addItem("Nenhuma impressora selecionada", None)
        
        available_printers = self.printer_manager.get_available_printers()
        for printer in available_printers:
            self.printer_combo.addItem(printer, printer)
        
        # Selecionar impressora padrão se disponível
        default_printer = self.printer_manager.get_default_printer()
        if default_printer and default_printer in available_printers:
            index = self.printer_combo.findData(default_printer)
            if index >= 0:
                self.printer_combo.setCurrentIndex(index)
                self.printer_manager.set_selected_printer(default_printer)
        
        self.update_printer_status()
    
    def on_printer_selected(self, printer_name):
        """Callback quando uma impressora é selecionada"""
        if printer_name == "Nenhuma impressora selecionada":
            self.printer_manager.selected_printer = None
        else:
            selected_data = self.printer_combo.currentData()
            if selected_data:
                self.printer_manager.set_selected_printer(selected_data)
        
        self.update_printer_status()
    
    def on_auto_print_toggled(self, checked):
        """Callback quando a impressão automática é alterada"""
        self.printer_manager.set_auto_print(checked)
    
    def update_printer_status(self):
        """Atualiza o status da impressora na interface"""
        if not self.printer_manager.selected_printer:
            self.printer_status_label.setText("❌ Nenhuma impressora selecionada")
            self.printer_status_label.setStyleSheet("color: #666; padding: 10px; background-color: #f0f0f0; border-radius: 5px; margin: 5px 0;")
        else:
            status = self.printer_manager.get_printer_status()
            if status['available']:
                status_text = f"✅ {status['name']}"
                if status['is_default']:
                    status_text += " (Padrão)"
                status_text += " - Pronta para imprimir"
                self.printer_status_label.setStyleSheet("color: #2E8B57; padding: 10px; background-color: #F0FFF0; border-radius: 5px; margin: 5px 0;")
            else:
                status_text = f"❌ {status['name']} - Não disponível"
                self.printer_status_label.setStyleSheet("color: #DC143C; padding: 10px; background-color: #FFE4E1; border-radius: 5px; margin: 5px 0;")
            
            self.printer_status_label.setText(status_text)
    
    def test_print(self):
        """Imprime uma página de teste"""
        if not self.printer_manager.selected_printer:
            QMessageBox.warning(self, "⚠️ Erro", "Selecione uma impressora primeiro!")
            return
        
        try:
            success = self.printer_manager.print_test_page()
            if success:
                QMessageBox.information(self, "✅ Sucesso", "Página de teste enviada para impressão!")
            else:
                QMessageBox.warning(self, "⚠️ Erro", "Falha ao enviar página de teste para impressão.")
        except Exception as e:
            QMessageBox.critical(self, "❌ Erro", f"Erro ao imprimir página de teste: {str(e)}")

    # ========================= MÉTODOS DE TEMAS =========================
    def apply_light_theme(self, checked):
        """Aplica o tema claro quando o radio button é selecionado"""
        if checked:
            self.apply_theme_light()
    
    def apply_dark_theme(self, checked):
        """Aplica o tema escuro quando o radio button é selecionado"""
        if checked:
            self.apply_theme_dark()
    
    def apply_selected_theme(self):
        """Aplica o tema selecionado baseado no radio button"""
        if self.light_theme_radio.isChecked():
            self.apply_theme_light()
            QMessageBox.information(self, "🎨 Tema", "Tema Claro aplicado com sucesso!")
        elif self.dark_theme_radio.isChecked():
            self.apply_theme_dark()
            QMessageBox.information(self, "🎨 Tema", "Tema Escuro aplicado com sucesso!")
    
    def apply_theme_light(self):
        """Aplica o tema claro"""
        # Estilo das abas para tema claro
        light_tab_style = """
            QTabWidget::pane {
                border: 1px solid #ddd;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #f0f0f0;
                color: #00FF00;
                padding: 15px 25px;
                margin-right: 2px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: #4CAF50;
                color: #00FF00;
                font-weight: bold;
            }
            QTabBar::tab:hover {
                background-color: #e0e0e0;
                color: #00FF00;
            }
        """
        
        # Aplicar estilo das abas
        self.tabs.setStyleSheet(light_tab_style)
        
        # Atualizar todas as tabelas para o tema claro
        self.apply_light_theme_to_tables()
        
        # Estilo do fundo principal
        self.setStyleSheet("""
            QMainWindow {
                background-color: white;
                color: black;
            }
            QWidget {
                background-color: white;
                color: black;
            }
            QLabel {
                color: black;
            }
        """)
    
    def apply_theme_dark(self):
        """Aplica o tema escuro"""
        # Estilo das abas para tema escuro
        dark_tab_style = """
            QTabWidget::pane {
                border: 1px solid #555;
                background-color: #2b2b2b;
            }
            QTabBar::tab {
                background-color: #404040;
                color: #00FF00;
                padding: 15px 25px;
                margin-right: 2px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: #4CAF50;
                color: #00FF00;
                font-weight: bold;
            }
            QTabBar::tab:hover {
                background-color: #555;
                color: #00FF00;
            }
        """
        
        # Aplicar estilo das abas
        self.tabs.setStyleSheet(dark_tab_style)
        
        # Atualizar todas as tabelas para o tema escuro
        self.apply_dark_theme_to_tables()
        
        # Estilo do fundo principal escuro
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1a1a1a;
                color: #00FF00;
            }
            QWidget {
                background-color: #2b2b2b;
                color: #00FF00;
            }
            QLabel {
                color: #00FF00 !important;
                background-color: #2b2b2b;
            }
            QPushButton {
                background-color: #404040;
                color: #00FF00 !important;
                border: 2px solid #555;
                padding: 8px;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #555;
                border: 2px solid #00FF00;
            }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
                background-color: #404040;
                color: #00FF00 !important;
                border: 2px solid #555;
                padding: 5px;
                border-radius: 4px;
            }
            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {
                border: 2px solid #00FF00;
            }
            QTableWidget {
                background-color: #2b2b2b;
                color: #00FF00 !important;
                gridline-color: #555;
                alternate-background-color: #404040;
                border: 1px solid #555;
            }
            QTableWidget::item {
                color: #00FF00 !important;
                background-color: #2b2b2b;
                padding: 5px;
            }
            QTableWidget::item:alternate {
                background-color: #404040;
            }
            QTableWidget::item:selected {
                background-color: #4CAF50;
                color: #000000 !important;
            }
            QHeaderView::section {
                background-color: #1a1a1a;
                color: #00FF00 !important;
                border: 1px solid #555;
                padding: 8px;
                font-weight: bold;
            }
            QRadioButton {
                color: #00FF00 !important;
                background-color: #2b2b2b;
            }
            QCheckBox {
                color: #00FF00 !important;
                background-color: #2b2b2b;
            }
            QDateEdit {
                background-color: #404040;
                color: #00FF00 !important;
                border: 2px solid #555;
                padding: 5px;
                border-radius: 4px;
            }
            QScrollBar:vertical {
                background-color: #404040;
                width: 15px;
                border-radius: 7px;
            }
            QScrollBar::handle:vertical {
                background-color: #00FF00;
                border-radius: 7px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #4CAF50;
            }
        """)
    
    def apply_dark_theme_to_tables(self):
        """Aplica tema escuro específico para todas as tabelas"""
        dark_table_style = """
            QTableWidget {
                background-color: #2b2b2b;
                color: #00FF00;
                gridline-color: #555;
                alternate-background-color: #404040;
                border: 1px solid #555;
            }
            QTableWidget::item {
                color: #00FF00;
                background-color: #2b2b2b;
                padding: 5px;
            }
            QTableWidget::item:alternate {
                background-color: #404040;
            }
            QTableWidget::item:selected {
                background-color: #4CAF50;
                color: #000000;
            }
            QHeaderView::section {
                background-color: #1a1a1a;
                color: #00FF00;
                border: 1px solid #555;
                padding: 8px;
                font-weight: bold;
            }
        """
        
        # Aplicar em todas as tabelas
        if hasattr(self, 'products_table'):
            self.products_table.setStyleSheet(dark_table_style)
        if hasattr(self, 'available_products_table'):
            self.available_products_table.setStyleSheet(dark_table_style)
        if hasattr(self, 'cart_table'):
            self.cart_table.setStyleSheet(dark_table_style)
        if hasattr(self, 'sales_table'):
            self.sales_table.setStyleSheet(dark_table_style)
        if hasattr(self, 'services_table'):
            self.services_table.setStyleSheet(dark_table_style)
        if hasattr(self, 'expenses_table'):
            self.expenses_table.setStyleSheet(dark_table_style)
    
    def apply_light_theme_to_tables(self):
        """Restaura tema claro específico para todas as tabelas"""
        # Restaurar estilos originais para cada tabela
        if hasattr(self, 'products_table'):
            self.products_table.setStyleSheet("""
                QTableWidget {
                    gridline-color: #ddd;
                    background-color: white;
                    alternate-background-color: #f9f9f9;
                    font-size: 11px;
                    color: #00FF00;
                }
                QHeaderView::section {
                    background-color: #4CAF50;
                    color: white;
                    padding: 12px;
                    border: none;
                    font-weight: bold;
                    font-size: 12px;
                }
                QTableWidget::item {
                    padding: 8px;
                    border-bottom: 1px solid #eee;
                    color: #00FF00;
                }
                QTableWidget::item:selected {
                    background-color: #E8F5E8;
                    color: #00FF00;
                }
            """)

    # ========================= MÉTODOS AUXILIARES =========================
    def refresh_all_data(self):
        """Atualiza todos os dados das abas"""
        self.refresh_products()
        self.refresh_available_products()
        self.refresh_sales()
        self.refresh_services()
        self.refresh_expenses()
        self.refresh_cash()
        self.refresh_cart_display()

# ========================= FUNÇÃO PRINCIPAL =========================
def main():
    print("🚀 Iniciando Sistema ADR INFO...")
    
    # Inicializar banco de dados
    init_db()
    print("📦 Banco de dados inicializado!")
    
    # Criar aplicação
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Usar estilo moderno
    print("🖥️ Aplicação PyQt6 criada!")
    
    # Criar e mostrar janela principal
    window = MainWindow()
    window.show()
    print("✅ Janela principal exibida!")
    
    print("🎯 Sistema completo pronto para uso!")
    print("📱 Todas as funcionalidades disponíveis:")
    print("   📦 Gestão de Produtos")
    print("   🛍️ Sistema de Vendas com Carrinho")
    print("   🔧 Gestão de Serviços")
    print("   💸 Controle de Gastos")
    print("   💰 Controle de Caixa e Lucros")
    print("   📄 Geração de Notas Fiscais")
    print("   🖨️ Sistema de Impressão")
    print("   ⚙️ Configurações")
    
    # Executar aplicação
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
