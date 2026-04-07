import customtkinter
import pyodbc
import pandas as pd
from tkinter import messagebox

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

# Функция для подключения к базе данных
def connect_to_database():
    global conn, cursor
    try:
        conn = pyodbc.connect('DRIVER={SQL Server};SERVER=DESKTOP-L1C9AQR;DATABASE=prodaz;Trusted_Connection=yes;')
        cursor = conn.cursor()
    except pyodbc.Error as e:
        messagebox.showerror("Database Connection Error", f"An error occurred while connecting to the database: {e}")
        root.destroy()

# Функция для запуска генерации отчёта
def generate_report(report_type, sale_date=None, customer_id=None, customer_name=None, product_name=None, category_name=None):
    try:
        # Проверка соединения с базой данных
        if not conn or not cursor:
            messagebox.showerror("Database Error", "Database connection not established. Please try again.")
            return

        # Проверка даты продажи
        if sale_date:
            try:
                pd.to_datetime(sale_date)
            except ValueError:
                messagebox.showerror("Input Error", "Invalid date format. Please use YYYY-MM-DD.")
                return

        # Проверка идентификатора клиента
        if customer_id:
            try:
                int(customer_id)
            except ValueError:
                messagebox.showerror("Input Error", "Invalid customer ID. Please enter a numeric value.")
                return

        # Проверка имени клиента
        if customer_name and not customer_name.replace(' ', '').isalnum():
            messagebox.showerror("Input Error", "Invalid customer name. Please enter alphanumeric characters only.")
            return

        if report_type == 'Sales Report':
            query = """
            SELECT Sales.SaleID, Sales.SaleDate, Sales.CustomerID, Customers.CustomerName, Sales.ProductID, Products.ProductName, Products.ProductDescription, Categories.CategoryName, Sales.Quantity, Sales.TotalAmount
            FROM Sales
            INNER JOIN Customers ON Sales.CustomerID = Customers.CustomerID
            INNER JOIN Products ON Sales.ProductID = Products.ProductID
            INNER JOIN Categories ON Products.CategoryID = Categories.CategoryID
            WHERE 1=1
            """
            params = []
            if sale_date:
                query += " AND Sales.SaleDate = ?"
                params.append(sale_date)
            if customer_id:
                query += " AND Sales.CustomerID = ?"
                params.append(customer_id)
            if customer_name:
                query += " AND Customers.CustomerName = ?"
                params.append(customer_name)
            if product_name:
                query += " AND Products.ProductName = ?"
                params.append(product_name)
            if category_name:
                query += " AND Categories.CategoryName = ?"
                params.append(category_name)
            cursor.execute(query, *params)
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror("No Data Found", "No sales data found with the provided filters.")
                return
            df = pd.DataFrame.from_records(rows, columns=['SaleID', 'SaleDate', 'CustomerID', 'CustomerName', 'ProductID', 'ProductName', 'ProductDescription', 'CategoryName', 'Quantity', 'TotalAmount'])
            df.to_excel('sales_report.xlsx', index=False)  # Save as Excel file
            messagebox.showinfo("Report Generated", "Sales report has been generated successfully.")

        elif report_type == 'Customers Report':
            query = """
            SELECT Customers.CustomerID, Customers.CustomerName, Customers.Address, Customers.ContactInfo, Products.ProductName, Categories.CategoryName
            FROM Customers
            INNER JOIN Sales ON Customers.CustomerID = Sales.CustomerID
            INNER JOIN Products ON Sales.ProductID = Products.ProductID
            INNER JOIN Categories ON Products.CategoryID = Categories.CategoryID
            WHERE 1=1
            """
            params = []
            if customer_id:
                query += " AND Customers.CustomerID = ?"
                params.append(customer_id)
            if customer_name:
                query += " AND Customers.CustomerName = ?"
                params.append(customer_name)
            if product_name:
                query += " AND Products.ProductName = ?"
                params.append(product_name)
            if category_name:  # Добавляем фильтрацию по имени категории
                query += " AND Categories.CategoryName = ?"
                params.append(category_name)
            cursor.execute(query, *params)
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror("No Data Found", "No customers data found with the provided filters.")
                return
            df = pd.DataFrame.from_records(rows, columns=['CustomerID', 'CustomerName', 'Address', 'ContactInfo', 'ProductName', 'CategoryName'])
            
            # Если прошел проверку , save the report to Excel
            df.to_excel('customers_report.xlsx', index=False)  # Save as Excel file
            messagebox.showinfo("Report Generated", "Customers report has been generated successfully.")

    except pyodbc.Error as e:
        messagebox.showerror("Database Error", f"An error occurred while executing the database query: {e}")
    except PermissionError:
        messagebox.showerror("Permission Error", "You do not have permission to write to the file. Please close any programs that might be using the file and try again.")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

# Функция для закрытия соединения с базой данных и выхода из приложения
def close_application():
    try:
        cursor.close()
        conn.close()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while closing the database connection: {e}")
    root.destroy()

# Создание главного окна приложения
root = customtkinter.CTk()
root.title("Report Generator")

# Подключение к базе данных при запуске приложения
connect_to_database()

# Создание полей ввода для выбора даты и фильтрации покупателей, продуктов и категорий
sale_date_entry = customtkinter.CTkEntry(root, placeholder_text="Sale Date (YYYY-MM-DD)")
sale_date_entry.pack(padx=10, pady=10)

customer_id_entry = customtkinter.CTkEntry(root, placeholder_text="Customer ID")
customer_id_entry.pack(padx=10, pady=10)

customer_name_entry = customtkinter.CTkEntry(root, placeholder_text="Customer Name")
customer_name_entry.pack(padx=10, pady=10)

product_name_entry = customtkinter.CTkEntry(root, placeholder_text="Product Name")
product_name_entry.pack(padx=10, pady=10)

category_name_entry = customtkinter.CTkEntry(root, placeholder_text="Category Name")
category_name_entry.pack(padx=10, pady=10)

# Создание кнопок для генерации отчетов
sales_report_button = customtkinter.CTkButton(root, text="Generate Sales Report", command=lambda: generate_report('Sales Report', sale_date_entry.get(), customer_id_entry.get(), customer_name_entry.get(), product_name_entry.get(), category_name_entry.get()))
sales_report_button.pack(fill='x', padx=20, pady=10)

customers_report_button = customtkinter.CTkButton(root, text="Generate Customers Report", command=lambda: generate_report('Customers Report', sale_date_entry.get(), customer_id_entry.get(), customer_name_entry.get(), product_name_entry.get(), category_name_entry.get()))
customers_report_button.pack(fill='x', padx=20, pady=10)

# Кнопка для выхода из приложения
exit_button = customtkinter.CTkButton(root, text="Exit", command=close_application)
exit_button.pack(fill='x', padx=20, pady=10)

# Запуск главного цикла обработки событий
root.protocol("WM_DELETE_WINDOW", close_application)  
root.mainloop()