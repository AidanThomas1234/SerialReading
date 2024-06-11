import os
import serial
import time
import mysql.connector
from datetime import datetime
import threading
import msvcrt  # For Windows getch() function
import gc
import re
# Used to print files 
import win32ui
import win32con

exit_flag = False

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Extracts the number from a string
def extract_number(input_string):
    pattern = r"\d+(\.\d+)?"
    match = re.search(pattern, input_string)
    if match:
        extracted_number = match.group(0)
        return int(float(extracted_number))  # Convert to integer
    else:
        print("No numbers found in the input string.")
        return None

# Database connection pool
connection_pool = mysql.connector.pooling.MySQLConnectionPool(
    pool_name="mypool",
    pool_size=5,
    host="plesk.remote.ac",
    user="ws330240_Alistair",
    password="ea#4M786q",
    database="ws330240_AandR"
)

def get_db_connection():
    return connection_pool.get_connection()

def print_file_to_printer(data_string):
    printer_name = "Alistairsprinter"  # Adjust to your printer
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)
    hdc.StartDoc("Weigh Head Scaled")
    hdc.StartPage()
    
    y = 100  # Starting Y position

    lines = data_string.split('\n')
    
    for line in lines:
        # Determine the font size and style based on the content of the line
        if 'BagID' in line:
            font_size = 200
            font_style = win32con.FW_BOLD
            font_italic = False
        elif 'Product' in line:
            font_size = 180
            font_style = win32con.FW_NORMAL
            font_italic = False
        else:
            font_size = 160
            font_style = win32con.FW_NORMAL
            font_italic = False

        # Create and select the font
        font = win32ui.CreateFont({
            "name": "Arial",  # Use the desired font name
            "height": font_size,
            "weight": font_style,
            "italic": font_italic,
        })
        hdc.SelectObject(font)
        
        # Write the actual line of text
        hdc.TextOut(100, y, line.strip())
        y += font_size * 10  # Increment Y position for next line based on font size

    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()

def read_serial_data(port, baud_rate, ser=None):
    global exit_flag
    if ser is None:
        ser = serial.Serial(port, baud_rate)  # Open the serial port
        print(f"Opened serial port {port}")

    try:
        read_count = 0  
        batch_number = get_batch_number()
        if batch_number is None:  
            return

        product = get_product_type()
        if product is None: 
            return

        while True:
            if exit_flag:
                print("Exiting to menu...")
                break  # Exit loop

            if ser.in_waiting > 0:
                data = ser.read(ser.in_waiting).decode("utf-8")
                lines = data.split('\n')
                
                for decoded_data in lines:
                    if decoded_data.startswith("Gross"):
                        number = extract_number(decoded_data)
                        read_count += 1

                        if read_count % 22 == 0:
                            batch_number = get_batch_number()
                            product = get_product_type()

                        current_time = datetime.now()
                        mydb = get_db_connection()
                        mycursor = mydb.cursor()
                        sql = "INSERT INTO `LakesWeighHead` (`BagID`, `GrossWeight`, `DateandTime`, `BatchNumb`, `ProductType`) VALUES ('', %s, %s, %s, %s)"
                        val = (number, current_time, batch_number, product)
                        mycursor.execute(sql, val)
                        mydb.commit()

                        idsql = "SELECT `BagID` FROM `LakesWeighHead` ORDER BY `BagID` DESC LIMIT 1;"
                        mycursor.execute(idsql)
                        CurrentID = mycursor.fetchone()
                        most_recent_id = CurrentID[0]

                        print(f"\n\nBatch: {batch_number}   Weight: {number}    Product: {product}    BagID: {most_recent_id}       Date and time: {current_time}")
                        
                        data_string = (
                            f"BagID: {most_recent_id}\n\n"
                            f"Product Type: {product}\n\n"
                            f"Batch: {batch_number}\n\n"
                            f"Weight: {number}\n\n"
                            f"Date and Time: {current_time}\n\n"
                        )

                        print_file_to_printer(data_string)

                        print(f"Data saved to database with BagID: {most_recent_id}")
                        mycursor.close()
                        mydb.close()

                        gc.collect()

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        ser.close()  # Close the serial port

def get_batch_number():
    while True:
        batch_number = input("Enter batch number (3 digits, e.g., 123) or 'exit' to return to menu: ").strip()
        if batch_number.lower() == 'exit':
            return None
        if len(batch_number) == 3 and batch_number.isdigit():
            return batch_number
        else:
            print("Invalid input. Please enter a 3-digit batch number.")

def get_product_type():
    while True:
        product_type = input("Enter product type or 'exit' to return to menu: ").strip()
        if product_type.lower() == 'exit':
            return None
        return product_type

def update_field():
    try:
        mydb = get_db_connection()
        mycursor = mydb.cursor()

        sql = "SELECT * FROM `LakesWeighHead`"
        mycursor.execute(sql)
        results = mycursor.fetchall()
        for row in results:
            print(row)

        record_id = input("Enter the BagID of the record you want to update: ")
        field_name = input("Enter the field you want to update (GrossWeight, BatchNumb, ProductType, DateandTime): ")
        new_value = input("Enter the new value: ")

        sql = f"UPDATE `LakesWeighHead` SET `{field_name}` = %s WHERE `BagID` = %s"
        val = (new_value, record_id)
        mycursor.execute(sql, val)
        mydb.commit()

        print(f"Record {record_id} updated successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        mycursor.close()
        mydb.close()

def weigh_head(port, baud_rate):
    read_serial_data(port, baud_rate)

def menu(port, baud_rate, ser):
    global exit_flag
    while True:
        print("\nWeigh Head System\n")
        print("1) Start Weighing")
        print("2) Update a Field")
        print("0) Exit")

        user_input = input().strip()
        if user_input == '0':
            exit_flag = True
            return
        elif user_input == '1':
            weigh_head(port, baud_rate)
        elif user_input == '2':
            update_field()
        else:
            print("Invalid input. Please enter a valid number.")

if __name__ == "__main__":
    port = 'COM3'
    baud_rate = 9600

    ser = serial.Serial(port, baud_rate)  # Open the serial port
    print(f"Opened serial port {port}")

    menu(port, baud_rate, ser)

    ser.close()  # Close the serial port
    input("Press Enter to exit...")
