import serial
import time
import win32print
import mysql.connector
import win32ui
import re
import datetime
from datetime import datetime
import pytz
import os



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


# Function to print text to a specific printer
def print_to_printer(text):
    printer_name = "Alistairsprinter"  # Adjust to your printer
    hdc = win32ui.CreateDC()


    hdc.CreatePrinterDC(printer_name)
    hdc.StartDoc("Weigh Head Scaled")
    hdc.StartPage()
    hdc.TextOut(100, 100, text)
    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()


# Database connection
mydb = mysql.connector.connect(
    host="#",
    user="#",
    password="#",
    database="#"
)

mycursor = mydb.cursor()


def read_serial_data(port, baud_rate):
    ser = serial.Serial(port, baud_rate)
    read_count = 0  # Counter to track the number of readings
    
    # Prompt for user input at the start
    user_input = input("Please enter the batch number to begin: ")
    product = input("Product Type: ")
    print("Batch:", user_input,"Product:",product)
    
    try:
        while True:
            if ser.in_waiting > 0:
                data = ser.readline().strip()
                decoded_data = data.decode("utf-8")

                if decoded_data.startswith("Gross"):
                    print(decoded_data)
                    print_to_printer(decoded_data)
                    number = extract_number(decoded_data)

                    # Increment the read count
                    read_count += 1
                    
                    # Prompt for user input every 22 readings
                    if read_count % 22 == 0:
                        user_input = input("Please enter the batch number to begin: ")
                        product = input("Product Type: ")
                        print("Batch:", user_input,"Product:",product)
                        
                    
                    # Insert into the database
                    current_time = datetime.now()
                    sql = "#"
                    val = (number, current_time, user_input,product)
                    mycursor.execute(sql, val)
                    mydb.commit()
            
            time.sleep(0.1)  # Small delay to prevent CPU overload

    except KeyboardInterrupt:
        print("Interrupted by user. Exiting.")

    finally:
        ser.close()  # Close the serial connection


if __name__ == "__main__":
    port = 'COM3'  # Set your COM port
    baud_rate = 9600  # Set the baud rate
    read_serial_data(port, baud_rate)
