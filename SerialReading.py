#Import Area 
import serial
import time
import win32print
import mysql.connector
import win32ui
import re
import datetime
from datetime import datetime
import pytz


def extract_number(input_string):
    # Regular expression to match numbers (integers and floats)
    # This pattern matches both integers and decimal numbers
    pattern = r"\d+(\.\d+)?"
    
    # Search for the pattern in the input string
    match = re.search(pattern, input_string)

    if match:
        # Extract the matched number
        extracted_number = match.group(0)  # This gets the first matching group
        
        # Convert to float (to handle both integers and floats)
        number = float(extracted_number)
        
        # If you prefer integers, you can convert to int
        integer_part = int(number)
        
        return integer_part
    else:
        print("No numbers found in the input string.")
        return None


##Printer Functions 
def print_to_printer(text):

    #PrinterName(Defult)
    #printer_name = win32print.GetDefaultPrinter()

    #PrinterName(specific)
    printer_name = "Alistairsprinter"

 
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)

  
    hdc.StartDoc("Python Printing Job")
    hdc.StartPage()

    
    hdc.TextOut(100, 100, text)


    hdc.EndPage()
    hdc.EndDoc()

 
    hdc.DeleteDC()

##SQL connections
mydb=mysql.connector.connect(
    host="#",
    user="#",
    password="#",
    database="#"
)  

 

  #sql statment


    
mycursor=mydb.cursor()

def read_serial_data(port, baud_rate):

    # Open the serial connection
    ser = serial.Serial(port, baud_rate)
    
    try:
      
        while True:
           
            if ser.in_waiting > 0:
               
                data = ser.readline().strip()

           
                decoded_data = data.decode("utf-8")

              #Only read lines begging with Gross
                if decoded_data.startswith("Gross"):
                    #1WeighHeadOutput=("Received from:",port,":", decoded_data)

                    #################################################################
                    #When the weigh head button is pressed this happens 

                    ##print to terminal 
                    print("Received from:",port,":", decoded_data)

                    #Print 
                    x=(decoded_data)
                    print_to_printer(x)
                  
                  
                    #Seperate the numbers from the text
                    number = extract_number(x)
                    print(number)

                    #Datetime
                    L=datetime.now()

                    ##SQL
                    sql=""
                    val=(number,L)
                    mycursor.execute(sql,val)
                    mydb.commit()
                    


















            
            time.sleep(0.1) 

    except KeyboardInterrupt:
        print("Interrupted by user. Exiting.")

    finally:
 
        ser.close()
#########################################################


if __name__ == "__main__":


    # Serial port settings
    port = 'COM3'  
    baud_rate = 9600  

   
    read_serial_data(port, baud_rate)


    #https://github.com/AidanThomas1234