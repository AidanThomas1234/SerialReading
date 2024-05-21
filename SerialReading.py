import serial
import time
import win32print
import mysql.connector
import win32ui
import re
from datetime import datetime, timedelta
import threading
import msvcrt  # For Windows getch() function
import csv
from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_TABLE_ALIGNMENT




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
def print_file_to_word_doc(file_path, report_date):
    doc = Document()
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'  # Apply table style

    # Add header row with column names
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'BagID'
    hdr_cells[1].text = 'GrossWeight'
    hdr_cells[2].text = 'BatchNumb'
    hdr_cells[3].text = 'ProductType'
    hdr_cells[4].text = 'DateandTime'

    try:
        with open(file_path, 'r') as file:
            csv_reader = csv.reader(file)
            next(csv_reader)  # Skip header row
            for row in csv_reader:
                row_cells = table.add_row().cells
                for i, cell_value in enumerate(row):
                    row_cells[i].text = cell_value
    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
        return
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        return

    # Format table
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)  # Adjust font size

    # Center align text in all cells
    for row in table.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER

    doc.save(f'report_{report_date.strftime("%Y-%m-%d_%H-%M-%S")}.docx') # Save Word document with date and time
    pathend=(f'report_{report_date.strftime("%Y-%m-%d_%H-%M-%S")}.docx')
    wrd_file_path = "C:/Users/Projects.PURNHOUSEFARM/"+ pathend
    print(wrd_file_path)



#Used to prints files 
def print_file_to_printer(file_path):
    printer_name = "Alistairsprinter"  # Adjust to your printer
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)
    hdc.StartDoc("Weigh Head Scaled")
    hdc.StartPage()
    y = 100  # Starting Y position

    try:
        with open(file_path, 'r') as file:
            for line in file:
                # Write a blank line to leave a gap between each line of text
                hdc.TextOut(100, y, "")
                y += 100  # Increment Y position for the blank line
                
                # Write the actual line of text
                hdc.TextOut(100, y, line.strip())
                y += 100  # Increment Y position for next line
    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
        return
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        return

    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()

# Database connection
mydb = mysql.connector.connect(
    host="plesk.remote.ac",
    user="ws330240_Alistair",
    password="ea#4M786q",
    database="ws330240_AandR"
)

mycursor = mydb.cursor()

exit_flag = False


def read_serial_data(port, baud_rate,ser):
    try:
        global exit_flag
        ser = serial.Serial(port, baud_rate)  # Open the serial port
        read_count = 0  # Counter to track the number of readings
        
        # Rest of your function code...

    except Exception as e:
        handle_error=(f"An error occurred while reading serial data: {e}")
    finally:
        if ser.is_open:
            ser.close()  # Close the serial port in the finally block
    global exit_flag
    ser = serial.Serial(port, baud_rate)

    read_count = 0  # Counter to track the number of readings
    
    # Prompt for user input at the start
    while True:
        user_input = input("Please enter the batch number to begin : ")
        if user_input == '0':
            menu(port, baud_rate,ser)
            return  
        1
        try:
            user_input = int(user_input)  # Try to convert the input to an integer
            break  # Break out of the2 loop if conversion is successful
        except ValueError:
            print("Invalid input. Please enter an integer.")
  
  
    product = None  # Initialize product to None

    product_input = int(input("""Enter the Product Type:\n
    1) Skins\n
    2) Sticks\n"""))

    if product_input == 1:
        product = "Skins"  # Assign "Skins" to product
    elif product_input == 2:
        product = "Sticks"  # Assign "Sticks" to product
    else:
        print("Invalid Input. Restarting...")
        read_serial_data(port, baud_rate,ser)


    try:
        while True:
            if exit_flag:
                print("Exiting to menu...")
                menu(port, baud_rate,ser)
                return  # Exit to menu

            if ser.in_waiting > 0:
                data = ser.readline().strip()
                decoded_data = data.decode("utf-8")

                if decoded_data.startswith("Gross"):
                    # print_to_printer(decoded_data)
                    number = extract_number(decoded_data)

                    # Increment the read count
                    read_count += 1
                    
                    # Prompt for user input every 22 readings
                    if read_count % 22 == 0:
                        while True:
                            user_input = input("Please enter the batch number to begin (or enter 0 to exit): ")
                            if user_input == '0':
                                menu(port, baud_rate,ser)
                                return  # Exit to menu
                            

                            product = None  # Initialize product to None

                            product_input = int(input("""Enter the Product Type:\n
                            1) Skins\n
                            2) Sticks\n"""))

                            if product_input == 1:
                                product = "Skins"  # Assign "Skins" to product
                            elif product_input == 2:
                                product = "Sticks"  # Assign "Sticks" to product
                            else:
                                print("Invalid Input. Restarting...")
                                read_serial_data(port, baud_rate,ser)

                
                            try:
                                user_input = int(user_input)  # Try to convert the input to an integer
                                break  # Break out of the loop if conversion is successful
                            except ValueError:
                                print("Invalid input. Please enter an integer.")
                    
                    # Insert into the database
                    current_time = datetime.now()
                    sql = "INSERT INTO `LakesWeighHead` (`BagID`, `GrossWeight`, `DateandTime`, `BatchNumb`, `ProductType`) VALUES ('', %s, %s, %s, %s)"
                    val = (number, current_time, user_input, product)
                    mycursor.execute(sql, val)
                    mydb.commit()

                    idsql = "SELECT `BagID` FROM `LakesWeighHead` ORDER BY `BagID` DESC LIMIT 1;"
                    mycursor.execute(idsql)
                    CurrentID = mycursor.fetchone()
                    most_recent_id = CurrentID[0]
                   
                    print(f"\n\nBatch:{user_input}   Weight:{number}    Product:{product}    BagID:{most_recent_id}       Date and time:{current_time}\n")

                    # This line should be commented during testing to prevent having 5000 pages 
                    #print_to_printer(f"\n\nBatch:{user_input}   Weight:{number}    Product:{product}    BagID:{most_recent_id}       Date and time:{current_time}\n")
            
            time.sleep(0.1)  # Small delay to prevent CPU overload

    except KeyboardInterrupt:
        print("Interrupted by user. Exiting to menu.")
        menu(port, baud_rate)
        return  # Exit to menu

    finally:
        ser.close()  # Close the serial connection
def update(port, baud_rate):
    print("Update is used if a bag is broken or incorrectly weighed\n")
    print("Please Enter the Bag ID for the bag that needs reweighing\n")
    while True:
        bag_id = input("--> ")
        try:
            bag_id = int(bag_id)  # Ensure the input is an integer
            break  # Break out of the loop if conversion is successful
        except ValueError:
            print("Invalid input. Please enter an integer.")
    
    ser = serial.Serial(port, baud_rate)
    print("Please place the bag on the scale to get the new weight...")

    try:
        while True:
            if ser.in_waiting > 0:
                data = ser.readline().strip()
                decoded_data = data.decode("utf-8")

                if decoded_data.startswith("Gross"):
                    new_weight = extract_number(decoded_data)
                    if new_weight is not None:
                        # Update the weight in the database
                        sql = "UPDATE `LakesWeighHead` SET `GrossWeight` = %s WHERE `BagID` = %s"
                        val = (new_weight, bag_id)
                        mycursor.execute(sql, val)
                        mydb.commit()
                        print(f"Bag ID {bag_id} updated with new weight {new_weight}.")
                        break  # Break the loop after updating the weight

            time.sleep(0.1)  # Small delay to prevent CPU overload

    except KeyboardInterrupt:
        print("Interrupted by user. Exiting to menu.")
    
    finally:
        ser.close()  # Close the serial connection
        menu(port, baud_rate,ser)  # Return to menu


def previous_day_reports(port, baud_rate):
    today = datetime.now().date()
    days = [(today - timedelta(days=i)) for i in range(1, 21)]  # Last 20 days

    print("Select a day to view the report:")
    for i, day in enumerate(days, 1):
        print(f"{i}) {day}")

    while True:
        user_input = input("Enter the number corresponding to the day (or 0 to exit): ")
        if user_input == '0':
            menu(port, baud_rate)
            return
        try:
            selected_day = int(user_input)
            if 1 <= selected_day <= 20:
                selected_date = days[selected_day - 1]
                break
            else:
                print("Invalid input. Please enter a number between 1 and 20.")
        except ValueError:
            print("Invalid input. Please enter a number.")

    next_day = selected_date + timedelta(days=1)
    
    sql = "SELECT * FROM `LakesWeighHead` WHERE `DateandTime` >= %s AND `DateandTime` < %s"
    val = (selected_date, next_day)
    mycursor.execute(sql, val)

    results = mycursor.fetchall()
    
    if results:
        csv_filename = f"previous_day_report_{selected_date}.csv"  # CSV filename based on selected date
        with open(csv_filename, mode='w', newline='') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(["BagID", "GrossWeight", "BatchNumb", "ProductType", "DateandTime"])  # Write header
            for row in results:
                csv_writer.writerow(row)  
        print(f"Previous day report saved to {csv_filename}.")
        csv_file_path = "C:/Users/Projects.PURNHOUSEFARM/"+ csv_filename
        
       
        print_file_to_word_doc(csv_file_path,selected_date)
        print("File Saved to:",csv_file_path)
        
    else:
        print(f"No entries found for {selected_date}.")

    menu(port, baud_rate,ser) 
def exit_listener():
    global exit_flag
    print("Press '0' to exit.")
    while True:
        if msvcrt.kbhit():  
            input_char = msvcrt.getch().decode('utf-8')
            if input_char == '0':
                exit_flag = True
                break
def menu(port, baud_rate,ser):
    print("")
    print("""\n
             Weigh Head System\n
            1) Start Printing\n
            2) Update a Field\n
            3) Reports\n
            """)
    
    MenuOption = input("-->")
    if MenuOption == "1":
        exit_thread = threading.Thread(target=exit_listener)
        exit_thread.daemon = True
        exit_thread.start()
        read_serial_data(port, baud_rate,ser)
    elif MenuOption == "2":
        update(port, baud_rate)
    elif MenuOption == "3":
        previous_day_reports(port, baud_rate)
    else:
        print("Invalid Input")
        menu(port, baud_rate,ser)  # Return to menu for invalid input

if __name__ == "__main__":
    ser=""
    port = 'COM3'  # Set your COM port
    baud_rate = 9600  # Set the baud rate

    menu(port, baud_rate,ser)
