import time
import pandas as pd
import pyautogui
import keyboard as k
import webbrowser as web
from urllib.parse import quote

def send_whatsapp(data_file_excel, message_file_text, x_cord=628, y_cord=869):
    # Reading the Excel file and assigning the contents to a DataFrame
    df = pd.read_excel(data_file_excel, dtype={"Contact": str})
    
    # Extracting the values in the "Name" and "Contact" columns
    name = df['Name'].values
    contact = df['Contact'].values
    
    # Reading the contents of the message file
    with open(message_file_text) as f:
        file_data = f.read()
    
    # Zipping together the values in name and contact into tuples
    zipped = zip(name, contact)
    
    # Initializing a counter variable to keep track of the number of messages sent
    counter = 0
    
    # Looping over each tuple in zipped
    for (a, b) in zipped:
        # Formatting the message text with the name of the recipient
        msg = file_data.format(a)
        
        # Opening the WhatsApp Web URL for the corresponding contact and message text
        web.open(f"https://web.whatsapp.com/send?phone={b}&text={quote(msg)}")
        
        # Adding a delay to allow the WhatsApp Web page to load
        time.sleep(20)  # Increased wait time to ensure WhatsApp Web is fully loaded
        
        # Simulating a mouse click at the specified coordinates to send the message
        pyautogui.click(x_cord, y_cord)
        
        # Adding a delay to allow time for the message to be sent
        time.sleep(5)  # Increased time to make sure the message is sent
        
        # Simulating the pressing of the "Return" key (instead of Enter)
        k.press_and_release('return')
        
        # Adding a delay to allow time for the pop-up to be dismissed
        time.sleep(2)
        
        # Simulating the pressing of the "Command + W" keys to close the WhatsApp Web tab (Mac shortcut)
        pyautogui.hotkey('command', 'w')  # Use the Mac shortcut for closing the tab
        
        # Adding a delay to allow time for the tab to be closed
        time.sleep(2)  # Ensuring the tab closes completely before continuing
        
        # Incrementing the counter variable and printing a message indicating that the message has been sent
        counter += 1
        print(counter, "- Message sent..!!")
    
    # Printing a message indicating that the function has completed execution
    print("Done!")

# Defining the paths to the Excel file and message text file to be used as inputs to the send_whatsapp function
excel_path = r"/Users/adityabhati/Desktop/RC mesages/0. Whatsapp Web Automation/Whatsapp List_Main.xlsx"
text_path = r"/Users/adityabhati/Desktop/RC mesages/0. Whatsapp Web Automation/WHATSDRAFT.txt"

send_whatsapp(excel_path, text_path)
