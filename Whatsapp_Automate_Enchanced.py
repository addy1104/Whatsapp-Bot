import keyboard as k    # for simulating keyboard keys
import pyautogui        # for automating mouse and keyboard actions
import time             # for adding delay in program execution
import pandas as pd     # for reading and manipulating Excel files
import webbrowser as web # for opening URLs in web browser
from urllib.parse import quote # for URL encoding special characters

def send_whatsapp_fast(data_file_excel, message_file_text, delay=8, batch_size=3):
    df = pd.read_excel(data_file_excel, dtype={"Contact": str})
    name = df['Name'].values
    contact = df['Contact'].values
    
    with open(message_file_text) as f:
        file_data = f.read()
    
    zipped = zip(name, contact)
    counter = 0
    batch_count = 0

    for (a, b) in zipped:
        msg = file_data.format(a)
        
        # Open WhatsApp web with the given contact and message
        web.open(f"https://web.whatsapp.com/send?phone={b}&text={quote(msg)}")
        
        # Reduce loading time
        time.sleep(delay)
        
        # Press "Enter" key to send message
        pyautogui.press("enter")
        
        # Small delay before closing tab
        time.sleep(1.5)
        
        # Close the current tab (Mac shortcut)
        pyautogui.hotkey('command', 'w')

        counter += 1
        batch_count += 1
        print(counter, "- Message sent..!!")

        # Open multiple messages in batches to improve speed
        if batch_count >= batch_size:
            print(f"Waiting {delay} seconds before opening the next batch...")
            time.sleep(delay)
            batch_count = 0  # Reset batch counter

    print("Done!")

# Paths to Excel file and message text file
excel_path = r"/Users/adityabhati/Desktop/RC mesages/0. Whatsapp Web Automation/Whatsapp List_Main.xlsx"
text_path = r"/Users/adityabhati/Desktop/RC mesages/0. Whatsapp Web Automation/WHATSDRAFT.txt"

send_whatsapp_fast(excel_path, text_path)
