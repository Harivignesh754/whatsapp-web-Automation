import gspread
from google.oauth2 import service_account
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
import logging
from datetime import datetime

# Configure logging to write logs to both console and file
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filename='log.txt')

# Google Sheets API credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = service_account.Credentials.from_service_account_file('credentials.json', scopes=scope)
client = gspread.authorize(creds)

def send_whatsapp_message(driver, contact, message, worksheet, row_index):
    try:
        logging.info(f"Sending message to {contact}")

        # Check if the update column already contains "Message Sent"
        status = worksheet.cell(row_index, 15).value
        if status == "Message Sent":
            logging.info(f"Message already sent for {contact} at {worksheet.cell(row_index, 3).value}")
            return

        # Find and send message
        search_box_locator = (By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/div/div[1]/div/div[2]/div[2]/div/div[1]/p')
        search_box = WebDriverWait(driver, 30).until(EC.presence_of_element_located(search_box_locator))
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.DELETE)
        search_box.send_keys(contact)
        time.sleep(5)  # Wait for search results
        search_box.send_keys(Keys.ENTER)

        message_box_locator = (By.XPATH, '/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')
        message_box = WebDriverWait(driver, 30).until(EC.presence_of_element_located(message_box_locator))
        
        # Send message with newline characters preserved
        for line in message.split('\n'):
            message_box.send_keys(line)
            message_box.send_keys(Keys.SHIFT, Keys.ENTER)  # Use SHIFT + ENTER to create a newline
        message_box.send_keys(Keys.ENTER)  # Send the message
        time.sleep(5)

        # Update Google Sheets with "Message Sent" status and current time
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        worksheet.update_cell(row_index, 15, f"Message Sent at {now}")
        logging.info(f"Message sent successfully to {contact} at {now}")

    except NoSuchElementException as e:
        logging.error(f"Element not found while sending message to {contact}: {e}")
    except TimeoutException as e:
        logging.error(f"Timeout occurred while sending message to {contact}: {e}")
    except Exception as e:
        logging.error(f"Error sending message to {contact}: {e}", exc_info=True)  # Log exception details

def main():
    try:
        # Open the Google Sheets document
        spreadsheet = client.open('MarlenMahalQuoteGenerator (Responses)')
        worksheet = spreadsheet.get_worksheet(0)  # Get the First worksheet

        # Ask user to enter from-date and to-date
        from_date_str = input("Enter from-date (DD/MM/YYYY): ").strip()
        to_date_str = input("Enter to-date (DD/MM/YYYY): ").strip()

        # Convert input strings to datetime objects
        from_date = datetime.strptime(from_date_str, "%d/%m/%Y")
        to_date = datetime.strptime(to_date_str, "%d/%m/%Y")

        # Read message from message.txt file
        with open('message.txt', 'r') as file:
            message = file.read()

        # Initialize Chrome WebDriver
        driver = webdriver.Chrome()

        # Open WhatsApp Web
        driver.get("https://web.whatsapp.com/")
        logging.info("Scan the QR code within the next 20 seconds.")
        time.sleep(20)

        # Get data from Google Sheets and send messages
        data = worksheet.get_all_values()
        for i, row in enumerate(data[1:], start=2):  # Start iterating from the second row
            contact = row[3]  # Assuming contact number is in the Fourth column
            date_str = row[1]  # Assuming follow-up date is in the second column

            # Convert date string to datetime object
            if date_str:  # Check if date_str is not empty
                date = datetime.strptime(date_str, "%d/%m/%Y")
            else:
                # Handle the case when date_str is empty
                logging.warning("Encountered an empty date string.")
                continue  # Skip this iteration

            # Send message if the date falls within the specified range
            if from_date <= date <= to_date:
                send_whatsapp_message(driver, contact, message, worksheet, i)

    except gspread.exceptions.SpreadsheetNotFound as e:
        logging.error("Spreadsheet not found:", e)
    except Exception as e:
        logging.error("An unexpected error occurred:", e)

    finally:
        # Close the browser after sending all messages
        if 'driver' in locals():
            driver.quit()
            logging.info("WebDriver quit")

if __name__ == "__main__":
    main()
