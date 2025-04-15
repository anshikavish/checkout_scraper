from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import sys
import datetime
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException


def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = r"json_file"
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=credentials)
    return service


def get_sheet(service, spreadsheet_id):
    # Get the spreadsheet data
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return spreadsheet


def update_sheet(service, spreadsheet_id, error_message, step, website):
    # Get the sheet by ID
    sheet = get_sheet(service, spreadsheet_id)
    # Get the first empty row in the 'Sheet1'
    sheet_name = 'sheet_name'  # Targeting the specific sheet name
    range_ = f'{sheet_name}!A:A'  # Range to check for the last row in Column A
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_).execute()
    values = result.get('values', [])
    # Find the next empty row
    next_row = len(values) + 1
    # Prepare data to be updated
    current_date = datetime.datetime.now().strftime(r"%d-%m-%Y %H:%M:%S")  # Current date in the desired format
    print(current_date)
    # Ensure the data is properly formatted
    values = [
        [current_date, error_message, step, website]  # [Date, Error, Step]
    ]
    # Now define the range to update the sheet with the data
    range_ = f'{sheet_name}!A{next_row}:D{next_row}'  # Update the range to columns A, B, C
    body = {
        'values': values
    }
    # Call the API to update the sheet


    try:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_,
            valueInputOption="RAW", body=body).execute()
        print("Data successfully updated in Google Sheets.")
    except HttpError as err:
        print(f"Error: {err}")


def log_error_to_sheet(error_message, step, website):
    service = authenticate_google_sheets()
    spreadsheet_id = 'sheet_id'  # Your Google Sheets ID
    update_sheet(service, spreadsheet_id, error_message, step, website)

def update_sheet_success(service, spreadsheet_id, website, order_id):
    sheet = get_sheet(service, spreadsheet_id)
    sheet_name = 'sheet_name'
    range_ = f'{sheet_name}!A:A'  # Range to check for the last row in Column A
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_).execute()
    values = result.get('values', [])
    # Find the next empty row
    next_row = len(values) + 1
    # Prepare data to be updated
    current_date = datetime.datetime.now().strftime(r"%d-%m-%Y %H:%M:%S")  # Current date in the desired format
    print(current_date)
    # Ensure the data is properly formatted
    values = [
        [current_date, website, order_id]  # [Date, Error, Step]
    ]
    # Now define the range to update the sheet with the data
    range_ = f'{sheet_name}!A{next_row}:C{next_row}'  # Update the range to columns A, B, C
    body = {
        'values': values
    }
    # Call the API to update the sheet


    try:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_,
            valueInputOption="RAW", body=body).execute()
        print("Data successfully updated in Google Sheets.")
    except HttpError as err:
        print(f"Error: {err}")

def order_id_to_sheet(website, order_id):
    service = authenticate_google_sheets()
    spreadsheet_id = 'sheet_id'  # Your Google Sheets ID
    update_sheet_success(service, spreadsheet_id, website, order_id)    


import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# Define email credentials
SENDER_EMAIL = "email_id"
RECEIVER_EMAIL = ["email_id"]
EMAIL_PASSWORD = "password"  #Use environment variables in production.
# Global variable to track last error message sent


def send_error_email(step):
    """Send an email with only the error step, without traceback."""
    try:
        # Setup email server
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()  # Secure connection
        server.login(SENDER_EMAIL, EMAIL_PASSWORD)
        # Email content setup (Without traceback)
        subject = f"subject"
        body = f"""An error occurred in the script.

Step : {step}

"""
        # Create and send the email
        message = MIMEMultipart()
        message["From"] = SENDER_EMAIL
        message["To"] =  ", ".join(RECEIVER_EMAIL)
        message["Subject"] = subject
        message.attach(MIMEText(body, "plain"))
        server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, message.as_string())
        print(f"Error email sent successfully")
    except Exception as email_exception:
        print(f"Failed to send error email: {email_exception}")
    finally:
        server.quit()  # Ensure the server connection is closed


# Initialize the WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
website = "BVO"
# Open the webpage


def open_website(url, step):
    try:
        # Step 1: Open the website
        driver.get(url)
        driver.maximize_window()
        print(f"Website opened: {url}")
    except Exception as e:
        print(f"Error occurred while {step}: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)
        driver.quit()
        sys.exit()
    # Step 2: Close modal if it exists
    try:
        close_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'close') or @aria-label='Close']"))
        )
        close_button.click()
        print("Modal closed")
    except Exception as e:
        print("No modal appeared or modal close failed:", e)
    # Step 3: Click 'No Thanks' button if it exists
    try:
        nothanks = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@id, 'secondaryButton')]"))
        )
        nothanks.click()
        print("I'll do this later button clicked")
    except Exception as e:
        print("No Thanks button not found")
        pass
    # Step 4: Wait for the page to load
    time.sleep(3)
    print("Page is loaded and ready.")


try:
    # Replace with actual driver initialization
    # Placeholder for the WebDriver object
    step = "opening website"
    url = "url"
    # Call the function
    open_website(url, step)
except Exception as e:
    print(f"Error in overall process: {e}")
    if 'driver' in locals() and driver:  # Ensure the driver is initialized before quitting
        driver.quit()


# Step 0: open bestseller page
def navigate_to_bestsellers(step):
    try:
        # Step 1: Locate the 'Bestsellers' link
        bestsellers_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[.//span[text()='Bestsellers']]"))
        )
        # Step 2: Extract the href attribute
        href = bestsellers_link.get_attribute("href")
        print("Bestsellers link opened")
        # Step 3: Navigate directly to the Bestsellers page
        driver.get(href)
        print("Navigated directly to the Bestsellers page.")
    except Exception as e:
        print(f"Error occurred while navigating to the Bestsellers page: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message,step,website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    # Replace with actual WebDriver instance
    step = "Opening Bestseller page"
    # Call the function to navigate to the Bestsellers page
    navigate_to_bestsellers(step)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 1: Select the first product
def click_on_first_product(step):
    try:
        # Step 1: Locate the first product
        first_product = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'card__media')]"))
        )
        # Step 2: Scroll into view and click on the product
        driver.execute_script("arguments[0].scrollIntoView(true);", first_product)
        ActionChains(driver).move_to_element(first_product).click().perform()
        print("Clicked on the first product PDP")
    except Exception as e:
        print(f"Error occurred while selecting the first product: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()
    # Step 3: Wait for the product detail page to load
    time.sleep(3)
    print("Product Detail Page (PDP) loaded.")


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    # Replace with actual WebDriver instance
    step = "At the first product"
    # Call the function to click on the first product
    click_on_first_product(step)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 2: Click the "Add to Cart" button
def click_add_to_cart(step):
    try:
        # Attempt to locate the first 'Add to Cart' button
        print("Attempting to locate the 'Add to Cart' button...")
        try:
            button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'product-form__submit')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            ActionChains(driver).move_to_element(button).click().perform()
            print("Add to Cart button clicked.")
        # If the first button isn't found, attempt to locate the second one
            time.sleep(10)
            try:
            # Check if the mini-cart navigation is visible
                mini_cart = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "ul.mini-cart__navigation"))
                )
                print("Mini cart is visible, proceeding to next step.")
            except:
                # If mini-cart is not visible, click the 'Add to Cart' button again
                print("Mini cart not visible, clicking 'Add to Cart' again.")
                button.click()  # Try clicking again
                time.sleep(5)  # Allow time for the page to update again
            
                # Check again if the mini-cart is visible
                try:
                    mini_cart = WebDriverWait(driver, 5).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, "ul.mini-cart__navigation"))
                    )
                    print("Mini cart is visible after second click, proceeding to next step.")
                except:
                    print("Mini cart still not visible. Check if there's an issue with the page or the button.")

        except:
            print("Attempting to locate the alternative 'Add to Cart' button...")
            button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='product-form-template--15143001129006__main--alt']/div/button"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            ActionChains(driver).move_to_element(button).click().perform()
            print("Add to Cart button clicked.")
            time.sleep(10)
    except Exception as e:
        print(f"Error occurred while clicking the 'Add to Cart' button: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    step = "Click the add to cart button"
    # Call the function to click the Add to Cart button
    click_add_to_cart(step)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 3: Click the "Checkout" button
def click_checkout_button(step):
    try:
        # Attempt to locate the 'Checkout' button
        print("Attempting to locate the checkout button...")
        checkout_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'zecpe-btn')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", checkout_button)
        ActionChains(driver).move_to_element(checkout_button).click().perform()
        time.sleep(5)
        print("Checkout button clicked.")
    except Exception as e:
        print(f"Error occurred while clicking the 'Checkout' button: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    step = "Click the 'Checkout' button"
    # Call the function to click the Checkout button
    click_checkout_button(step)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 4: Enter the mobile number
def enter_mobile_number_and_get_otp(step, mobile_number):
    try:
        # Step 1: Locate the mobile number input field and enter the number
        print("Attempting to enter mobile number...")
        mobile_number_input = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[contains(@class, 'zecpe-enter-phone__input')]"))
        )
        mobile_number_input.send_keys(mobile_number)  # Replace with provided mobile number
        print("Mobile number entered.")
        # Step 2: Locate the 'Get OTP' button and click on it
        otp_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'zecpe-button')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", otp_button)
        ActionChains(driver).move_to_element(otp_button).click().perform()
        print("Clicked on 'Get OTP' button.")
    except Exception as e:
        print(f"Error occurred while entering mobile number or clicking 'Get OTP' button: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    step = "Enter the mobile number"
    mobile_number = "mobile_number"  # Replace with the mobile number you want to enter
    # Call the function to enter the mobile number and get OTP
    enter_mobile_number_and_get_otp(step, mobile_number)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 5: Wait until OTP is filled
def enter_otp(step, otp):
    try:
        # Step 1: Wait for the OTP input field to be present and interactable
        print("Waiting for OTP input field to be present...")
        otp_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "zecpe-otp-input-0"))
        )
        # Step 2: Enter the OTP
        otp_field.send_keys(otp)
        time.sleep(1)
        print(f"OTP entered: {otp}. Proceeding to the next step.")
        time.sleep(5)
    except Exception as e:
        print(f"Error occurred while entering OTP: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()
    # try:
    #     print("Waiting for OTP to be filled...")
    #     WebDriverWait(driver, 300).until(
    #     lambda driver: driver.find_element(By.ID, "zecpe-otp-input-3").get_attribute("value").strip() != ""
    # )
    #     print("OTP filled. Proceeding to the next step.")
    # except Exception as e:
    #     print("Error occurred while waiting for OTP to be filled:", e)
    #     driver.quit()
    #     exit()
    # time.sleep(10)


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    step = "OTP entered"
    otp = "otp"  # Replace with the OTP you want to enter
    # Call the function to enter OTP
    enter_otp(step, otp)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 6: Click on shipping address continue button
def click_shipping_address_continue_button(step):
    try:
        print("Attempting to locate and click the 'Shipping Address Continue' button...")
        try:
            # First attempt: Locate the button using WebDriverWait
            continue_button = WebDriverWait(driver, 50).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@class='zecpe-button']"))
            )
            # Scroll into view to ensure visibility
            driver.execute_script("arguments[0].scrollIntoView(true);", continue_button)
            # Click the button using ActionChains
            actions = ActionChains(driver)
            actions.move_to_element(continue_button).click().perform()
            print("'Shipping Address Continue' button clicked successfully.")
        except Exception:
            print("Attempting to locate and click the 'Shipping Address Continue' button...by 2nd way")
            # Second attempt: Locate the button using WebDriverWait with different XPath
            continue_button = WebDriverWait(driver, 50).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='zecpe-contents']/div[3]/div[2]/div/button/span"))
            )
            # Scroll into view to ensure visibility
            driver.execute_script("arguments[0].scrollIntoView(true);", continue_button)
            # Click the button using ActionChains
            actions = ActionChains(driver)
            actions.move_to_element(continue_button).click().perform()
            print("'Shipping Address Continue' button clicked successfully.")
    except Exception as e:
        print(f"An unexpected error occurred while interacting with the 'Continue' button: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)
        driver.quit()
        sys.exit()


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    step = "Click on shipping address continue button"
    # Call the function to click the 'Shipping Address Continue' button
    click_shipping_address_continue_button(step)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 7: Click on payment method (Cash On Delivery)
def select_payment_method_cod(step):
    try:
        print("Attempting to locate and click the 'Cash on Delivery' payment method...")
        # Locate the "Cash on Delivery" payment method using its unique data attribute
        cod_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.XPATH, "//div[@class='zecpe-payment-method__header' and @data-zecpe-payment-method='Cash on Delivery']"
            ))
        )
        # Scroll into view to ensure visibility
        driver.execute_script("arguments[0].scrollIntoView(true);", cod_option)
        # Click the payment option
        cod_option.click()
        print("Cash on Delivery payment method selected.")
    except Exception as e:
        print(f"An error occurred: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    step = "Click on payment method (CASH ON DELIVERY)"
    # Call the function to select the 'Cash on Delivery' payment method
    select_payment_method_cod(step)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 8: Click on payment method (Cash On Delivery)
def click_pay_on_delivery(step):
    try:
        print("Attempting to locate and click the 'Pay on Delivery' option...")
        # Locate the "Cash on Delivery" option by its ID
        cod_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "zecpe-payment-screen__cod-btn"))
        )
        # Scroll into view to ensure visibility
        driver.execute_script("arguments[0].scrollIntoView(true);", cod_option)
        # Click the payment option
        cod_option.click()
        print("Clicked on 'Pay on Delivery'")
    except Exception as e:
        print(f"Error occurred while clicking Pay on Delivery: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    step = "Click on Pay on Delivery"
    # Call the function to click the 'Pay on Delivery' button
    click_pay_on_delivery(step)
except Exception as e:
    print(f"Error in overall process: {e}")


# Step 9: Click on confirm order
def click_confirm_order(step):
    try:
        print("Attempting to locate and click the 'Confirm Order' button...")
        # Wait for the 'Confirm Order' button to be clickable
        confirm_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "zecpe-cod-confirmation-modal__btn--confirm"))
        )
        # Scroll into view (optional, if the button is not in the visible viewport)
        driver.execute_script("arguments[0].scrollIntoView(true);", confirm_button)
        # Click the button
        confirm_button.click()
        print("Confirm button clicked successfully.")
        # Optional wait after the click (you can adjust this as needed)
        time.sleep(40)
    except Exception as e:
        print(f"Error occurred while clicking the confirm button: {e}")
        error_message = str(e)
        log_error_to_sheet(error_message, step, website)
        send_error_email(step)  # Assuming you have a function to send the error email
        driver.quit()
        sys.exit()


try:
    # Initialize driver (This is a placeholder for actual driver initialization)
    step = "Click on Confirm Order"
    # Call the function to click the 'Confirm Order' button
    click_confirm_order(step)
except Exception as e:
    print(f"Error in overall process: {e}")


try:
    order_number_element = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'os-order-number'))
    )
    # Extract the text
    order_number_text = order_number_element.text
    order_number = order_number_text.replace('Order', '').strip()
    print(f"{order_number}")
    order_id_to_sheet(website, order_number)
except Exception as e:
    print("not found", e)


import smtplib
# Create an Email object to store sender and receiver details
class Email:
    def __init__(self, sender, receiver, password):
        self.sender = sender
        self.receiver = receiver
        self.password = password
# Define email credentials using the Email object
email_config = Email(
    sender="email_id",
    receiver=["email_id"],
    password="password"  #Never hardcode passwords in production
)
# Secure the connection to the email server
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls() 


try:
    # Log in to the server
    server.login(email_config.sender, email_config.password)
    # Construct the email message
    subject = "BVO : Order Cancellation Request"
    body = f'''Order successfully

Please cancel this order : {order_number}

Kindly confirm once the cancellation has been processed.

    '''
    #body = f"please delete this order: {order_number_text}"
    email_message = f"Subject: {subject}\n\n{body}"
    # Send the email
    server.sendmail(email_config.sender, email_config.receiver, email_message)
    print("Mail sent successfully!")
except Exception as e:
    print(f"Error sending email: {e}")
finally:
    # Close the connection to the server
    server.quit()
# Step 7: Close the browser
driver.quit()
print("Browser closed.")