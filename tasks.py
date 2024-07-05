from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import shutil
import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot into the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    # Create necessary output directories
    create_output_directories()

    # Configure the browser with a delay between actions
    browser.configure(slowmo=200)

    # Perform the necessary steps for ordering robots
    open_robot_order_website()
    download_orders_file()
    process_orders_from_csv()
    archive_receipts()
    clean_up()

def create_output_directories():
    """
    Create directories to store receipts and screenshots.
    """
    os.makedirs("output/receipts", exist_ok=True)
    os.makedirs("output/screenshots", exist_ok=True)

def open_robot_order_website():
    """
    Opens the robot order website and closes any modal that appears.
    """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_modal_if_present()

def close_modal_if_present():
    """
    Closes the modal window if it is present on the page.
    """
    page = browser.page()
    if page.is_visible("text=OK"):
        page.click("text=OK")

def download_orders_file():
    """
    Downloads the orders file from the specified URL.
    """
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def process_orders_from_csv():
    """
    Reads the CSV file and processes each order.
    """
    csv_file = Tables()
    robot_orders = csv_file.read_table_from_csv("orders.csv")
    for order in robot_orders:
        fill_and_submit_order(order)

def fill_and_submit_order(order):
    """
    Fills and submits the robot order form with the data from the order.
    Retries the submission if it fails.
    """
    page = browser.page()

    # Select the robot head
    page.select_option("#head", order["Head"])

    # Select the robot body
    page.click(f"input[name='body'][value='{order['Body']}']")

    # Enter the legs part number
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])

    # Enter the address
    page.fill("#address", order["Address"])

    # Submit the form with retry logic in case of failure
    submit_form_with_retry(order)

def submit_form_with_retry(order):
    """
    Submits the robot order form and retries if submission fails.
    """
    page = browser.page()
    while True:
        page.click("#order")
        if page.is_visible("#order-another"):
            # Process the receipt and screenshot after a successful order
            process_receipt_and_screenshot(order)
            # Click the "order another" button to continue with the next order
            page.click("#order-another")
            close_modal_if_present()
            break

def process_receipt_and_screenshot(order):
    """
    Processes the receipt and screenshot for a given order.
    """
    order_number = int(order["Order number"])
    pdf_path = save_receipt_as_pdf(order_number)
    screenshot_path = take_robot_screenshot(order_number)
    embed_screenshot_in_receipt(screenshot_path, pdf_path)

def save_receipt_as_pdf(order_number):
    """
    Saves the order receipt as a PDF file.
    """
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = f"output/receipts/{order_number}.pdf"
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path

def take_robot_screenshot(order_number):
    """
    Takes a screenshot of the ordered robot.
    """
    page = browser.page()
    screenshot_path = f"output/screenshots/{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_in_receipt(screenshot_path, pdf_path):
    """
    Embeds the screenshot of the robot into the PDF receipt.
    """
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path, source_path=pdf_path, output_path=pdf_path)

def archive_receipts():
    """
    Archives all the receipt PDFs into a single ZIP file.
    """
    archive = Archive()
    archive.archive_folder_with_zip("output/receipts", "output/receipts.zip")

def clean_up():
    """
    Cleans up the directories where receipts and screenshots are saved.
    """
    shutil.rmtree("output/receipts")
    shutil.rmtree("output/screenshots")
