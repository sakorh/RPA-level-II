from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    browser.configure(
        slowmo=100,
    )
    open_the_robot_order_website()
    orders = get_orders()
    for order in orders:
        close_annoying_modal()
        fill_and_submit_order(order)
        embed_screenshot_to_receipt([f"output/screenshots/receipt_{order['Order number']}.png"], f"output/receipts/receipt_{order['Order number']}.pdf")
    archive_receipts()            


def open_the_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """
    Downloads the order file from the given URL.
    Reads it as a table and returns the result.
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    return orders

def close_annoying_modal():
    """Closes the modal popup"""
    page = browser.page()
    page.click('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[3]')

def fill_and_submit_order(order):
    """
    Fills the order form for each robot order.
    Previews the robot.
    Submits the order.
    """
    page = browser.page()

    page.select_option('//*[@id="head"]', order['Head'])
    page.check(f'#id-body-{order["Body"]}.form-check-input')
    page.fill('//*[@placeholder="Enter the part number for the legs"]', order['Legs'])
    page.fill('//*[@id="address"]', order['Address'])
    page.click('//*[@id="preview"]')
    page.click('//*[@id="order"]')
    while True:
        if page.is_visible('//*[@id="order-another"]'):
            store_receipt_as_pdf(order['Order number'])
            screenshot_robot(order['Order number'])
            page.click('//*[@id="order-another"]')
            break
        else:
            page.click('//*[@id="order"]')

    
def store_receipt_as_pdf(order_number):
    """Exports the receipt to a pdf file"""
    page = browser.page()
    receipt_html = page.locator('//*[@id="receipt"]').inner_html()
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, f"output/receipts/receipt_{order_number}.pdf")

def screenshot_robot(order_number):
    """Takes a screenshot of the page"""
    page = browser.page()
    page.locator('#robot-preview-image').screenshot(path=f"output/screenshots/receipt_{order_number}.png", omit_background=True)

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the screenshot of the robot to the PDF receipt"""
    pdf = PDF()
    pdf.add_files_to_pdf(files=screenshot, target_document=pdf_file, append=True)

def archive_receipts():
    """Archives all receipts into a single zip file"""
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'receipts.zip')