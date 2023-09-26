from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
import os
import zipfile


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    
    """
    open_robot_order_website()
    orders = get_orders()
    """browser.configure(
        slowmo=100,
    )"""
    page = browser.page()
    for order in orders:
        close_annoying_modal()
        fill_the_form(order)
        orderNumber = page.locator('//p[@class="badge badge-success"]').inner_text()
        store_receipt_as_pdf(orderNumber)
        screenshot_robot(orderNumber)
        embed_screenshot_to_receipt(orderNumber+".png",orderNumber+".pdf")
        page.click("#order-another")
    archive_receipts()
    
def archive_receipts():
    """Archive pdf receipts as a ZIP"""
    receipts_folder_path = 'output/receipts/'
    zip_file_name = 'output/pdf_receipts.zip'
   
    with zipfile.ZipFile(zip_file_name, 'w') as pdf_receipts_zip:    
        for foldername, subfolders, filenames in os.walk(receipts_folder_path):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)               
                if filename.lower().endswith('.pdf') and 'temp' not in foldername:
                    pdf_receipts_zip.write(file_path, filename)

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """merge screenshot to pdf"""
    pdf=PDF()
    pdf.add_files_to_pdf(
        files=["output/receipts/screenshots/"+screenshot,"output/receipts/temp/"+pdf_file],
        target_document="output/receipts/"+pdf_file
    )

def screenshot_robot(order_number):
    """take a screenshot from receipt"""
    page = browser.page()
    page.screenshot(path='output/receipts/screenshots/'+order_number+'.png')

def store_receipt_as_pdf(order_number):
    """save html receipt as PDF"""
    pdf = PDF()
    sales_results_html = browser.page().locator("#receipt").inner_html()
    pdf.html_to_pdf(sales_results_html, 'output/receipts/temp/'+order_number+'.pdf')


def fill_the_form(csv):
    """fill the Form"""
    page = browser.page()
    page.select_option("#head", csv["Head"])
    page.click("#id-body-" + csv["Body"])
    page.fill('//input[@placeholder="Enter the part number for the legs"]', csv["Legs"])
    page.fill('//input[@placeholder="Shipping address"]', csv["Address"])
    while(True):
        page.click("#order")
        error = page.locator('//div[@class="alert alert-danger"]').is_visible()
        if error:
            continue
        else:
            break

def close_annoying_modal():
    """Close web page popup"""
    page = browser.page()
    page.click("text=Yep")

def open_robot_order_website():
    """Open robot sparebin industries website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Download file and return it as csv"""
    download_csv_file()
    library = Tables()

    ordersCsv = library.read_table_from_csv(
    "orders.csv", columns=["Order Number", "Head", "Body", "Legs", "Address"]
    )
    return ordersCsv

