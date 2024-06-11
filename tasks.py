from datetime import datetime

from robocorp import browser
from robocorp.tasks import task
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables


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
        slowmo=200,
    )
    open_robot_order_website()
    close_annoying_modal()
    fill_the_form()


def open_robot_order_website():
    """Open a browser and download the orders file"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv")
    return orders


def close_annoying_modal():
    page = browser.page()
    page.click("button:text('Yep')")


def fill_the_form():
    page = browser.page()
    orders = get_orders()
    fs = FileSystem()

    directory = "output/receipts_" + datetime.now().strftime("%d.%m.%Y") + "_"
    fs.create_directory(directory)
    fs.create_directory("output/screenshots")

    for row in orders:

        order_number = row["Order number"]
        receipt_file = directory + "/" + order_number + ".pdf"

        page.select_option("#head", (row["Head"]))
        page.click("input[id='id-body-" + str(row["Body"]) + "']")
        page.fill('input[placeholder="Enter the part number for the legs"]', (row["Legs"]))
        page.fill("#address", (row["Address"]))
        page.click("#preview")

        for _ in range(5):
            try:

                page.click("#order")
                store_receipt_as_pdf(order_number, receipt_file, directory)
                break

            except Exception as OrderError:
                print("An error occurred: {OrderError}. Retrying...")
            else:
                print("Max retries exceeded.")

        page.click("#order-another")
        close_annoying_modal()
    archive_receipts(directory)


def store_receipt_as_pdf(order_number, receipt_file, directory):
    page = browser.page()
    receipt_text = page.locator("#order-completion").inner_html(timeout=1000)
    screenshotlist = screenshot_robot(order_number)

    pdf = PDF()
    pdf.html_to_pdf(receipt_text, directory + "/" + order_number + ".pdf")
    pdf.add_files_to_pdf(files=screenshotlist, target_document=receipt_file, append=True)

    return order_number


def screenshot_robot(order_number):
    screenshot = "output/screenshots/screenshot_" + order_number + ".png"
    screenshotlist = [screenshot]
    page = browser.page()
    page.screenshot(path=screenshot)
    return screenshotlist


def archive_receipts(directory):
    archive = Archive()
    archive.archive_folder_with_zip(
        folder=directory, archive_name="output/receipts_" + datetime.now().strftime("%d.%m.%Y") + ".zip"
    )
