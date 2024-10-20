import os
import logging
import csv
from tkinter import Tk, filedialog, simpledialog
import mammoth
from bs4 import BeautifulSoup
import datetime
from docx2pdf import convert
from PyPDF2 import PdfReader
from collections import defaultdict

# Logging setup
logging.basicConfig(filename='link_manager.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def select_folder():
    root = Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Válassza ki a kimeneti mappát")
    return folder_path


def initialize_log_file(output_folder):
    log_file_path = os.path.join(output_folder,
                                 f"HivatkozasKezelo_Naplo_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    logging.info("Naplózás kezdete: " + str(datetime.datetime.now()))
    return log_file_path


def log_message(log_file, message):
    logging.info(message)
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"{datetime.datetime.now()} - {message}\n")


def convert_to_pdf(doc_path, output_folder):
    pdf_path = os.path.join(output_folder, "converted_document.pdf")
    try:
        convert(doc_path, pdf_path)
        logging.info(f"Document converted to PDF: {pdf_path}")
    except Exception as e:
        logging.error(f"Error converting to PDF: {e}")
        raise
    return pdf_path


def extract_hyperlinks_and_bookmarks(doc_path):
    logging.info(f"Extracting hyperlinks and bookmarks from {doc_path}")
    with open(doc_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value

    soup = BeautifulSoup(html, "html.parser")
    hyperlinks = []
    bookmarks = set()
    order = 0

    for a in soup.find_all("a", href=True):
        hyperlink = a["href"]
        text = a.get_text()
        hyperlinks.append((hyperlink, text, order))
        order += 1
        logging.info(f"Extracted hyperlink: {hyperlink} with text: {text}")

    for bookmark in soup.find_all("a", id=True):
        bookmarks.add(bookmark["id"])
        logging.info(f"Found bookmark: {bookmark['id']}")

    logging.info("Extraction of hyperlinks and bookmarks completed.")
    return hyperlinks, bookmarks


def extract_page_numbers(pdf_path, hyperlinks):
    logging.info(f"Extracting page numbers from {pdf_path}")
    reader = PdfReader(pdf_path)
    link_pages = defaultdict(list)

    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text()
        for hyperlink, link_text, order in hyperlinks:
            if link_text in text:
                link_pages[hyperlink].append((page_number, order))
                logging.info(f"Hyperlink: {hyperlink} found on page {page_number}")

    return link_pages


def save_csv(data, file_path):
    logging.info(f"Saving CSV to {file_path}")
    with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(["Cél", "Link Szöveg", "Státusz", "Oldalszám", "Módosításra került?"])
        csvwriter.writerows(data)
    logging.info(f"CSV file successfully saved: {file_path}")


def list_and_manage_links():
    logging.info("Starting link extraction and management process.")
    doc_path = filedialog.askopenfilename(title="Válassza ki a Word dokumentumot",
                                          filetypes=[("Word dokumentumok", "*.docx")])
    if not doc_path:
        logging.warning("No file selected. Exiting.")
        print("Nem választott ki fájlt. A program leáll.")
        return

    output_folder = select_folder()
    if not output_folder:
        logging.warning("No output folder selected. Exiting.")
        print("Nem választott ki kimeneti mappát. A program leáll.")
        return

    log_file = initialize_log_file(output_folder)
    log_message(log_file, "Dokumentum feldolgozásának kezdete")

    # Convert to PDF first
    pdf_path = convert_to_pdf(doc_path, output_folder)

    # Ask for starting page number
    root = Tk()
    root.withdraw()
    start_page = simpledialog.askinteger("Kezdő oldalszám", "Adja meg a kezdő oldalszámot:", initialvalue=1, minvalue=1)
    if not start_page:
        start_page = 1

    # Extract hyperlinks and bookmarks
    hyperlinks, bookmarks = extract_hyperlinks_and_bookmarks(doc_path)

    # Extract page numbers for all pages
    link_pages = extract_page_numbers(pdf_path, hyperlinks)

    links_array = []
    for hyperlink, text, order in hyperlinks:
        if hyperlink.startswith("#"):
            if hyperlink[1:] in bookmarks:
                link_status = "Belső hivatkozás (kereszthivatkozás)"
            else:
                link_status = "Belső hivatkozás (szellem hivatkozás)"
        else:
            link_status = "Külső hivatkozás"

        page_numbers = link_pages.get(hyperlink, [])
        for page_number, original_order in page_numbers:
            if page_number >= start_page:
                links_array.append([
                    hyperlink,
                    text,
                    link_status,
                    page_number,
                    "NEM",
                    original_order
                ])
                logging.info(
                    f"Processed hyperlink: {hyperlink}, Status: {link_status}, Text: {text}, Page: {page_number}")

    # Sort links_array by page number, then by original order
    links_array.sort(key=lambda x: (x[3], x[5]))

    # Remove duplicates while preserving order
    seen = set()
    unique_links_array = []
    for link in links_array:
        key = (link[0], link[1], link[3])  # Cél, Link Szöveg, Oldalszám
        if key not in seen:
            seen.add(key)
            unique_links_array.append(link[:-1])  # Remove the temporary order column

    csv_path = os.path.join(output_folder, "Frissített_Hivatkozások.csv")
    save_csv(unique_links_array, csv_path)

    log_message(log_file, "Dokumentum mentve")
    log_message(log_file, "Naplózás vége: " + str(datetime.datetime.now()))
    logging.info("Link extraction and management process completed.")


if __name__ == "__main__":
    list_and_manage_links()
