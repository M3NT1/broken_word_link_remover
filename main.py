import os
import logging
import csv
from tkinter import Tk, filedialog
import mammoth
from bs4 import BeautifulSoup
import datetime

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


def extract_hyperlinks_and_bookmarks(doc_path):
    logging.info(f"Extracting hyperlinks and bookmarks from {doc_path}")
    with open(doc_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value

    # Parse the HTML to find hyperlinks and bookmarks
    soup = BeautifulSoup(html, "html.parser")
    hyperlinks = []
    bookmarks = set()

    for a in soup.find_all("a", href=True):
        hyperlink = a["href"]
        text = a.get_text()
        hyperlinks.append((hyperlink, text))
        logging.info(f"Extracted hyperlink: {hyperlink} with text: {text}")

    for bookmark in soup.find_all("a", id=True):
        bookmarks.add(bookmark["id"])
        logging.info(f"Found bookmark: {bookmark['id']}")

    logging.info("Extraction of hyperlinks and bookmarks completed.")
    return hyperlinks, bookmarks


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

    # Extract hyperlinks and bookmarks
    hyperlinks, bookmarks = extract_hyperlinks_and_bookmarks(doc_path)

    links_array = []
    for hyperlink, text in hyperlinks:
        if hyperlink.startswith("#"):
            if hyperlink[1:] in bookmarks:
                link_status = "Belső hivatkozás (kereszthivatkozás)"
            else:
                link_status = "Belső hivatkozás (szellem hivatkozás)"
        else:
            link_status = "Külső hivatkozás"

        links_array.append([
            hyperlink,
            text,
            link_status,
            "N/A",  # Page number extraction is not directly supported
            "NEM"
        ])
        logging.info(f"Processed hyperlink: {hyperlink}, Status: {link_status}, Text: {text}")

    csv_path = os.path.join(output_folder, "Frissített_Hivatkozások.csv")
    save_csv(links_array, csv_path)

    log_message(log_file, "Dokumentum mentve")
    log_message(log_file, "Naplózás vége: " + str(datetime.datetime.now()))
    logging.info("Link extraction and management process completed.")


if __name__ == "__main__":
    list_and_manage_links()
