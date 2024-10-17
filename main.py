import os
import datetime
import logging
from tkinter import Tk, filedialog
import docxpy
import csv
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

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
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"{datetime.datetime.now()} - {message}\n")


def extract_hyperlinks_and_bookmarks(doc_path):
    # Open the docx file as a zip archive
    with ZipFile(doc_path, 'r') as zipf:
        # Extract external hyperlinks from the relationships file
        hyperlinks = []
        if 'word/_rels/document.xml.rels' in zipf.namelist():
            with zipf.open('word/_rels/document.xml.rels') as rels_file:
                rels_xml = rels_file.read()
                rels_root = ET.fromstring(rels_xml)
                for rel in rels_root:
                    if 'External' in rel.attrib.get('TargetMode', ''):
                        hyperlinks.append((rel.attrib['Id'], rel.attrib['Target']))

        # Extract internal links and bookmarks from the document XML
        internal_links = []
        bookmarks = []
        if 'word/document.xml' in zipf.namelist():
            with zipf.open('word/document.xml') as doc_xml_file:
                doc_xml = doc_xml_file.read()
                root = ET.fromstring(doc_xml)

                # Namespace map for WordprocessingML
                nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

                # Find all hyperlinks and bookmarks
                for elem in root.iter():
                    if elem.tag == f"{{{nsmap['w']}}}hyperlink":
                        anchor = elem.attrib.get(f"{{{nsmap['w']}}}anchor")
                        if anchor:
                            internal_links.append(anchor)
                    elif elem.tag == f"{{{nsmap['w']}}}bookmarkStart":
                        bookmark_name = elem.attrib.get(f"{{{nsmap['w']}}}name")
                        if bookmark_name:
                            bookmarks.append(bookmark_name)

    return hyperlinks, internal_links, bookmarks


def save_csv(data, file_path):
    with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(["Név", "Cél", "Státusz", "Oldalszám", "Módosításra került?"])
        csvwriter.writerows(data)
    logging.info(f"CSV fájl sikeresen mentve: {file_path}")


def list_and_manage_links():
    doc_path = filedialog.askopenfilename(title="Válassza ki a Word dokumentumot",
                                          filetypes=[("Word dokumentumok", "*.docx")])
    if not doc_path:
        print("Nem választott ki fájlt. A program leáll.")
        return

    output_folder = select_folder()
    if not output_folder:
        print("Nem választott ki kimeneti mappát. A program leáll.")
        return

    log_file = initialize_log_file(output_folder)
    log_message(log_file, "Dokumentum feldolgozásának kezdete")

    # Extract hyperlinks and bookmarks
    hyperlinks, internal_links, bookmarks = extract_hyperlinks_and_bookmarks(doc_path)

    links_array = []
    for link_id, target in hyperlinks:
        link_status = "Külső hivatkozás" if target.startswith("http") else "Belső hivatkozás"
        links_array.append([
            link_id,  # Link ID or text
            target,
            link_status,
            "N/A",  # Page number extraction is not directly supported
            "NEM"
        ])

    for internal_link in internal_links:
        links_array.append([
            internal_link,
            f"#{internal_link}",
            "Belső hivatkozás",
            "N/A",
            "NEM"
        ])

    for bookmark in bookmarks:
        links_array.append([
            bookmark,
            f"#{bookmark}",
            "Könyvjelző",
            "N/A",
            "NEM"
        ])

    csv_path = os.path.join(output_folder, "Frissített_Hivatkozások.csv")
    save_csv(links_array, csv_path)

    log_message(log_file, "Dokumentum mentve")
    log_message(log_file, "Naplózás vége: " + str(datetime.datetime.now()))


if __name__ == "__main__":
    list_and_manage_links()
