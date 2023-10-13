import os
import logging
from configparser import ConfigParser
from concurrent.futures import ProcessPoolExecutor
from docx2pdf import convert

def doc_to_pdf(doc_file_path, pdf_file_path):
    convert(doc_file_path, pdf_file_path)

def create_directory_if_not_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def process_doc_file(doc_file, doc_folder, pdf_folder):
    file_name, extension_name = os.path.splitext(doc_file)
    if extension_name != ".docx":
        return

    doc_file_path = os.path.join(doc_folder, doc_file)
    pdf_file_path = os.path.join(pdf_folder, file_name + ".pdf")

    print("⌛️正在处理：" + doc_file_path)
    doc_to_pdf(doc_file_path, pdf_file_path)
    print(pdf_file_path + "处理完成✅")

def main():
    logging.getLogger().setLevel(logging.ERROR)

    config_parser = ConfigParser()
    config_parser.read("config.cfg")
    config = config_parser["default"]

    pdf_folder = config["pdf_folder"]
    doc_folder = config["doc_folder"]
    max_workers = int(config["max_worker"])

    if not os.path.exists(doc_folder):
        print(f"'{pdf_folder}'文件夹不存在，请创建该文件夹。")
        return

    create_directory_if_not_exists(pdf_folder)

    doc_files = [file for file in os.listdir(doc_folder) if file.endswith(".docx")]

    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        for doc_file in doc_files:
            executor.submit(process_doc_file, doc_file, doc_folder, pdf_folder)

    print("所有文件处理完成✅")

if __name__ == "__main__":
    main()
