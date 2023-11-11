#sub_menu.py
import os
import pdf_operations
import image_module

# Functions to get the input files
def get_file_paths(valid_pdf_paths):
    pdf_paths = input("Enter the PDF paths (comma-separated): ").split(",")

    for pdf_path in pdf_paths:
        while not os.path.isfile(pdf_path.strip()):
            print(f"Invalid PDF path: {pdf_path.strip()}. Please try again.")
            pdf_path = input("Enter the PDF path: ")
        valid_pdf_paths.append(pdf_path.strip())

def get_valid_file_path(prompt, file_extension=None):
    while True:
        file_path = input(prompt).strip()
        if os.path.isfile(file_path) and (file_extension is None or file_path.endswith(file_extension)):
            return file_path
        print(f"Invalid file path or file type: '{file_path}'. Please try again.")

def get_valid_file_name(prompt, required_extension):
    while True:
        file_name = input(prompt).strip()
        if file_name.endswith(required_extension):
            return file_name
        print(f"Invalid file name. Please ensure it ends with '{required_extension}'.")

# Menu functions
def mirror_image_menu():
    image_path = input("Enter the image path: ")
    while not os.path.isfile(image_path):
        print("Invalid image path. Please try again.")
        image_path = input("Enter the image path: ")

    direction = int(input("Enter the direction (1 for left-right mirror, 2 for top-bottom mirror): "))
    while direction not in [1, 2]:
        direction = int(input("Enter the direction (1 for left-right mirror, 2 for top-bottom mirror): "))

    image_module.mirror_image(image_path, direction)

def convert_image_menu():
    image_path = get_valid_file_path("Enter the image path: ")
    output_format = input("Enter the output format: ")
    image_module.convert_image(image_path, output_format)

def images_to_pdf_menu():
    image_paths = input("Enter the image paths (comma-separated): ").split(",")
    valid_image_paths = [get_valid_file_path(f"Enter the path for '{path.strip()}': ") for path in image_paths]
    pdf_name = get_valid_file_name("Enter the PDF name: ", '.pdf')
    image_module.images_to_pdf(valid_image_paths, pdf_name)

def split_pdf_menu():
    filename = get_valid_file_path("Enter the PDF filename: ", '.pdf')
    page_ranges = input("Enter the page ranges (comma-separated): ").split(",")
    output_filename = get_valid_file_name("Enter the output filename: ", '.pdf')
    pdf_operations.split_pdf(filename, page_ranges, output_filename)

def merge_pdf_files_menu():
    output_filename = get_valid_file_name("Enter the output PDF filename: ", '.pdf')
    all_files_choice = int(input("Press 1 to specify PDF files or 2 to merge all in the directory: "))
    merge_all = all_files_choice == 2
    pdf_operations.merge_pdf_files(output_filename, all_files=merge_all)

def pdf_to_word_menu():
    valid_pdf_paths = []
    valid_pdf_paths = get_file_paths(valid_pdf_paths)
    pdf_operations.pdf_to_word(valid_pdf_paths)

def pdf_to_images_menu():
    valid_pdf_paths = []
    get_file_paths(valid_pdf_paths)
    pdf_operations.pdf_to_images(valid_pdf_paths)

def convert_pdf_to_excel_menu():
    input_pdf_path = input("Enter the input PDF path: ")
    while not os.path.isfile(input_pdf_path.strip()):
        print(f"Invalid PDF path: {input_pdf_path.strip()}. Please try again.")
        input_pdf_path = input("Enter the input PDF path: ")

    pdf_operations.pdf_to_excel(input_pdf_path)