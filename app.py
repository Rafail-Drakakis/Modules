import os, tempfile, glob, pdf2image, pdf2docx, PyPDF2, tabula, openpyxl, warnings ,PIL
from flask import Flask, render_template_string, request, send_file
from werkzeug.utils import secure_filename
import pandas as pd

def pdf_to_excel(input_pdf_path):
    # Define output Excel file path
    output_excel_path = os.path.splitext(input_pdf_path)[0] + ".xlsx"

    # Create a temporary directory for storing temporary CSV files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_csv_files = []

        # Convert each page of the PDF to a separate CSV file
        pdf_reader = PyPDF2.PdfReader(input_pdf_path)
        num_pages = len(pdf_reader.pages)

        for page_num in range(num_pages):
            temp_csv_file = os.path.join(temp_dir, f"Page{page_num + 1}.csv")
            temp_csv_files.append(temp_csv_file)

            try:
                tabula.convert_into(input_pdf_path, temp_csv_file, output_format="csv", pages=page_num + 1)
            except Exception as e:
                print(f"An error occurred while converting page {page_num + 1} to CSV: {e}")

        # Merge all the separate CSV files into one DataFrame
        dfs = []
        for temp_csv_file in temp_csv_files:
            df = pd.read_csv(temp_csv_file)
            dfs.append(df)

        # Concatenate all DataFrames into a single DataFrame, skipping the title of every page after the first page
        merged_data_frame = dfs[0]  # Initialize with the first page
        for df in dfs[1:]:  # Skip the first page's title
            merged_data_frame = pd.concat([merged_data_frame, df], ignore_index=True)

        # Save the merged DataFrame to Excel
        merged_data_frame.to_excel(output_excel_path, index=False)

        # Load the workbook and select the active worksheet
        workbook = openpyxl.load_workbook(output_excel_path)
        sheet = workbook.active

        # Iterate through columns
        for column in sheet.columns:
            max_length = 0
            column_letter = openpyxl.utils.get_column_letter(column[0].column)

            # Find the maximum length in each column
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass

            # Adjust the column width
            adjusted_width = (max_length + 3)
            sheet.column_dimensions[column_letter].width = adjusted_width

        # Save the workbook
        workbook.save(output_excel_path)
        
        # Check if the user wants to shift empty cells
        option = input("Do you want to shift empty cells? (yes/no) ")
        if option == "yes":
            excel_workbook = openpyxl.load_workbook(output_excel_path)
            for worksheet_name in excel_workbook.sheetnames:
                excel_worksheet = excel_workbook[worksheet_name]
                for row in excel_worksheet.iter_rows():
                    empty_cell_list = []
                    for current_cell in row:
                        if current_cell.value is None:
                            empty_cell_list.append(current_cell)
                        else:
                            for empty_cell_to_fill in empty_cell_list:
                                empty_cell_to_fill.value = current_cell.value
                                current_cell.value = None
                                empty_cell_list.remove(empty_cell_to_fill)
                                empty_cell_list.append(current_cell)
                                break
            excel_workbook.save(output_excel_path)

        print(f"PDF converted to Excel: {output_excel_path}")

def split_pdf(filename, page_ranges, output_filename):
    try:
        pdf = PyPDF2.PdfReader(filename)
        output_pdf = PyPDF2.PdfWriter()

        for range_string in page_ranges:
            pages = []
            ranges = range_string.split(",")
            for rng in ranges:
                if "-" in rng:
                    start, end = map(int, rng.split("-"))
                    pages.extend(range(start, end + 1))
                else:
                    pages.append(int(rng))
            for page_num in pages:
                output_pdf.add_page(pdf.pages[page_num - 1])

        with open(output_filename, "wb") as output_file:
            output_pdf.write(output_file)

    except Exception as e:
        print(f"An error occurred: {str(e)}")

def merge_pdf_files(output_filename, all_files):
    try:
        if all_files:
            files = sorted(glob.glob(os.path.join(os.getcwd(), '*.pdf')))
            if not files:
                print("Error: No PDF files found in the current directory.")
                return

            merger = PyPDF2.PdfMerger()
            for filename in files:
                merger.append(filename)
            merger.write(output_filename)
            merger.close()
        else:
            filenames = []
            while True:
                try:
                    filename = input("Enter the filename (or 'done' to finish): ")
                    if filename == "done":
                        break
                    if not os.path.exists(filename):
                        print(f"Error: File '{filename}' does not exist.")
                        continue
                    filenames.append(filename)
                except Exception as e:
                    print(f"An error occurred: {e}")

            if not filenames:
                print("No files specified.")
                return

            merger = PyPDF2.PdfMerger()
            for filename in filenames:
                merger.append(filename)
            merger.write(output_filename)
            merger.close()

    except Exception as e:
        print(f"An error occurred: {e}")

def pdf_to_word(pdf_paths):
    try:
        docx_paths = []
        for pdf_path in pdf_paths:
            try:
                docx_path = pdf_path.replace(".pdf", ".docx")
                pdf2docx.parse(pdf_path, docx_path)
                docx_paths.append(docx_path)
            except Exception as e:
                print(f"Error converting PDF to Word: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

def pdf_to_images(pdf_paths):
    try:
        for pdf_path in pdf_paths:
            try:
                #ignore warnings
                warnings.simplefilter('ignore', PIL.Image.DecompressionBombWarning)
                images = pdf2image.convert_from_path(pdf_path, dpi=1000)
                for idx, img in enumerate(images):
                    img.save(f'page_{idx + 1}.jpg', 'JPEG', quality=80)
            except Exception as e:
                print(f"Error converting PDF to images: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

def mirror_image(input_path, direction, output_dir=None, output_format='png'):
    if not os.path.isfile(input_path):
        return f"Error: {input_path} does not exist"
    if output_dir is None:
        output_dir = os.path.dirname(input_path)
    output_filename = os.path.splitext(os.path.basename(input_path))[0]
    try:
        with PIL.Image.open(input_path) as img:
            if direction == 1:
                mirror_img = img.transpose(PIL.Image.FLIP_LEFT_RIGHT)
                mirror_output_path = os.path.join(output_dir, f"{output_filename}_mirror.{output_format.lower()}")
                mirror_img.save(mirror_output_path)
                return mirror_output_path
            elif direction == 2:
                mirror_img = img.transpose(PIL.Image.FLIP_TOP_BOTTOM)
                mirror_output_path = os.path.join(output_dir, f"{output_filename}_flip.{output_format.lower()}")
                mirror_img.save(mirror_output_path)
                return mirror_output_path
            else:
                return "Invalid direction specified"
    except OSError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error: {e}"

def convert_image(input_path, output_path):
    try:
        image = PIL.Image.open(input_path) # Open the image file
        image = image.convert("RGB") # Convert the image to the RGB mode
        temp_path = "temp.png" # Create a temporary path to save the image in PNG format
        image.save(temp_path, "PNG") # Save the image in PNG format
        temp_image = PIL.Image.open(temp_path) # Load the temporary image
        temp_image.save(output_path, quality=95) # Save the temporary image in the desired output format with specified quality
        os.remove(temp_path) # Remove the temporary image file
    except FileNotFoundError:
        print("Input file not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def images_to_pdf(images, pdf_name):
    try:
        pdf_images = []
        for image in images:
            img = PIL.Image.open(image)
            pdf_images.append(img)

        if pdf_images:
            pdf_images[0].save(pdf_name, "PDF", resolution=100.0, save_all=True, append_images=pdf_images[1:])
        else:
            print("Error: No images found.")
    except Exception as e:
        print("Error: Failed to convert images to PDF.\nError:", str(e))

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

    mirror_image(image_path, direction)

def convert_image_menu():
    image_path = get_valid_file_path("Enter the image path: ")
    output_format = input("Enter the output format: ")
    convert_image(image_path, output_format)

def images_to_pdf_menu():
    image_paths = input("Enter the image paths (comma-separated): ").split(",")
    valid_image_paths = [get_valid_file_path(f"Enter the path for '{path.strip()}': ") for path in image_paths]
    pdf_name = get_valid_file_name("Enter the PDF name: ", '.pdf')
    images_to_pdf(valid_image_paths, pdf_name)

def split_pdf_menu():
    filename = get_valid_file_path("Enter the PDF filename: ", '.pdf')
    page_ranges = input("Enter the page ranges (comma-separated): ").split(",")
    output_filename = get_valid_file_name("Enter the output filename: ", '.pdf')
    split_pdf(filename, page_ranges, output_filename)

def merge_pdf_files_menu():
    output_filename = get_valid_file_name("Enter the output PDF filename: ", '.pdf')
    all_files_choice = int(input("Press 1 to specify PDF files or 2 to merge all in the directory: "))
    merge_all = all_files_choice == 2
    merge_pdf_files(output_filename, all_files=merge_all)

def pdf_to_word_menu():
    valid_pdf_paths = []
    valid_pdf_paths = get_file_paths(valid_pdf_paths)
    pdf_to_word(valid_pdf_paths)

def pdf_to_images_menu():
    valid_pdf_paths = []
    get_file_paths(valid_pdf_paths)
    pdf_to_images(valid_pdf_paths)

def convert_pdf_to_excel_menu():
    input_pdf_path = input("Enter the input PDF path: ")
    while not os.path.isfile(input_pdf_path.strip()):
        print(f"Invalid PDF path: {input_pdf_path.strip()}. Please try again.")
        input_pdf_path = input("Enter the input PDF path: ")

    pdf_to_excel(input_pdf_path)

app = Flask(__name__)

# Configure upload folder
UPLOAD_FOLDER = tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF Toolkit</title>
</head>
<body>
    <h1>PDF Toolkit</h1>
    <form action="/process" method="post" enctype="multipart/form-data">
        <label for="action">Choose Action:</label>
        <select id="action" name="action" required>
            <option value="pdf_to_excel">PDF to Excel</option>
            <option value="split_pdf">Split PDF</option>
            <option value="merge_pdfs">Merge PDFs</option>
            <option value="pdf_to_word">PDF to Word</option>
            <option value="pdf_to_images">PDF to Images</option>
        </select>

        <br><br>

        <label for="file">Choose File(s):</label>
        <input type="file" id="file" name="file" multiple>

        <br><br>

        <label for="additional_input">Additional Input (if needed):</label>
        <input type="text" id="additional_input" name="additional_input">

        <br><br>

        <button type="submit">Process</button>
    </form>
</body>
</html>
"""

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/process', methods=['POST'])
def process():
    action = request.form.get('action')
    additional_input = request.form.get('additional_input')
    files = request.files.getlist('file')

    if not action or not files:
        return "Invalid input. Please select an action and upload files.", 400

    # Save uploaded files temporarily
    input_files = []
    for file in files:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(file_path)
        input_files.append(file_path)

    try:
        if action == 'pdf_to_excel':
            pdf_to_excel(input_files[0])
            output_file = input_files[0].replace('.pdf', '.xlsx')
        elif action == 'split_pdf':
            if not additional_input:
                return "Page ranges are required for splitting PDFs.", 400
            output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'split_output.pdf')
            split_pdf(input_files[0], additional_input.split(','), output_file)
        elif action == 'merge_pdfs':
            output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'merged_output.pdf')
            merge_pdf_files(output_file, all_files=False)
        elif action == 'pdf_to_word':
            pdf_to_word(input_files)
            output_file = input_files[0].replace('.pdf', '.docx')
        elif action == 'pdf_to_images':
            pdf_to_images(input_files)
            output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'page_1.jpg')
        else:
            return "Invalid action.", 400

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        return f"An error occurred: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
