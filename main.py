#main.py
import sub_menu

if __name__ == "__main__":
    try:
        choice = int(input("1. Mirror Image\n2. Convert Image to another image format\n3. Convert Images to PDF\n4. Merge PDF Files\n5. Convert PDF to Word \n6. Convert PDF to Images\n7. Split PDF\n8. Convert PDF to Excel\nEnter your choice: "))
        if choice not in [1, 2, 3, 4, 5, 6, 7, 8]:
            print("Enter a number from 1 to 8")

        elif choice == 1:
            sub_menu.mirror_image_menu()
        elif choice == 2:
            sub_menu.convert_image_menu()
        elif choice == 3:
            sub_menu.images_to_pdf_menu()
        elif choice == 4:
            sub_menu.merge_pdf_files_menu()
        elif choice == 5:
            sub_menu.pdf_to_word_menu()
        elif choice == 6:
            sub_menu.pdf_to_images_menu()
        elif choice == 7:
            sub_menu.split_pdf_menu()
        elif choice == 8:
            sub_menu.convert_pdf_to_excel_menu()
    except ValueError:
        print("Enter an integer")