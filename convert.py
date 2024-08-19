import os
import comtypes.client

def convert_docx_to_pdf(docx_file, pdf_file):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(docx_file)
    doc.SaveAs(pdf_file, FileFormat=17)
    doc.Close()
    word.Quit()

def convert_all_docx_in_folder(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            docx_file = os.path.join(folder_path, filename)
            pdf_file = os.path.join(folder_path, f"{os.path.splitext(filename)[0]}.pdf")
            convert_docx_to_pdf(docx_file, pdf_file)
            print(f"Converted: {docx_file} to {pdf_file}")

if __name__ == "__main__":
    folder_path = r"C:\Users\jamil.neto\Desktop\Notebook LM\MEs"  # Replace with the path to your folder
    convert_all_docx_in_folder(folder_path)
