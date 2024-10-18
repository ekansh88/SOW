from docx import Document
from html2docx import html2docx

# all the constants
vConstants = {
    "DOC_FILE_LOCATION" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\output.docx',
    "INPUT_1" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\input1.txt',
    "INPUT_2" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\input2.txt',
    "INPUT_3" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\input3.txt'
}

# get all the inputs fields in place of this
input1 = open(vConstants["INPUT_1"], 'r')
input1Text = input1.read();
input2 = open(vConstants["INPUT_2"], 'r')
input2Text = input1.read();
input3 = open(vConstants["INPUT_3"], 'r')
input3Text = input1.read();




def fAddOverWordDoc(text, file_name):
    doc = Document()
    
    doc.add_paragraph(text)
    doc.save(file_name)



def convert_html_to_docx(html_content, output_file):
    buf = html2docx(html_content, title="My Document")
    print(buf.getvalue())
    with open(output_file, "wb") as fp:
        fp.write(buf.getvalue())

if __name__ == "__main__":
    print("ekansh")
    convert_html_to_docx(fileText, vConstants["DOC_FILE_LOCATION"])

