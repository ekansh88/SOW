import zipfile
from lxml import etree

def update_content_controls(docx_path, output_path, replacements):
    # Open the .docx as a zipfile
    with zipfile.ZipFile(docx_path, 'r') as docx:
        # Extract the document.xml which contains the main content
        xml_content = docx.read('word/document.xml')
    
    # Parse the XML content using lxml
    tree = etree.XML(xml_content)

    # XPath to find all content controls (SDTs), including nested ones
    sdt_elements = tree.xpath('//w:sdt', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})

    # Iterate through all found SDTs (content controls)
    for sdt in sdt_elements:
        # Extract the content control's title (if exists)
        alias_element = sdt.xpath('.//w:alias', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        tag_element = sdt.xpath('.//w:tag', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})

        control_title = alias_element[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if alias_element else "No title"
        control_tag = tag_element[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if tag_element else "No tag"

        # Extract the actual text inside the content control
        text_elements = sdt.xpath('.//w:t', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})

        current_text = ''.join([t.text for t in text_elements if t.text])
        

        # Check if the current text matches one of the specified keys (a, b, c, d)
        if control_tag in replacements:
            # Replace the text with "Ekansh Sharma"
            print("//got the tag",control_tag)
            # text_elements.text = replacements[control_tag]
            iterator = 0
            for t in text_elements:
                if iterator==0:
                    t.text = replacements[control_tag]  # Replace text with the new value
                else:
                    t.text = ""
                iterator += 1
    # Convert the updated XML tree back to string
    updated_xml_content = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone='yes')

    # Write the updated XML back to a new .docx file
    with zipfile.ZipFile(docx_path, 'r') as docx:
        with zipfile.ZipFile(output_path, 'w') as output_docx:
            # Copy all files from the original .docx except document.xml
            for item in docx.infolist():
                if item.filename != 'word/document.xml':
                    output_docx.writestr(item.filename, docx.read(item.filename))
            # Write the updated document.xml
            output_docx.writestr('word/document.xml', updated_xml_content)

# Specify the replacements for specific content controls
replacements = {
    'Assumptions': 'this is me ekansh',
    'Requirements': 'Ekansh Sharma'
}

# Call the function to update the content controls and save to a new .docx file
update_content_controls(r'C:\Codes\ROPs\SOW\Sow\RichTextPy\templateTesting.docx', r'C:\Codes\ROPs\SOW\Sow\RichTextPy\templateTestingOutput.docx', replacements)
