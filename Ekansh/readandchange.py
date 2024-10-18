import zipfile
from lxml import etree
from html2docx import html2docx
from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET
import html2text
from docx import Document
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls

vConstants = {
    "DOC_FILE_LOCATION_INPUT" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\templateTesting.docx',
    "DOC_FILE_LOCATION_OUTPUT" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\templateTestingOutput.docx',
    "INPUT_1" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\input1.txt',
    "INPUT_2" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\input2.txt',
    "INPUT_3" : r'C:\Codes\ROPs\SOW\Sow\RichTextPy\input3.txt',
}

# all the contentControl tags that we what to change their text
vReplacement = {
    'Assumptions': None,
    'Requirements': None
}

# get all the inputs fields in place of this
vInput1 = open(vConstants["INPUT_1"], 'r')
vInput1Text = vInput1.read();
vInput2 = open(vConstants["INPUT_2"], 'r')
vInput2Text = vInput2.read();
vInput3 = open(vConstants["INPUT_3"], 'r')
vInput3Text = vInput3.read();

class cRichTextConverter:
    def __init__(self) -> None:
        pass

    def fConvertHtmlToRichText(self, htmlContent):
        buf = html2docx(htmlContent, title="my text")
        return buf.getvalue()
    
    def fConvertHtmlToXML(self, htmlContent):
        # Parse the HTML content
        soup = BeautifulSoup(htmlContent, 'html.parser')

        # Create the root of the XML document
        root = ET.Element('Document')

        # Recursively convert the HTML structure to XML
        def parse_element(element, xml_parent):
            # Create an XML element for each HTML tag
            xml_element = ET.SubElement(xml_parent, element.name)

            # Copy attributes to the XML element
            for attr, value in element.attrs.items():
                xml_element.set(attr, value)

            # Add text if available
            if element.string:
                xml_element.text = element.string.strip()

            # Recursively process child elements
            for child in element.children:
                if child.name:  # Ensure the child is a tag
                    parse_element(child, xml_element)
                elif child.string:  # If it's a text node, add it to the current element
                    xml_element.text = (xml_element.text or '') + child.string.strip()

        # Start parsing from the body of the HTML
        for item in soup.contents:  # Using soup.contents to include all elements at the root level
            if item.name:  # Only parse if the item has a name (it's an element)
                parse_element(item, root)
        print(soup.contents)
        print()
        # Convert the XML tree to a string
        # print(ET.tostring(root, encoding='unicode', method='xml'))
        return ET.tostring(root, encoding='unicode', method='xml')

    def fHtmlToText(self, htmlContent, ignoreLink = True, bypassTable = True):
        # Ignore converting links from HTML
        # html2text.ignore_links = ignoreLink
        # return html2text.html2text(htmlContent)
    
        text_maker = html2text.HTML2Text()
        text_maker.ignore_links = ignoreLink
        text_maker.bypass_tables = bypassTable
        text = text_maker.handle(htmlContent).replace("*", "")
        return text

class cTemplateMaker:
    def __init__(self) -> None:
        pass

    def fReplaceContentControls(self, inputDocxPath, outputDocxPath, replacements):
        # Open the .docx as a zipfile
        with zipfile.ZipFile(inputDocxPath, 'r') as docx:
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


            control_tag = tag_element[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if tag_element else "No tag"

            # Extract the actual text inside the content control
            text_elements = sdt.xpath('.//w:t', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})


            # Check if the current text matches one of the specified keys (a, b, c, d)
            if control_tag in replacements:
                # Split the replacement text into lines, if any line breaks are present
                replacement_lines = replacements[control_tag].split('\n')

                # Replace the text in the first <w:t> element and insert line breaks for the remaining text
                for i, t in enumerate(text_elements):
                    if i == 0:
                        t.text = replacement_lines[0]  # Replace the text with the first line
                    else:
                        t.text = ""  # Clear any extra <w:t> elements
                
                # For the subsequent lines (if any), insert line breaks and new text
                current_element = text_elements[0].getparent()  # Get the parent of the first <w:t>
                for line in replacement_lines[1:]:
                    # Create a line break element
                    br_element = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')
                    current_element.append(br_element)  # Add the line break
                    
                    # Create a new text element for the next line of text
                    new_text_element = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    new_text_element.text = line
                    current_element.append(new_text_element)  # Append the new text element

        # Convert the updated XML tree back to string
        updated_xml_content = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone='yes')

        # Write the updated XML back to a new .docx file
        with zipfile.ZipFile(inputDocxPath, 'r') as docx:
            with zipfile.ZipFile(outputDocxPath, 'w') as output_docx:
                # Copy all files from the original .docx except document.xml
                for item in docx.infolist():
                    if item.filename != 'word/document.xml':
                        output_docx.writestr(item.filename, docx.read(item.filename))
                # Write the updated document.xml
                output_docx.writestr('word/document.xml', updated_xml_content)
    
    def _fReplaceContentControlsContentWithXML(self, doc, contentControlTag, contentControlText):
        # Iterate over all the SDT (Structured Document Tags)
        for sdt in doc.element.findall('.//w:sdt', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
            # Check if the content control has a title
            title = sdt.find('.//w:title', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            
            # Optionally, check for a specific title
            if title is not None and title.text == contentControlTag:
                # Find the content inside the content control
                content = sdt.find('.//w:sdtContent', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if content is not None:
                    # Clear existing paragraphs inside the content control
                    for paragraph in content.findall('.//w:p', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        content.remove(paragraph)
                    
                    # Create a new paragraph and run to hold the new text
                    new_paragraph = OxmlElement('w:p')
                    new_run = OxmlElement('w:r')
                    new_text_element = OxmlElement('w:t')
                    new_text_element.text = contentControlText
                    
                    # Append the new text to the run, and the run to the paragraph
                    new_run.append(new_text_element)
                    new_paragraph.append(new_run)
                    
                    # Add the new paragraph back to the content control
                    content.append(new_paragraph)


if __name__ == "__main__":

    oTemplateMaker = cTemplateMaker()

    oRichTextConverter = cRichTextConverter()
    vReplacement["Assumptions"] = oRichTextConverter.fHtmlToText(vInput1Text)
    vReplacement["Requirements"] = oRichTextConverter.fHtmlToText(vInput2Text)

 
    oTemplateMaker.fReplaceContentControls(vConstants["DOC_FILE_LOCATION_INPUT"], vConstants["DOC_FILE_LOCATION_OUTPUT"], vReplacement)