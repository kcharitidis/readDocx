#specific to extracting information from word documents
import os
import zipfile

#useful tool for extracting information from XML
import re

#to pretty print our xml:
import xml.dom.minidom

#to check files in the current directory, use a single period
# os.listdir('.')
#to check files in the directory above the current directory, use double periods
# os.listdir('..')
#the sample word document is in the folder entitled "docs"
# os.listdir('../docs')
document = zipfile.ZipFile('document.docx')
print document.namelist()
#name = 'word/people.xml'
#we can see who the document author is: Stephen C. Phillips
uglyXml = xml.dom.minidom.parseString(document.read('word/document.xml')).toprettyxml(indent='  ')
text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
prettyXml = text_re.sub('>\g<1></', uglyXml)
print(prettyXml)
uglyXml = xml.dom.minidom.parseString(document.read('word/fontTable.xml')).toprettyxml(indent='  ')
text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
prettyXml = text_re.sub('>\g<1></', uglyXml)
print(prettyXml)
#first to turn the xml content into a string:
xml_content = document.read('word/document.xml')
document.close()
xml_str = str(xml_content)
link_list = re.findall('http.*?\<',xml_str)[1:]
link_list = [x[:-1] for x in link_list]
print link_list
print xml_str
