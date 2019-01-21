import docx2html

html = docx2html.convert('doc.docx')

with open("result.html", "w") as file:
    file.write(html)