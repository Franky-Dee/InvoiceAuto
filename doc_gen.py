from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")     # Template File

doc.render({"name" : "John"})                   # Example
doc.save("new_invoice.docx")                    # New File