from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")     # Template File

invoice_list = [[2, "item", 0.5, 1]]

doc.render({"name" : "John", "invoice_list" : invoice_list})                   # Example of fields to be inserted into template
doc.save("new_invoice.docx")                    # New File