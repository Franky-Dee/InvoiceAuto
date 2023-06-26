from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")     

invoice_list = [[2, "item", 0.5, 1]]

doc.render({"name" : "John",
            "invoice_list" : invoice_list})    
             
doc.save("new_invoice.docx")                    