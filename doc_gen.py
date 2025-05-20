from docxtpl import DocxTemplate


doc = DocxTemplate("invoice_template.docx")

invoice_list = [[2, "Pen", 5.00, 1],
                [1, "Notebook", 50.00, 2],
                [2, "Paper", 2.00, 5]]


doc.render({"name":"Karim",
            "phone":"017564325",
            "invoice_list": invoice_list,
            "subtotal":115,
            "salestax":"5%",
            "total": 115})

doc.save("new_invoice.docx")