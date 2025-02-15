from docx import Document
import pypandoc

template_path = "./template.docx"

students = [
    {"name":"Ali","roll":"55"},
    {"name":"Noufan Elachola","roll":"69"},
    {"name":"Razin","roll":"72"},
    {"name":"Hashim","roll":"67"},
]

for student in students:
    new_doc = Document(template_path)

    for para in new_doc.paragraphs:
        if "{name}" in para.text:
            para.text = para.text.replace("{name}", student["name"])
        if "{roll}" in para.text:
            para.text = para.text.replace("{roll}", student["roll"])

    docx_path = f"output/{student['roll']}_{student['name'].replace(' ', '_')}.docx"
    pdf_path = f"output/{student['roll']}_{student['name'].replace(' ', '_')}.pdf"
    new_doc.save(docx_path)

    pypandoc.convert_file(docx_path, "pdf", outputfile=pdf_path)

print("Documents generated successfully!")