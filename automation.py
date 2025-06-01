from flask import Flask, render_template, request, send_file
from docx import Document
import os

app = Flask(__name__)

# Fill Word template with values

def fill_docx(template_path, output_path, context):
    doc = Document(template_path)

    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        full_text = paragraph.text
        for key, value in context.items():
            full_text = full_text.replace(f"{{{{{key}}}}}", str(value))
        # Rebuild paragraph with replaced text
        if paragraph.text != full_text:
            for run in paragraph.runs:
                run.text = ''
            paragraph.runs[0].text = full_text if paragraph.runs else paragraph.add_run(full_text)

    # Replace in tables too (if any)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    full_text = paragraph.text
                    for key, value in context.items():
                        full_text = full_text.replace(f"{{{{{key}}}}}", str(value))
                    if paragraph.text != full_text:
                        for run in paragraph.runs:
                            run.text = ''
                        paragraph.runs[0].text = full_text if paragraph.runs else paragraph.add_run(full_text)

    doc.save(output_path)

@app.route("/")
def home():
    return render_template("main.html")

@app.route("/appointment_letter.html", methods=["GET", "POST"])
def appointment():
    if request.method == "POST":
        data = request.form.to_dict()
        data["name"] = f"{data['first_name']} {data['last_name']}"
        filename = f"{data['name'].replace(' ', '_')}_Appointment_Letter.docx"
        output_path = os.path.join("generated", filename)
        fill_docx("Appointment_letter.docx", output_path, data)
        return send_file(output_path, as_attachment=True)
    return render_template("appointment_letter.html")

@app.route("/experience_letter.html", methods=["GET", "POST"])
def experience():
    if request.method == "POST":
        data = request.form.to_dict()
        filename = f"{data['name'].replace(' ', '_')}_Experience_Letter.docx"
        output_path = os.path.join("generated", filename)
        fill_docx("Experience_letter.docx", output_path, data)
        return send_file(output_path, as_attachment=True)
    return render_template("experience_letter.html")

@app.route("/internship_letter.html", methods=["GET", "POST"])
def internship():
    if request.method == "POST":
        data = request.form.to_dict()
        filename = f"{data['name'].replace(' ', '_')}_Internship_Letter.docx"
        output_path = os.path.join("generated", filename)
        fill_docx("Internship_letter.docx", output_path, data)
        return send_file(output_path, as_attachment=True)
    return render_template("internship_letter.html")

if __name__ == "__main__":
    os.makedirs("generated", exist_ok=True)
    app.run(debug=True)
