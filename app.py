from flask import Flask, request, render_template_string
import pandas as pd
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import os

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>CT Institute Admission Form</title>

    <style>

        body{
            margin:0;
            padding:0;
            font-family:Arial, sans-serif;
            background:linear-gradient(135deg,#ff9a9e,#fad0c4,#fbc2eb,#a18cd1);
            min-height:100vh;
        }

        .container{
            width:90%;
            max-width:650px;
            background:white;
            margin:30px auto;
            padding:30px;
            border-radius:20px;
            box-shadow:0 0 20px rgba(0,0,0,0.3);
        }

        .logo{
            width:160px;
            height:160px;
            border-radius:50%;
            background:black;
            color:gold;
            font-size:70px;
            font-weight:bold;
            display:flex;
            align-items:center;
            justify-content:center;
            margin:auto;
            border:5px solid gold;
        }

        h1{
            text-align:center;
            color:#111;
        }

        .title{
            text-align:center;
            color:#555;
            margin-bottom:20px;
            font-size:18px;
        }

        label{
            font-weight:bold;
            display:block;
            margin-top:15px;
            color:#333;
        }

        input,
        textarea,
        select{
            width:100%;
            padding:12px;
            margin-top:5px;
            border-radius:10px;
            border:2px solid #ddd;
            box-sizing:border-box;
            font-size:16px;
        }

        textarea{
            height:100px;
            resize:none;
        }

        small{
            color:gray;
        }

        .submit-btn{
            width:100%;
            background:linear-gradient(90deg,#ff512f,#dd2476);
            color:white;
            border:none;
            padding:15px;
            border-radius:12px;
            font-size:20px;
            margin-top:25px;
            cursor:pointer;
        }

        .footer{
            text-align:center;
            margin-top:20px;
            color:#444;
        }

    </style>

</head>

<body>

<div class="container">

    <div class="logo">CT</div>

    <h1>CT Institute</h1>

    <div class="title">
        Online Admission Form <br>
        Admission Through Vikash Maurya <br>
        Location: Deoria, Uttar Pradesh
    </div>

    <form method="POST">

        <label>Student Name</label>
        <input
            type="text"
            name="student_name"
            placeholder="Enter Full Name"
            pattern="[A-Za-z ]+"
            title="Only letters allowed"
            required
        >
        <small>Only alphabets allowed</small>

        <label>Father Name</label>
        <input
            type="text"
            name="father_name"
            placeholder="Enter Father Name"
            pattern="[A-Za-z ]+"
            title="Only letters allowed"
            required
        >

        <label>Mother Name</label>
        <input
            type="text"
            name="mother_name"
            placeholder="Enter Mother Name"
            pattern="[A-Za-z ]+"
            title="Only letters allowed"
            required
        >

        <label>Mobile Number</label>
        <input
            type="number"
            name="mobile"
            placeholder="Enter 10 Digit Mobile Number"
            oninput="if(this.value.length > 10) this.value = this.value.slice(0,10);"
            required
        >
        <small>Only 10 digit number allowed</small>

        <label>Aadhaar Number</label>
        <input
            type="number"
            name="aadhaar"
            placeholder="Enter 12 Digit Aadhaar Number"
            oninput="if(this.value.length > 12) this.value = this.value.slice(0,12);"
            required
        >
        <small>Only 12 digit Aadhaar number allowed</small>

        <label>Address</label>
        <textarea
            name="address"
            placeholder="Enter Full Address"
            required
        ></textarea>

        <label>Select Course</label>

        <select name="course">

            <option>ADCA</option>
            <option>DCA</option>
            <option>CCC</option>
            <option>O LEVEL</option>
            <option>TALLY</option>

        </select>

        <label>Qualification</label>

        <input
            type="text"
            name="qualification"
            placeholder="10th / 12th / Graduate"
            required
        >

        <button class="submit-btn" type="submit">
            Submit Admission Form
        </button>

    </form>

    <div class="footer">
        © CT Institute | Powered By Vikash Maurya
    </div>

</div>

</body>
</html>
"""


@app.route('/', methods=['GET', 'POST'])
def home():

    if request.method == 'POST':

        data = {
            "Date": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
            "Student Name": request.form['student_name'],
            "Father Name": request.form['father_name'],
            "Mother Name": request.form['mother_name'],
            "Mobile": request.form['mobile'],
            "Aadhaar": request.form['aadhaar'],
            "Address": request.form['address'],
            "Course": request.form['course'],
            "Qualification": request.form['qualification']
        }

        # ================= EXCEL SAVE =================

        excel_file = "admission_data.xlsx"

        if os.path.exists(excel_file):
            old_data = pd.read_excel(excel_file)
            new_data = pd.concat(
                [old_data, pd.DataFrame([data])],
                ignore_index=True
            )
        else:
            new_data = pd.DataFrame([data])

        new_data.to_excel(excel_file, index=False)

        # ================= PDF SAVE =================

        if not os.path.exists("pdf_receipts"):
            os.makedirs("pdf_receipts")

        pdf_file = f"pdf_receipts/{request.form['student_name']}.pdf"

        doc = SimpleDocTemplate(pdf_file)
        styles = getSampleStyleSheet()
        story = []

        story.append(
            Paragraph("CT Institute Admission Form", styles['Title'])
        )

        story.append(Spacer(1, 20))

        for key, value in data.items():
            story.append(
                Paragraph(f"<b>{key}:</b> {value}", styles['BodyText'])
            )
            story.append(Spacer(1, 10))

        doc.build(story)

        return f"""
        <h1 style='text-align:center;
        color:green;
        margin-top:100px;
        font-family:Arial;'>

        Thank You {request.form['student_name']}! <br><br>

        Your Admission Form Has Been Submitted Successfully 🎉

        </h1>
        """

    return render_template_string(HTML)


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)