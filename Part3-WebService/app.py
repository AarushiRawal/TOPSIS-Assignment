from flask import Flask, render_template, request
import os
from topsis import run_topsis
import smtplib
from email.message import EmailMessage
import re
import pandas as pd

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"



SENDER_EMAIL = "aarushirawal@gmail.com"
SENDER_PASSWORD = "xnae jada knqz ofnv"


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        file = request.files["file"]
        weights = request.form["weights"]
        impacts = request.form["impacts"]
        email = request.form["email"]

        error = validate_inputs(file, weights, impacts, email)
        if error:
            return error

        filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        file.save(filepath)

        output_file = "result.xlsx"

        run_topsis(
            filepath,
            weights,
            impacts,
            output_file
        )

        send_email(email, output_file)

        return "Result sent to your email!"

    return render_template("index.html")


def send_email(receiver_email, attachment):

    msg = EmailMessage()
    msg["Subject"] = "TOPSIS Result"
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg.set_content("Please find your TOPSIS result attached.")

    with open(attachment, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="octet-stream",
            filename=attachment
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
        smtp.send_message(msg)

def validate_inputs(file, weights, impacts, email):

    
    email_regex = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    if not re.match(email_regex, email):
        return "Invalid email format."

    try:
        weight_list = list(map(float, weights.split(",")))
    except:
        return "Weights must be numeric and comma separated."

    impact_list = impacts.split(",")

    for i in impact_list:
        if i not in ["+", "-"]:
            return "Impacts must be '+' or '-' only."

    df = pd.read_excel(file)

    criteria_count = df.shape[1] - 1

    if len(weight_list) != criteria_count:
        return f"Number of weights must be {criteria_count}."

    if len(impact_list) != criteria_count:
        return f"Number of impacts must be {criteria_count}."

    return None

if __name__ == "__main__":
    app.run(debug=True)
