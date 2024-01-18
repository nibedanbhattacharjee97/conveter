import pandas as pd
from flask import Flask, render_template, request
import os

app = Flask(__name__)

def excel_to_csv(input_excel_path, output_csv_path):
    try:
        # Load Excel file into a pandas DataFrame
        df = pd.read_excel(input_excel_path)

        # Save DataFrame to a CSV file
        df.to_csv(output_csv_path, index=False)

        return True, f"Conversion successful. CSV file saved at: {output_csv_path}"

    except Exception as e:
        return False, f"Error during conversion: {e}"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Check if a file was provided in the request
        if "file" not in request.files:
            return render_template("index.html", error="No file provided.")

        file = request.files["file"]

        # Check if the file has an allowed extension (e.g., Excel)
        if file and file.filename.endswith((".xls", ".xlsx")):
            input_excel_path = os.path.join("uploads", file.filename)
            output_csv_path = os.path.join("uploads", file.filename.rsplit(".", 1)[0] + ".csv")

            # Save the uploaded file to the server
            file.save(input_excel_path)

            # Call the function to convert Excel to CSV
            success, message = excel_to_csv(input_excel_path, output_csv_path)

            return render_template("index.html", success=success, message=message)
        else:
            return render_template("index.html", error="Invalid file format. Please upload an Excel file.")

    return render_template("index.html")

if __name__ == "__main__":
    if not os.path.exists("uploads"):
        os.makedirs("uploads")
    
    app.run(debug=True)
