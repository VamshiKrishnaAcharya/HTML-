from flask import Flask, request, jsonify
from openpyxl import Workbook, load_workbook
from pathlib import Path
from datetime import datetime
import threading

app = Flask(__name__)

# Excel file
EXCEL_FILE = Path("registrations.xlsx")

# Prevent simultaneous writes
excel_lock = threading.Lock()


HEADERS = [
    "Registration ID",
    "Date & Time",
    "Institution",
    "District",
    "Guru / Principal",
    "Contact Phone",
    "Team Name",

    "Student 1",
    "Age 1",

    "Student 2",
    "Age 2",

    "Student 3",
    "Age 3",

    "Student 4",
    "Age 4",

    "Student 5",
    "Age 5",

    "UTR / Transaction ID",
    "Amount",
    "Payment Status"
]


def create_excel_if_missing():

    if not EXCEL_FILE.exists():

        workbook = Workbook()

        worksheet = workbook.active
        worksheet.title = "Registrations"

        worksheet.append(HEADERS)

        # Freeze header
        worksheet.freeze_panes = "A2"

        # Column widths
        widths = {
            "A": 20,
            "B": 22,
            "C": 30,
            "D": 25,
            "E": 25,
            "F": 16,
            "G": 25,
            "H": 25,
            "I": 10,
            "J": 25,
            "K": 10,
            "L": 25,
            "M": 10,
            "N": 25,
            "O": 10,
            "P": 25,
            "Q": 10,
            "R": 22,
            "S": 12,
            "T": 18
        }

        for column, width in widths.items():
            worksheet.column_dimensions[column].width = width

        workbook.save(EXCEL_FILE)


def generate_registration_id():

    create_excel_if_missing()

    with excel_lock:

        workbook = load_workbook(EXCEL_FILE)
        worksheet = workbook["Registrations"]

        registration_number = worksheet.max_row

        registration_id = (
            f"ANTY-2026-{registration_number:04d}"
        )

        workbook.close()

    return registration_id


@app.route("/register", methods=["POST"])
def register():

    try:

        data = request.get_json()

        if not data:
            return jsonify({
                "message": "No registration data received"
            }), 400

        # Generate unique registration ID
        registration_id = generate_registration_id()

        date_time = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        row = [
            registration_id,
            date_time,

            data.get("instName", ""),
            data.get("district", ""),
            data.get("guruName", ""),
            data.get("phone", ""),
            data.get("teamName", ""),

            data.get("student1", ""),
            data.get("age1", ""),

            data.get("student2", ""),
            data.get("age2", ""),

            data.get("student3", ""),
            data.get("age3", ""),

            data.get("student4", ""),
            data.get("age4", ""),

            data.get("student5", ""),
            data.get("age5", ""),

            data.get("utr", ""),

            500,
            "Paid / UTR Submitted"
        ]

        with excel_lock:

            create_excel_if_missing()

            workbook = load_workbook(EXCEL_FILE)
            worksheet = workbook["Registrations"]

            worksheet.append(row)

            workbook.save(EXCEL_FILE)

            workbook.close()

        return jsonify({
            "success": True,
            "registration_id": registration_id,
            "message": "Registration saved successfully"
        })

    except Exception as error:

        print("ERROR:", error)

        return jsonify({
            "success": False,
            "message": str(error)
        }), 500


@app.route("/", methods=["GET"])
def home():

    return """
    <h2>Spiritual Antyakshari Registration Server</h2>
    <p>Registration server is running.</p>
    """


if __name__ == "__main__":

    create_excel_if_missing()

    print("------------------------------------------")
    print("Spiritual Antyakshari Registration Server")
    print("------------------------------------------")
    print("Excel file:", EXCEL_FILE.absolute())
    print("Server: http://127.0.0.1:5000")
    print("------------------------------------------")

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )