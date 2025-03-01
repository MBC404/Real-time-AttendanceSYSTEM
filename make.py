import serial.tools.list_ports
import firebase_admin
from firebase_admin import credentials, db
import pandas as pd
import datetime
import time

# Initialize serial communication
try:
    serialInst = serial.Serial()
    serialInst.baudrate = 9600
    serialInst.port = "COM3"  # Update this if needed
    serialInst.open()
except Exception as e:
    print(f"⚠️ Serial Port Error: {e}")
    exit()

# Load student data from Excel
try:
    df = pd.read_excel('student.xlsx')

    if 'UID' not in df.columns or 'Student Name' not in df.columns:
        print("⚠️ Error: 'UID' or 'Student Name' column not found in Excel file.")
        exit()

    # Normalize UID column (Remove spaces, dots & ensure consistent format)
    df['UID'] = df['UID'].astype(str).str.replace("[ .]", "", regex=True).str.strip()

except Exception as e:
    print(f"⚠️ Excel File Error: {e}")
    exit()

# Initialize Firebase
try:
    cred = credentials.Certificate('your.json.file')
    firebase_admin.initialize_app(cred, {
        'databaseURL': "yourdatabaselink"
    })
except Exception as e:
    print(f"⚠️ Firebase Initialization Error: {e}")
    exit()

print("📡 Ready to scan RFID tags...")

while True:
    try:
        if serialInst.in_waiting:
            packet = serialInst.readline().decode('utf-8').strip()
            print(f"🔍 Raw RFID Data: {packet}")

            if "TAG" in packet or len(packet) < 5:
                continue

            # Normalize scanned UID (Remove spaces & dots)
            scanned_uid = packet.replace(" ", "").replace(".", "").strip()
            print(f"✅ Scanned RFID: {scanned_uid}")

            # Search for the student using cleaned UID
            matched_student = df[df['UID'] == scanned_uid]

            if not matched_student.empty:
                student_name = matched_student.iloc[0]['Student Name']
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # Write to Firebase
                ref = db.reference(f"attendance/{scanned_uid}")
                ref.set({
                    "UID": scanned_uid,
                    "Student Name": student_name,
                    "Timestamp": current_time
                })

                print(f"📌 Attendance recorded: {student_name} at {current_time}")

            else:
                print("❌ No matching student found in the database.")

            time.sleep(2)  # Prevents rapid scanning

    except Exception as e:
        print(f"⚠️ Error: {e}")
