import os
from datetime import datetime
import time
import smtplib
import json
import pandas as pd
import psutil
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def is_program_running(program_name):
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'].lower() == program_name.lower():
            return True
    return False

def send_email(total_hours_worked, data):
    # Email content
    sender_email = data["your_email"]
    receiver_email = data["recipient_email"]
    subject = data["email_subject"]
    body = str(total_hours_worked)

    # Create the email message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    # Connect to the SMTP server (Outlook in this case)
    with smtplib.SMTP(data["SMTP_server"], 587) as server:
        server.starttls()  # Start TLS encryption
        server.login(sender_email, data["email_password"])  # Replace "your_password" with your actual password
        server.sendmail(sender_email, receiver_email, message.as_string())

def hours_calc(csv_file_path, data):
    hours_worked_series = pd.read_csv(csv_file_path, usecols=["Hours_worked"])["Hours_worked"]
    total_hours_worked = hours_worked_series.sum()

    send_email(total_hours_worked, data)

def log_hours(start_datetime, end_datetime, flag, csv_file_path, data):
    
    date_worked = start_datetime.strftime("%d/%m/%Y")
    time_difference = (end_datetime - start_datetime)
    hours_worked = time_difference.total_seconds() / 3600
    
    hour_log_df = pd.DataFrame({
        "Date": [date_worked],
        "Hours_worked": [hours_worked], 
    })
    
    if os.path.exists(csv_file_path):
        mode = 'a'
        header = False
    else:
        mode = 'w'
        header = True
    
    hour_log_df.to_csv(csv_file_path, mode=mode, header=header, index=False)

    if flag == 1:
        hours_calc(csv_file_path, data)

def main():
    with open('./config.json', 'r') as json_file:
        data = json.load(json_file)
    target_program = data["target_program"]  # Replace with your target program name
    program_was_running = False

    
    
    while True:
            program_is_running = is_program_running(target_program)

            if not program_was_running and program_is_running:
                # Program just started running, record the start time
                start_datetime = datetime.now()
            
            # Program was running but is not anymore
            if program_was_running and not program_is_running:
                end_datetime = datetime.now()

                current_date = end_datetime.strftime("%d")
                pay_date = datetime.strptime("24", "%d").strftime("%d")

                csv_file_path = data["CSV_file_path"]

                if current_date != pay_date:
                    flag = 0
                    log_hours(start_datetime, end_datetime, flag, csv_file_path, data)
                elif current_date == pay_date:
                    flag = 1
                    log_hours(start_datetime, end_datetime, flag, csv_file_path, data)
                
                sys.exit()
                
            program_was_running = program_is_running
            
            # Wait for some time before checking again
            time.sleep(1)

if __name__ == "__main__":
    main()
