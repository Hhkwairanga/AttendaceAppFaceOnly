import configparser
import shutil
import tkinter as tk
from tkinter import PhotoImage, ttk, simpledialog
import face_recognition
import cv2
import numpy as np
import sqlite3
import os
from face_recognition import face_locations, face_encodings
import csv
from datetime import datetime, timedelta
import subprocess
from tkinter import messagebox
from tkinter import filedialog
from collections import defaultdict
import pyttsx3
from tkcalendar import Calendar, DateEntry
import schedule
import threading
import time
from datetime import datetime, timedelta
import calendar
from docx import Document
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import urllib.request
import ssl

class WelcomeFrame(tk.Frame):
   def connect_to_database(self):
       # Get the current working directory
       current_directory = os.getcwd()

       # Create the path to the database file using os.path.join()
       db_path = os.path.join(current_directory, "enrollment_data.db")

       # Connect to the SQLite3 database
       connection = sqlite3.connect(db_path)

       return connection

   def recognize_faces(self, database_filename):
       # Connect to the SQLite database in the current working directory
       conn = sqlite3.connect(database_filename)
       cursor = conn.cursor()

       # Retrieve data from the database
       cursor.execute('SELECT name, face_image_path FROM corpers UNION SELECT name, face_image_path FROM interns')
       rows = cursor.fetchall()

       # Create a list to store face encodings and corresponding names
       known_face_encodings = []
       known_face_names = []

       # Load known face encodings and names from the database
       for name, face_folder_path in rows:
           print("Current Folder Path:", face_folder_path)

           for i in range(1, 4):  # Assuming there are 10 images per folder
               filename = f"capture_{i}.png"  # Adjust the extension if needed
               image_path = os.path.join(face_folder_path, filename)
               print("Current Image Path:", image_path)

               if os.path.exists(image_path):
                   # Load the image and compute face encoding
                   image = face_recognition.load_image_file(image_path)
                   face_locations = face_recognition.face_locations(image)
                   face_encoding = face_recognition.face_encodings(image)

                   if face_encoding:
                       known_face_encodings.append(face_encoding[0])
                       known_face_names.append(name)
                   else:
                       print(f"No face found in {filename} for {name}.")
               else:
                   print(f"File not found: {filename} for {name}.")

       conn.close()

       return known_face_encodings, known_face_names

   def __init__(self, master=None):
       super().__init__(master)
       self.master = master
       self.master.geometry(f"{self.master.winfo_screenwidth()}x{self.master.winfo_screenheight()}")
       self.pack()
       self.create_widgets()
       self.check_expired_exemptions()
       # Load known face encodings and names from the database
       # self.known_face_encodings, self.known_face_names = self.recognize_faces('enrollment_data.db')#

       # Store the original size of the background logo
       self.background_logo_original = self.background_logo

       # Center the logo initially
       self.center_logo()

       # Bind the resize event to adjust the logo position
       self.master.bind("<Configure>", self.handle_resize)

       # Initialize settings as instance variables
       self.load_settings()

       self.check_first_run()


   def init_tasks(self):
       # Call the method every 1hour
       self.schedule_log_absentees()

       # Call the method to check due dates and remove users
       self.check_due_dates_and_remove_users()

       # call schedule monthly report
       self.schedule_monthly_report()

   def check_first_run(self):
       # Check if a marker file exists
       marker_file = "first_run_marker.txt"
       if os.path.isfile(marker_file):
           print("App has been run before.")
           # If it's not the first run, proceed with initialization tasks
           self.init_tasks()
       else:
           print("First time running the app. Showing configuration prompt.")

           # Display a message box to prompt the user to configure settings
           messagebox.showinfo("First Run",
                               "Welcome! It seems this is your first time running the application. Please configure your settings.")

           # Create a marker file to indicate that the app has been run before
           with open(marker_file, "w") as f:
               f.write("This file is created on the first run of the application.")

   def load_settings(self):
       # Load settings from a configuration file
       config = configparser.ConfigParser()
       config.read('config.ini')

       # Admin credentials
       self.admin_username = config.get('Admin', 'Username', fallback='admin')
       self.admin_password = config.get('Admin', 'Password', fallback='admin123')

       # Time settings
       late_time_str = config.get('TimeSettings', 'LateTime', fallback='08:40:00')
       closing_time_str = config.get('TimeSettings', 'ClosingTime', fallback='14:00:00')

       self.late_time = datetime.strptime(late_time_str, "%H:%M:%S").time()
       self.closing_time = datetime.strptime(closing_time_str, "%H:%M:%S").time()

       # Max days late and max days absent settings
       self.max_days_late = int(config.get('TimeSettings', 'MaxDaysLate', fallback='3'))
       self.max_days_absent = int(config.get('TimeSettings', 'MaxDaysAbsent', fallback='3'))

       # App email and password settings
       self.app_email = config.get('TimeSettings', 'AppEmail', fallback='')
       self.app_password = config.get('TimeSettings', 'AppPassword', fallback='')

       print("Loaded App Email:", self.app_email)
       print("Loaded App Password:", self.app_password)

   def save_settings(self):
       # Save settings to a configuration file
       config = configparser.ConfigParser()
       config['Admin'] = {'Username': self.admin_username, 'Password': self.admin_password}
       config['TimeSettings'] = {'LateTime': self.late_time.strftime("%H:%M:%S"),
                                 'ClosingTime': self.closing_time.strftime("%H:%M:%S"),
                                 'MaxDaysLate': str(self.max_days_late),
                                 'MaxDaysAbsent': str(self.max_days_absent)}
       config['AppSettings'] = {'AppEmail': self.app_email, 'AppPassword': self.app_password}

       with open('config.ini', 'w') as configfile:
           config.write(configfile)

   def schedule_log_absentees(self):
       current_day = self.get_current_day()
       # Check the current time every hour and call log_absentees if closing time is reached
       self.log_absentees(current_day)
       self.master.after(3600000, self.schedule_log_absentees)  # 3600000 milliseconds = 1 hour

   def create_widgets(self):
       # Load background logo with expansion
       self.background_logo = PhotoImage(file="top_left_logo.png")
       self.background_logo = self.background_logo.zoom(3)
       self.logo_label = tk.Label(self.master, image=self.background_logo)
       self.logo_label.place(x=0, y=0)  # Initially at (0, 0)

       # Create a label for the welcome message with a larger font size and starting on a new line
       self.welcome_label = tk.Label(self, text="", font=("Helvetica", 36, "bold"), fg="dark blue")
       self.welcome_label.pack(pady=20)

       # Set a flag to track whether the welcome animation is complete
       self.animation_complete = False

       # Create "Start Attendance" button after the animation is complete
       self.start_attendance_button = ttk.Button(self, text="Start Attendance", command=self.start_attendance,
                                                 style='TButton')
       # Do not pack the button initially
       # self.start_attendance_button.pack(pady=10)

       # Create "View Record" button
       self.view_record_button = ttk.Button(self, text="View Record", command=self.view_record, style='TButton')

       # Create "Time Management" button
       self.time_management_button = ttk.Button(self, text="Configurations", command=self.time_management,
                                                style='TButton')

       # Create "Data Management" button
       self.data_management_button = ttk.Button(self, text="Data Management", command=self.data_management,
                                                style='TButton')

       # Configure button style with hovering effects
       self.master.style = ttk.Style()
       self.master.style.configure('TButton',
                                   foreground='black',
                                   background='lightgray',
                                   font=('Helvetica', 14, 'bold'),
                                   padding=10)
       self.master.style.map('TButton',
                             foreground=[('active', 'green')],
                             background=[('active', 'lightgreen')])

   def animate_welcome_message(self):
       welcome_text = "Welcome to\nNDIC Corper/Intern Attendance Management System"
       for i in range(1, len(welcome_text) + 1):
           self.welcome_label.config(text=welcome_text[:i])
           self.master.update_idletasks()
           self.master.after(100)

       self.animation_complete = True
       self.start_attendance_button.pack(pady=10, side=tk.LEFT, anchor=tk.CENTER)  # Pack the "Start Attendance" button

       # Pack the "View Record" and "Data Management" buttons at the center
       self.view_record_button.pack(pady=10, side=tk.LEFT, padx=70)
       self.data_management_button.pack(pady=10, side=tk.RIGHT, padx=0)
       self.time_management_button.pack(pady=10, side=tk.TOP, anchor=tk.CENTER)

   def center_logo(self):
       # Center the logo based on the window size
       x = (self.master.winfo_width() - self.background_logo.width()) / 2
       y = (self.master.winfo_height() - self.background_logo.height()) / 2
       self.logo_label.place(x=x, y=y)

   def handle_resize(self, event):
       # Adjust the logo position when the window is resized
       self.center_logo()

   def get_current_day(self):
       # Get the current date and time
       current_datetime = datetime.now()

       # Extract and return the current day as a string
       return current_datetime.strftime("%A")

   def check_due_dates_and_remove_users(self):
       # Get the current date
       current_date = datetime.now().date()

       # Connect to the SQLite database in the current working directory
       conn = sqlite3.connect('enrollment_data.db')
       cursor = conn.cursor()

       # Check the corpers_table for due passing out dates
       cursor.execute('SELECT * FROM corpers')
       corpers = cursor.fetchall()

       # Assuming the passing_out_date is the fifth column in corpers_table
       passing_out_date_index = 4

       for corper in corpers:
           due_passing_out_date = datetime.strptime(corper[passing_out_date_index], "%Y-%m-%d").date()
           if due_passing_out_date <= current_date:
               # Delete the user from the corpers_table
               cursor.execute('DELETE FROM corpers WHERE id = ?', (corper[0],))  # Assuming id is the first column
               conn.commit()

               # Delete the user's image folder
               user_folder_path = os.path.join(os.getcwd(), corper[5])  # Assuming face_image_path is the sixth column
               if os.path.exists(user_folder_path):
                   shutil.rmtree(user_folder_path)

       # Check the interns_table for due school resumption dates
       cursor.execute('SELECT * FROM interns')
       interns = cursor.fetchall()

       # Assuming the school_resumption_date is the fourth column in interns_table
       school_resumption_date_index = 4

       for intern in interns:
           due_resumption_date_str = str(intern[school_resumption_date_index])  # Convert to string
           due_resumption_date = datetime.strptime(due_resumption_date_str, "%Y-%m-%d").date()

           if due_resumption_date <= current_date:
               # Delete the user from the interns_table
               cursor.execute('DELETE FROM interns WHERE id = ?', (intern[0],))  # Assuming id is the first column
               conn.commit()

               # Delete the user's image folder
               user_folder_path = os.path.join(os.getcwd(), intern[5])  # Assuming face_image_path is the fifth column
               if os.path.exists(user_folder_path):
                   shutil.rmtree(user_folder_path)

       # Close the database connection
       conn.close()
   def create_cds_csv(self):
       #Connect to the database
       conn = sqlite3.connect('enrollment_data.db')
       cursor = conn.cursor()

       # Get the current day
       current_day = self.get_current_day()

       # Select corpers who are on CDS on the current day
       cursor.execute('SELECT name FROM corpers WHERE cds_day = ?', (current_day,))
       cds_corpers = cursor.fetchall()

       # Close the database connection
       conn.close()

       # Create a CSV file with the names of corpers on CDS
       csv_filename = f'cds_corpers_{current_day}.csv'
       with open(csv_filename, 'w', newline='') as csvfile:
           fieldnames = ['Name']
           writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

           writer.writeheader()
           for corper in cds_corpers:
               writer.writerow({'Name': corper[0]})

       print(f"CSV file '{csv_filename}' created with corpers on CDS for {current_day}.")

   def say_attendance_status(self, attendance_status):
       text = f"{attendance_status}, thank you."
       if os.name == 'posix':  # macOS or Linux
           os.system(f'say "Thank You, {attendance_status}"')
       elif os.name == 'nt':  # Windows
           os.system(
               f'echo "Thank You, {attendance_status}" | powershell Add-Type -AssemblyName System.speech; $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer; $speak.Speak([Console]::In.ReadToEnd())')
       else:
           print("Text-to-speech is not supported on this platform.")

   def start_attendance(self):
       if self.animation_complete:
           print("Starting Attendance!")

           known_face_encodings, known_face_names = self.recognize_faces('enrollment_data.db')

           # Open a connection to the webcam (usually 0 or 1)
           cap = cv2.VideoCapture(0)

           # Get the current date in the format YYYY-MM-DD
           current_date = datetime.now().strftime("%Y-%m-%d")

           # Create the CSV file with the current date as the filename
           csv_filename = f'attendance_for:{current_date}.csv'
           fieldnames = ['Name', 'Time In', 'Time Out']

           # Check if the file exists, if not, write the header
           if not os.path.isfile(csv_filename):
               with open(csv_filename, mode='w', newline='') as csv_file:
                   writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                   writer.writeheader()

           while True:
               # Capture video frame-by-frame
               ret, frame = cap.read()

               # Check if the frame is valid
               if not ret:
                   print("Error: Unable to capture frame.")
                   break

               # Resize the frame for processing
               frame = cv2.resize(frame, (0, 0), fx=0.5, fy=0.5)

               # Find all face locations in the current frame
               face_locs = face_locations(np.array(frame))

               for top, right, bottom, left in face_locs:
                   # Draw a rectangle around the face
                   cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)

                   # Compute the face encoding
                   face_encoding = face_encodings(frame, [(top, right, bottom, left)])[0]

                   # Convert the frame to grayscale and apply histogram equalization
                   gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                   equalized_frame = cv2.equalizeHist(gray_frame)

                   # Adjust the distance threshold (tolerance) for face recognition
                   matches = face_recognition.compare_faces(known_face_encodings, face_encoding, tolerance=0.4)

                   # Draw a progress bar based on the confidence level
                   progress_bar_width = int(right - left)
                   confidence = face_recognition.face_distance(known_face_encodings, face_encoding)

                   # Check if there are any face matches
                   if confidence.size > 0:
                       confidence_percentage = int((1 - min(confidence)) * 100)
                   else:
                       confidence_percentage = 0

                   # Draw a light shaded bar
                   cv2.rectangle(frame, (left, bottom + 5), (right, bottom + 10), (200, 200, 200), -1)

                   # Draw a colored bar based on confidence level
                   color = (0, 255, 0)  # Green color for high confidence
                   if confidence_percentage < 65:
                       color = (0, 0, 255)  # Red color for low confidence

                   progress = int((confidence_percentage / 100) * progress_bar_width)
                   cv2.rectangle(frame, (left, bottom + 5), (left + progress, bottom + 10), color, -1)

                   # Draw a colored bar based on confidence level
                   color = (0, 255, 0)  # Green color for high confidence
                   if confidence_percentage < 65:
                       color = (0, 0, 255)  # Red color for low confidence

                   progress = int((confidence_percentage / 100) * progress_bar_width)
                   cv2.rectangle(frame, (left, bottom + 5), (left + progress, bottom + 10), color, -1)

                   # Check if any known face encoding matches the captured face encoding
                   if True in matches and confidence_percentage >= 65:  # Check confidence level
                       first_match_index = matches.index(True)
                       name = known_face_names[first_match_index]

                       # Get the current time
                       current_time = datetime.now().strftime("%H:%M:%S")

                       # Check if time in has already been recorded for the person on that day
                       if not self.is_attendance_recorded(csv_filename, name, 'Time In'):
                           # Log attendance in the CSV file and display attendance mark
                           time_in = current_time
                           self.log_attendance(csv_filename, name, time_in)
                           attendance_status = f"Clocked In for {name}"
                           if not current_time >= self.closing_time.strftime(
                                   "%H:%M:%S") and not self.is_attendance_recorded(
                                   csv_filename, name, 'Time Out'):
                               self.say_attendance_status(attendance_status)
                           # Check if user is clocking in late
                           current_time_obj = datetime.strptime(current_time, "%H:%M:%S").time()
                           if self.late_time <= current_time_obj < self.closing_time:
                               # Create a separate CSV file for late attendance
                               late_csv_filename = f'late_on_{datetime.now().strftime("%Y-%m-%d")}.csv'

                               # Check if the name already exists in the late entries
                               name_already_exists = False
                               if os.path.isfile(late_csv_filename):
                                   with open(late_csv_filename, mode='r') as late_csv_file:
                                       existing_late_reader = csv.DictReader(late_csv_file)
                                       for row in existing_late_reader:
                                           if row['Name'] == name:
                                               name_already_exists = True
                                               break

                               if not name_already_exists:
                                   # Check if the late CSV file exists, if not, write the header
                                   if not os.path.isfile(late_csv_filename):
                                       with open(late_csv_filename, mode='w', newline='') as late_csv_file:
                                           late_writer = csv.DictWriter(late_csv_file, fieldnames=['Name', 'Time In'])
                                           late_writer.writeheader()

                                   with open(late_csv_filename, mode='a', newline='') as late_csv_file:
                                       late_writer = csv.DictWriter(late_csv_file, fieldnames=['Name', 'Time In'])
                                       late_writer.writerow({'Name': name, 'Time In': time_in})
                                       print(f"Late attendance recorded for {name}.")
                       else:
                           attendance_status = f"{name} Has Clocked in already"
                           if not current_time >= self.closing_time.strftime(
                                   "%H:%M:%S") and not self.is_attendance_recorded(
                                   csv_filename, name, 'Time Out'):
                               self.say_attendance_status(attendance_status)

                       # Check if it's closing time or beyond and time out is not marked
                       if current_time >= self.closing_time.strftime("%H:%M:%S") and not self.is_attendance_recorded(
                               csv_filename, name, 'Time Out'):
                           self.log_attendance(csv_filename, name, '', current_time)
                           attendance_status = f"Clocked Out for {name}"
                           self.say_attendance_status(attendance_status)
                   else:
                       name = "Unknown"
                       attendance_status = ""

                   # Display the name and attendance status on the frame
                   cv2.putText(frame, name, (left + 6, bottom - 6), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
                   cv2.putText(frame, attendance_status, (left + 6, bottom + 20), cv2.FONT_HERSHEY_SIMPLEX, 0.5,
                               (0, 255, 0), 2)

               # Display the resulting frame
               cv2.imshow('Video', frame)

               # Break the loop if 'q' is pressed
               if cv2.waitKey(1) & 0xFF == ord('q'):
                   break

           # Release the webcam and close the window
           cap.release()
           cv2.destroyAllWindows()

   def is_attendance_recorded(self, csv_filename, name, column):
       with open(csv_filename, mode='r', newline='') as csv_file:
           reader = csv.DictReader(csv_file)
           for record in reader:
               if record['Name'] == name and record[column] and not record['Time Out']:
                   return True
       return False

   def log_attendance(self, csv_filename, name, time_in, time_out=''):
       # Get the current date in the format YYYY-MM-DD
       current_date = datetime.now().strftime("%Y-%m-%d")

       # Write attendance information to the main attendance table
       with open(csv_filename, mode='a', newline='') as csv_file:
           fieldnames = ['Name', 'Time In', 'Time Out']
           writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

           # Check if the person already has an entry for 'Time In'
           existing_records = []
           with open(csv_filename, mode='r', newline='') as csv_file:
               reader = csv.DictReader(csv_file)
               existing_records = [record['Name'] for record in reader]

           # If the person has 'Time In' recorded, update the existing record
           if name in existing_records:
               rows = []
               with open(csv_filename, mode='r', newline='') as csv_file:
                   reader = csv.DictReader(csv_file)
                   for record in reader:
                       if record['Name'] == name and record['Time In'] and not record['Time Out']:
                           record['Time Out'] = time_out
                       rows.append(record)

               with open(csv_filename, mode='w', newline='') as csv_file:
                   writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                   writer.writeheader()
                   writer.writerows(rows)

           else:
               writer.writerow(
                   {'Name': name, 'Time In': time_in, 'Time Out': time_out})

   def get_all_user_names(self):
       # Connect to the SQLite database in the current working directory
       conn = sqlite3.connect('enrollment_data.db')
       cursor = conn.cursor()

       # Retrieve only the names from both corpers and interns tables
       cursor.execute('SELECT name FROM corpers UNION SELECT name FROM interns')
       user_names = [row[0] for row in cursor.fetchall()]

       # Close the database connection
       conn.close()

       return user_names

   def get_all_users_emails_from_database(self, database_path):
       try:
           connection = sqlite3.connect(database_path)
           cursor = connection.cursor()

           # Query to select names and email addresses from the 'corpers' table
           cursor.execute("SELECT name, email FROM corpers")
           corpers_rows = cursor.fetchall()

           # Query to select names and email addresses from the 'interns' table
           cursor.execute("SELECT name, email FROM interns")
           interns_rows = cursor.fetchall()

           # Extract names and email addresses from the 'corpers' and 'interns' rows
           corpers_data = [(row[0], row[1]) for row in corpers_rows]
           interns_data = [(row[0], row[1]) for row in interns_rows]

           # Combine corpers and interns data
           combined_data = corpers_data + interns_data

           return combined_data

       except sqlite3.Error as e:
           print("Error reading data from the database:", e)
           return []

   def log_absentees(self, current_day):
       current_time = datetime.now().time()
       current_date = datetime.now()
       current_day = current_date.strftime("%A")
       current_date_str = current_date.strftime("%Y-%m-%d")

       # Check if absence query has been sent for the current date
       query_sent = self.check_query_sent_status(current_date_str)

       if not query_sent and current_day != "Saturday" and current_day != "Sunday" and current_time >= self.closing_time:
           all_user_names = self.get_all_user_names()
           present_users = self.get_present_users()
           exempted_users = self.get_exempted_users()
           cds_corpers = self.get_cds_corpers(current_day) or []

           absentees = [user for user in all_user_names if
                        user not in present_users and user not in exempted_users and user not in cds_corpers]

           if absentees:
               csv_filename = f'absent_on_{current_date_str}.csv'
               with open(csv_filename, mode='w', newline='') as csvfile:
                   fieldnames = ['Name']
                   writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                   writer.writeheader()
                   for absentee in absentees:
                       writer.writerow({'Name': absentee})

               print(f"CSV file '{csv_filename}' created with absentees for {current_date_str}.")

               # Send query email to absentees
               subject = "Absence From Duty"
               body = "Dear {name},\nOur Attendance record shows that you were absent from work on: {current_date}. Please reply to this email with the reason for your absence.\n\nRegards,\nThe Attendance Management Team."

               sender_email = f'{self.app_email}'
               sender_password = f'{self.app_password}'

               # Filter recipient data to only include absentees
               recipient_data = [(name, email) for name, email in
                                 self.get_all_users_emails_from_database('enrollment_data.db') if name in absentees]

               for name, email in recipient_data:
                   self.send_absent_query(sender_email, sender_password, email, subject,
                                          body.format(name=name, current_date=current_date_str))

               # Update absence query report after sending queries
               self.update_query_report(current_date_str, "Sent")
           else:
               print("No absentees for today.")
       else:
           print("Absence query already sent for today.")

   def update_query_report(self, current_date_str, status):
       report_filename = "absent_query_report.csv"
       if not os.path.exists(report_filename):
           with open(report_filename, mode='w', newline='') as report_file:
               fieldnames = ['date', 'status']
               writer = csv.DictWriter(report_file, fieldnames=fieldnames)
               writer.writeheader()

       with open(report_filename, mode='a', newline='') as report_file:
           writer = csv.writer(report_file)
           writer.writerow([current_date_str, status])

   def check_query_sent_status(self, current_date_str):
       report_filename = "absent_query_report.csv"
       if not os.path.exists(report_filename):
           return False

       with open(report_filename, mode='r') as report_file:
           reader = csv.DictReader(report_file)
           for row in reader:
               if row['date'] == current_date_str and row['status'] == "Sent":
                   return True
       return False

   def send_absent_query(self, sender_email, sender_password, recipient_email, subject, body):
       try:
           msg = MIMEMultipart()
           msg['From'] = sender_email
           msg['To'] = recipient_email
           msg['Subject'] = subject
           msg.attach(MIMEText(body, 'plain'))

           server = smtplib.SMTP('smtp.gmail.com', 587)
           server.starttls()
           server.login(sender_email, sender_password)
           server.sendmail(sender_email, recipient_email, msg.as_string())
           server.quit()

           print(f"Email sent to {recipient_email} successfully.")
       except Exception as e:
           print(f"Failed to send email to {recipient_email}: {e}")

   def get_cds_corpers(self, current_day):
       try:
           # CSV filename based on the current day
           csv_filename = f'cds_corpers_{current_day}.csv'

           # Check if the CSV file exists
           if not os.path.isfile(csv_filename):
               print(f"No CSV file found for CDS corpers on {current_day}.")
               return []

           # Read corpers from the CSV file
           with open(csv_filename, mode='r', newline='') as csvfile:
               reader = csv.DictReader(csvfile)
               cds_corpers = [row['Name'] for row in reader]

           print("CDS Corpers:", cds_corpers)
           return cds_corpers

       except Exception as e:
           print("Error reading CDS corpers CSV:", e)
           return []

   def get_exempted_users(self):
       try:
           # CSV filename for exemptions
           csv_filename = 'exemptions.csv'

           # Check if the CSV file exists
           if not os.path.isfile(csv_filename):
               print("No exemptions CSV file found.")
               return []

           # Read exemptions from the CSV file
           with open(csv_filename, mode='r', newline='') as csvfile:
               reader = csv.DictReader(csvfile)
               exempted_users = [row['Name'] for row in reader]

           print("Exempted Users:", exempted_users)
           return exempted_users

       except Exception as e:
           print("Error reading exemptions CSV:", e)
           return []

   def get_present_users(self):
       # Get the current date in the format YYYY-MM-DD
       current_date = datetime.now().strftime("%Y-%m-%d")

       # Create the CSV file with the current date as the filename
       csv_filename = f'attendance_for:{current_date}.csv'

       # Check if the file exists
       if os.path.isfile(csv_filename):
           # Read the names of users who have marked attendance
           with open(csv_filename, mode='r', newline='') as csv_file:
               reader = csv.DictReader(csv_file)
               present_users = [record['Name'] for record in reader]

           return present_users
       else:
           return []

   def is_valid_exemption(self, exemption, current_date):
       # Parse the end date from the exemption
       end_date = datetime.strptime(exemption['End Date'], "%Y-%m-%d").date()

       # Check if the end date has passed
       return end_date >= current_date

   def view_record(self):
       # Create a Tkinter root window to host the dialog
       root = tk.Tk()
       root.withdraw()  # Hide the main window

       # Create a simple dialog for username and password
       class LoginDialog(simpledialog.Dialog):
           def __init__(self, parent, title):
               self.admin_username = parent.admin_username
               self.admin_password = parent.admin_password
               super().__init__(parent, title)

           def body(self, master):
               tk.Label(master, text="Admin Username:").grid(row=0, column=0)
               tk.Label(master, text="Admin Password:").grid(row=1, column=0)

               self.username_entry = tk.Entry(master)
               self.username_entry.grid(row=0, column=1)

               self.password_entry = tk.Entry(master, show='*')
               self.password_entry.grid(row=1, column=1)

               return self.username_entry  # Set the initial focus to the username entry

           def apply(self):
               username = self.username_entry.get()
               password = self.password_entry.get()

               if username == self.admin_username and password == self.admin_password:
                   self.result = (username, password)
               else:
                   messagebox.showerror("Invalid Credentials", "Invalid username or password.")
                   self.result = None

       # Create an instance of the login dialog
       login_dialog = LoginDialog(self, "Admin Login")

       # Check if the user pressed cancel or provided invalid credentials
       if login_dialog.result is None or login_dialog.result == ('', ''):
           return

       # Extract username and password from the result
       username, password = login_dialog.result

       # Get the current directory
       current_directory = os.getcwd()

       # Get all CSV files in the current directory
       csv_files = [file for file in os.listdir(current_directory) if file.endswith(".csv")]

       if not csv_files:
           messagebox.showinfo("No Records", "No CSV files found.")
           return

       # Ask the user to select a CSV file
       selected_file = filedialog.askopenfilename(
           initialdir=current_directory,
           title="Select CSV File",
           filetypes=[("CSV Files", "*.csv")])

       # Check if the user selected a file
       if selected_file:
           # Use subprocess to open the selected CSV file with the default application
           subprocess.run(["open", selected_file], check=True)

   def time_management(self):
       # Create a Tkinter root window to host the dialog
       root = tk.Tk()
       root.withdraw()  # Hide the main window

       # Create a simple dialog for username and password
       class LoginDialog(simpledialog.Dialog):
           def __init__(self, parent, title):
               self.admin_username = parent.admin_username
               self.admin_password = parent.admin_password
               super().__init__(parent, title)

           def body(self, master):
               tk.Label(master, text="Admin Username:").grid(row=0, column=0)
               tk.Label(master, text="Admin Password:").grid(row=1, column=0)

               self.username_entry = tk.Entry(master)
               self.username_entry.grid(row=0, column=1)

               self.password_entry = tk.Entry(master, show='*')
               self.password_entry.grid(row=1, column=1)

               return self.username_entry  # Set the initial focus to the username entry

           def apply(self):
               username = self.username_entry.get()
               password = self.password_entry.get()

               if username == self.admin_username and password == self.admin_password:
                   self.result = (username, password)
               else:
                   messagebox.showerror("Invalid Credentials", "Invalid username or password.")
                   self.result = None

       # Create an instance of the login dialog
       login_dialog = LoginDialog(self, "Admin Login")

       # Check if the user pressed cancel or provided invalid credentials
       if login_dialog.result is None or login_dialog.result == ('', ''):
           return
       print("Admin login successful. Access granted to time management.")
       # Update late_time and closing_time as needed

       # Create a new top-level window for time management
       time_window = tk.Toplevel(self.master)
       time_window.title("Time Management")

       # Create entry widgets for late time and closing time
       tk.Label(time_window, text="Late Time (HH:MM:SS):").grid(row=0, column=0, padx=10, pady=5)
       late_time_entry = tk.Entry(time_window)
       late_time_entry.grid(row=0, column=1, padx=10, pady=5)

       tk.Label(time_window, text="Closing Time (HH:MM:SS):").grid(row=1, column=0, padx=10, pady=5)
       closing_time_entry = tk.Entry(time_window)
       closing_time_entry.grid(row=1, column=1, padx=10, pady=5)

       # Set default values in the entry widgets
       late_time_entry.insert(0, self.late_time.strftime("%H:%M:%S"))
       closing_time_entry.insert(0, self.closing_time.strftime("%H:%M:%S"))

       # Create dropdowns for max days late and max days absent
       tk.Label(time_window, text="Max Days Late:").grid(row=2, column=0, padx=10, pady=5)
       max_days_late_var = tk.StringVar(time_window)
       max_days_late_var.set(str(self.max_days_late))  # Default value from settings
       max_days_late_dropdown = tk.OptionMenu(time_window, max_days_late_var, *range(1, 11))
       max_days_late_dropdown.config(width=16)  # Set width explicitly
       max_days_late_dropdown.grid(row=2, column=1, padx=10, pady=5)

       tk.Label(time_window, text="Max Days Absent:").grid(row=3, column=0, padx=10, pady=5)
       max_days_absent_var = tk.StringVar(time_window)
       max_days_absent_var.set(str(self.max_days_absent))  # Default value from settings
       max_days_absent_dropdown = tk.OptionMenu(time_window, max_days_absent_var, *range(1, 11))
       max_days_absent_dropdown.config(width=16)  # Set width explicitly
       max_days_absent_dropdown.grid(row=3, column=1, padx=10, pady=5)

       # Load App Email and App Password from the configuration file
       config = configparser.ConfigParser()
       config.read('config.ini')
       app_email = config.get('TimeSettings', 'AppEmail', fallback='')
       app_password = config.get('TimeSettings', 'AppPassword', fallback='')

       # Create labels for App Email and App Password
       app_email_label = tk.Label(time_window, text="App Email:", anchor="e")
       app_email_label.grid(row=4, column=0, padx=10, pady=5)

       app_password_label = tk.Label(time_window, text="App Password:", anchor="e")
       app_password_label.grid(row=5, column=0, padx=10, pady=5)

       # Create entry widgets for App Email and App Password
       app_email_entry = tk.Entry(time_window, width=20)
       app_email_entry.grid(row=4, column=1, padx=10, pady=5)
       app_email_entry.insert(0, app_email)  # Insert stored email

       app_password_entry = tk.Entry(time_window, width=20, show='')
       app_password_entry.grid(row=5, column=1, padx=10, pady=5)
       app_password_entry.insert(0, app_password)  # Insert stored password

       # Function to automatically add spaces to the password as the user types
       def add_spaces(event):
           # Get current cursor position
           cursor_pos = app_password_entry.index(tk.INSERT)
           # Insert a space every 4 characters
           if cursor_pos % 5 == 4:
               app_password_entry.insert(tk.INSERT, ' ')

       app_password_entry.bind('<Key>', add_spaces)  # Bind the function to the Key event

       # Create a button to update the times, settings, email, and password
       update_button = tk.Button(time_window, text="Save Times, Settings, Email, and Password",
                                 command=lambda: self.save_times_and_settings(late_time_entry, closing_time_entry,
                                                                              max_days_late_var, max_days_absent_var,
                                                                              time_window, app_email_entry,
                                                                              app_password_entry))
       update_button.grid(row=6, columnspan=2, pady=10)

       # Add a button to upload clearance template
       upload_template_button = tk.Button(time_window, text="Upload Clearance Template", command=self.upload_template)
       upload_template_button.grid(row=7, columnspan=2, pady=10)

   def save_times_and_settings(self, late_time_entry, closing_time_entry, max_days_late_var, max_days_absent_var,
                               time_window, app_email_entry, app_password_entry):
       try:
           late_time_str = late_time_entry.get()
           closing_time_str = closing_time_entry.get()
           max_days_late = int(max_days_late_var.get())
           max_days_absent = int(max_days_absent_var.get())
           app_email = app_email_entry.get()
           app_password = app_password_entry.get()

           # Update the instance variables directly
           self.late_time = datetime.strptime(late_time_str, "%H:%M:%S").time()
           self.closing_time = datetime.strptime(closing_time_str, "%H:%M:%S").time()
           self.max_days_late = max_days_late
           self.max_days_absent = max_days_absent

           # Save settings to a configuration file
           config = configparser.ConfigParser()
           config['Admin'] = {'Username': self.admin_username, 'Password': self.admin_password}
           config['TimeSettings'] = {'LateTime': self.late_time.strftime("%H:%M:%S"),
                                     'ClosingTime': self.closing_time.strftime("%H:%M:%S"),
                                     'MaxDaysLate': str(self.max_days_late),
                                     'MaxDaysAbsent': str(self.max_days_absent),
                                     'AppEmail': app_email,
                                     'AppPassword': app_password}

           with open('config.ini', 'w') as configfile:
               config.write(configfile)

           messagebox.showerror("Success",
                                f"Late Time updated to: {self.late_time}\n"
                                f"Closing Time updated to: {self.closing_time}\n"
                                f"Max Days Late updated to: {self.max_days_late}\n"
                                f"Max Days Absent updated to: {self.max_days_absent}\n"
                                f"App Email updated to: {app_email}\n"
                                f"App Password updated.")

           # Print or perform any other necessary actions with the updated times, settings, email, and password
           print(f"Late Time updated to: {self.late_time}")
           print(f"Closing Time updated to: {self.closing_time}")
           print(f"Max Days Late updated to: {self.max_days_late}")
           print(f"Max Days Absent updated to: {self.max_days_absent}")
           print(f"App Email updated to: {app_email}")
           print("App Password updated.")

           # Close the time management window
           time_window.destroy()
       except ValueError:
           messagebox.showerror("Invalid Input",
                                "Please enter valid time formats (HH:MM:SS) and integer values for max days late and max days absent.")

   def upload_template(self):
       # Open a file dialog to select the clearance template file
       file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx"), ("All files", "*.*")])

       # Process the uploaded template file
       if file_path:
           # Check if the selected file is a .docx file
           if file_path.endswith('.docx'):
               # Create a folder named "templates" if it doesn't exist
               templates_folder = os.path.join(os.getcwd(), "templates")
               if not os.path.exists(templates_folder):
                   os.makedirs(templates_folder)

               # Save the uploaded file as "clearance_template.docx" in the "templates" folder
               destination_file = os.path.join(templates_folder, "clearance_template.docx")
               shutil.copyfile(file_path, destination_file)

               print("Template file uploaded and saved as 'clearance_template.docx' in 'templates' folder.")
               messagebox.showinfo("Success", "Clearance template uploaded and saved successfully.")
           else:
               print("Invalid file type. Please select a .docx file.")
               messagebox.showerror("Error", "Invalid file type. Please select a .docx file.")

   def get_all_users(self):
       # Connect to the SQLite database in the current working directory
       conn = sqlite3.connect('enrollment_data.db')
       cursor = conn.cursor()

       # Retrieve all columns from both corpers and interns tables
       cursor.execute('SELECT * FROM corpers UNION SELECT * FROM interns')
       user_details = cursor.fetchall()

       # Close the database connection
       conn.close()

       return user_details

   def data_management(self):
       # Create a Tkinter root window to host the dialog
       root = tk.Tk()
       root.withdraw()  # Hide the main window

       # Create a simple dialog for username and password
       class LoginDialog(simpledialog.Dialog):
           def __init__(self, parent, title):
               self.admin_username = parent.admin_username
               self.admin_password = parent.admin_password
               super().__init__(parent, title)

           def body(self, master):
               tk.Label(master, text="Admin Username:").grid(row=0, column=0)
               tk.Label(master, text="Admin Password:").grid(row=1, column=0)

               self.username_entry = tk.Entry(master)
               self.username_entry.grid(row=0, column=1)

               self.password_entry = tk.Entry(master, show='*')
               self.password_entry.grid(row=1, column=1)

               return self.username_entry  # Set the initial focus to the username entry

           def apply(self):
               username = self.username_entry.get()
               password = self.password_entry.get()

               if username == self.admin_username and password == self.admin_password:
                   self.result = (username, password)
               else:
                   messagebox.showerror("Invalid Credentials", "Invalid username or password.")
                   self.result = None

       # Create an instance of the login dialog
       login_dialog = LoginDialog(self, "Admin Login")

       # Check if the user pressed cancel or provided invalid credentials
       if login_dialog.result is None or login_dialog.result == ('', ''):
           return
       print("Admin login successful. Access granted to time management.")

       # Retrieve user details from the database
       users = self.get_all_users()

       # Create a new top-level window for data management
       data_management_window = tk.Toplevel(self.master)
       data_management_window.title("Data Management")

       # Create a listbox to display user details
       user_listbox = tk.Listbox(data_management_window, width=120)
       for user in users:
           user_listbox.insert(tk.END, user)

       # Pack the listbox
       user_listbox.pack(pady=10)

       # Create "Give Exemption" button
       # Create "Give Exemption" button
       give_exemption_button = ttk.Button(data_management_window, text="Exemptions",
                                          command=lambda: self.give_exemption(data_management_window))
       give_exemption_button.pack(side=tk.LEFT, padx=5)

       # Create "Delete User" button
       delete_user_button = ttk.Button(data_management_window, text="Delete User",
                                       command=lambda: self.delete_user(user_listbox.get(tk.ACTIVE)))
       delete_user_button.pack(side=tk.LEFT, padx=5)

       # Create "Generate Report" button
       generate_report_button = ttk.Button(data_management_window, text="Generate Report",
                                           command=lambda: self.generate_report(data_management_window))
       generate_report_button.pack(side=tk.LEFT, padx=5)

       # Create "Change Admin Credentials" button
       change_credentials_button = ttk.Button(data_management_window, text="Change Admin Credentials",
                                              command=self.change_admin_credentials)
       change_credentials_button.pack(side=tk.LEFT, padx=5)

   def get_all_names_only(self):
       # Connect to the SQLite database in the current working directory
       conn = sqlite3.connect('enrollment_data.db')
       cursor = conn.cursor()

       # Retrieve only the 'Name' column from both corpers and interns tables
       cursor.execute('SELECT Name FROM corpers UNION SELECT Name FROM interns')
       names_only = cursor.fetchall()

       # Close the database connection
       conn.close()

       # Extract names from the list of tuples
       names = [name[0] for name in names_only]

       return names

   def give_exemption(self, data_management_window):
       # Create a frame within the data_management_window for the exemption dialog
       exemption_frame = ttk.Frame(data_management_window)
       exemption_frame.pack()

       # Create and grid labels and entry widgets for user name, start date, and end date
       tk.Label(exemption_frame, text="User Name:").grid(row=0, column=0)

       # Get all names for the drop-down list
       all_names = self.get_all_names_only()

       # Create a dropdown list for user names
       user_name_var = tk.StringVar(exemption_frame)
       user_name_dropdown = ttk.Combobox(exemption_frame, textvariable=user_name_var, values=all_names)
       user_name_dropdown.grid(row=0, column=1)

       tk.Label(exemption_frame, text="Start Date:").grid(row=1, column=0)
       start_date_entry = DateEntry(exemption_frame, width=12, background='darkblue',
                                    foreground='white', borderwidth=2, date_pattern="yyyy-mm-dd")
       start_date_entry.grid(row=1, column=1)

       tk.Label(exemption_frame, text="End Date:").grid(row=2, column=0)
       end_date_entry = DateEntry(exemption_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern="yyyy-mm-dd")
       end_date_entry.grid(row=2, column=1)

       # Create a button to log exemption details
       log_button = ttk.Button(exemption_frame, text="Log Exemption",
                               command=lambda: self.log_exemption(user_name_var.get(),
                                                                   start_date_entry.get(),
                                                                   end_date_entry.get(),
                                                                   exemption_frame))
       log_button.grid(row=3, columnspan=2, pady=10)

   def log_exemption(self, user_name, start_date, end_date, frame):
       # Check if the user provided input for all fields
       if user_name and start_date and end_date:
           # Log exemption details to a CSV file
           csv_filename = 'exemptions.csv'  # Provide the desired CSV file name
           fieldnames = ['Name', 'Start Date', 'End Date']

           with open(csv_filename, mode='a', newline='') as csvfile:
               writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

               # Write header if the file is empty
               if csvfile.tell() == 0:
                   writer.writeheader()

               # Write exemption details
               writer.writerow({'Name': user_name, 'Start Date': start_date, 'End Date': end_date})

           messagebox.showinfo("Exemption Success", f"Exemption details for '{user_name}' logged successfully.")
           frame.destroy()  # Destroy the exemption_frame after logging
       else:
           messagebox.showwarning("Exemption Canceled", "Exemption process canceled. Please provide valid input.")

   def check_expired_exemptions(self):
       # Get the current date
       current_date = datetime.now().date()

       # Set the CSV filename
       csv_filename = 'exemptions.csv'  # Update with your actual CSV filename

       # Check if the CSV file exists
       if os.path.isfile(csv_filename):
           # Read the exemptions from the CSV file
           with open(csv_filename, mode='r', newline='') as csvfile:
               reader = csv.DictReader(csvfile)
               exemptions = list(reader)

           # Check and remove expired exemptions
           updated_exemptions = []
           for exemption in exemptions:
               end_date = datetime.strptime(exemption['End Date'], "%Y-%m-%d").date()
               if end_date >= current_date:
                   updated_exemptions.append(exemption)

           # Update the CSV file with non-expired exemptions
           with open(csv_filename, mode='w', newline='') as csvfile:
               fieldnames = ['Name', 'Start Date', 'End Date']
               writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
               writer.writeheader()
               writer.writerows(updated_exemptions)
       else:
           print(f"CSV file '{csv_filename}' does not exist. No action taken.")


   def delete_user(self, username):
       # Prompt the user for confirmation and get the name to be deleted
       user_to_delete = simpledialog.askstring("Delete User", f"Enter the name to delete ({username}):")

       if user_to_delete:
           # Connect to the SQLite database
           conn = sqlite3.connect('enrollment_data.db')
           cursor = conn.cursor()

           try:
               # Delete the user from the database
               cursor.execute('DELETE FROM corpers WHERE name = ?', (user_to_delete,))
               cursor.execute('DELETE FROM interns WHERE name = ?', (user_to_delete,))
               conn.commit()

               # Delete the user's image folder
               image_folder_path = f"./{user_to_delete}"  # Provide the actual path to the images folder
               shutil.rmtree(image_folder_path, ignore_errors=True)
               messagebox.showinfo("success", f"User '{user_to_delete}' deleted successfully.")

           except sqlite3.Error as e:
               conn.rollback()
               messagebox.showinfo("Error", "Error deleting user from the database:", e)

           finally:
               # Close the database connection
               conn.close()
       else:
           print("Deletion canceled.")

   def generate_report(self, data_management_window):
       # Create a frame within the data_management_window for the report dialog
       report_frame = ttk.Frame(data_management_window)
       report_frame.pack()

       # Create and grid labels and entry widgets for FROM date and TO date
       tk.Label(report_frame, text="FROM Date:").grid(row=0, column=0)
       from_date_entry = DateEntry(report_frame, width=12, background='darkblue',
                                   foreground='white', borderwidth=2, date_pattern="yyyy-mm-dd")
       from_date_entry.grid(row=0, column=1)

       tk.Label(report_frame, text="TO Date:").grid(row=1, column=0)
       to_date_entry = DateEntry(report_frame, width=12, background='darkblue',
                                 foreground='white', borderwidth=2, date_pattern="yyyy-mm-dd")
       to_date_entry.grid(row=1, column=1)

       # Create a button to generate the report
       generate_button = ttk.Button(report_frame, text="Generate Report",
                                    command=lambda: self.generate_report_logic(from_date_entry.get(),
                                                                               to_date_entry.get(),
                                                                               report_frame))
       generate_button.grid(row=2, columnspan=2, pady=10)

   def generate_report_logic(self, from_date, to_date, frame):
       try:
           # Parse the input dates
           from_date = datetime.strptime(from_date, "%Y-%m-%d").date()
           to_date = datetime.strptime(to_date, "%Y-%m-%d").date()

           # Generate a list of dates within the specified range
           date_range = [from_date + timedelta(days=x) for x in range((to_date - from_date).days + 1)]

           # Create the CSV file for the report
           report_csv_filename = f'Report_for_{from_date}_{to_date}.csv'
           with open(report_csv_filename, mode='w', newline='') as report_csv:
               fieldnames = ['Names', 'Days Present', 'Days Absent', 'Days Late']
               writer = csv.DictWriter(report_csv, fieldnames=fieldnames)
               writer.writeheader()

               # Iterate through each user
               all_user_names = self.get_all_user_names()  # Replace with your method to get user names
               for user_name in all_user_names:
                   days_present = 0
                   days_absent = 0
                   days_late = 0

                   # Iterate through each date in the range
                   for current_date in date_range:
                       # Construct the CSV file names for attendance, late, and absent records
                       attendance_csv_filename = f'attendance_for:{current_date}.csv'
                       late_csv_filename = f'late_on_{current_date}.csv'
                       absent_csv_filename = f'absent_on_{current_date}.csv'

                       # Read and process the attendance CSV file
                       attendance_records = self.read_csv(attendance_csv_filename)
                       # Check if the user was present
                       if any(record['Name'] == user_name for record in attendance_records):
                           days_present += 1

                       # Read and process the late CSV file
                       late_records = self.read_csv(late_csv_filename)
                       # Check if the user was late
                       if any(record['Name'] == user_name for record in late_records):
                           days_late += 1

                       # Read and process the absent CSV file
                       absent_records = self.read_csv(absent_csv_filename)
                       # Check if the user was absent
                       if any(record['Name'] == user_name for record in absent_records):
                           days_absent += 1

                   # Write user details to the report CSV file
                   writer.writerow({'Names': user_name, 'Days Present': days_present,
                                    'Days Absent': days_absent, 'Days Late': days_late})

           result = messagebox.askquestion("Report Generated",
                                           f"Report has been generated at {report_csv_filename}. Do you want to open it?",
                                           icon='info')

           if result == 'yes':
               # Open the report using subprocess
               subprocess.run(['open', report_csv_filename], check=True)
           else:
               # Display a message box indicating that the report is saved
               messagebox.showinfo("Report Saved", f"Report has been saved at {report_csv_filename}.")

           print("Process complete.")

       except ValueError:
           messagebox.showerror("Invalid Input", "Invalid date format. Please use YYYY-MM-DD.")

       frame.destroy()  # Destroy the report_frame after generating the report

   def read_csv(self, csv_filename):
       # Read and return the contents of a CSV file
       records = []
       if os.path.isfile(csv_filename):
           with open(csv_filename, mode='r', newline='') as csv_file:
               reader = csv.DictReader(csv_file)
               records = list(reader)
       return records

   def change_admin_credentials(self):
       # Create a Tkinter root window to host the dialog
       root = tk.Tk()
       root.withdraw()  # Hide the main window

       # Create a simple dialog for new admin credentials
       class ChangeCredentialsDialog(simpledialog.Dialog):
           def body(self, master):
               tk.Label(master, text="Enter New Admin Username:").grid(row=0, column=0)
               tk.Label(master, text="Enter New Admin Password:").grid(row=1, column=0)

               self.new_username_entry = tk.Entry(master)
               self.new_username_entry.grid(row=0, column=1)

               self.new_password_entry = tk.Entry(master, show='*')
               self.new_password_entry.grid(row=1, column=1)

               return self.new_username_entry  # Set the initial focus to the username entry

           def apply(self):
               new_username = self.new_username_entry.get()
               new_password = self.new_password_entry.get()

               if new_username and new_password:  # Check if both fields are filled
                   self.result = (new_username, new_password)
               else:
                   messagebox.showerror("Invalid Input", "Please enter both username and password.")
                   self.result = None

       # Create an instance of the credentials change dialog
       credentials_dialog = ChangeCredentialsDialog(root, "Change Admin Credentials")

       # Check if the user pressed cancel or provided invalid input
       if credentials_dialog.result is None:
           return

       # Extract new username and password from the result
       new_username, new_password = credentials_dialog.result

       # Update admin credentials
       self.admin_username = new_username
       self.admin_password = new_password

       # Save the updated credentials (implement this method as needed)
       self.save_settings()

       messagebox.showinfo("Credentials Updated", "Admin credentials updated successfully.")

   def capture_faces(self):
       # Load face recognition models
       known_face_encodings, known_face_names = self.recognize_faces('enrollment_data.db')

       # Open webcam
       cap = cv2.VideoCapture(0)

       while True:
           ret, frame = cap.read()

           # Convert the image from BGR color (which OpenCV uses) to RGB color
           rgb_frame = frame[:, :, ::-1]

           # Find all face locations and face encodings in the current frame of video
           face_locs = face_recognition.face_locations(rgb_frame)
           if len(face_locs) > 0:
               # Only use the first face found in the frame
               top, right, bottom, left = face_locs[0]
               face_encoding = face_recognition.face_encodings(rgb_frame, [(top, right, bottom, left)])[0]

               # Print face locations
               print("Face Locations:", face_locs)

               # Print face encoding
               print("Face Encoding:", face_encoding)

               # Compare the face encoding with known face encodings
               matches = face_recognition.compare_faces(known_face_encodings, face_encoding)

               # Print the result of the comparison
               print("Matches:", matches)

               name = "Unknown"

               # If a match is found, use the name associated with the known face
               if True in matches:
                   first_match_index = matches.index(True)
                   name = known_face_names[first_match_index]
                   print(f"Match found: {name}")
               else:
                   print("Unknown")

               # Draw a rectangle around the face
               cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)

               # Display the name on the frame
               cv2.putText(frame, name, (left + 6, bottom - 6), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)

           # Display the resulting image
           cv2.imshow('Video', frame)

           # If the 'q' key is pressed, break from the loop
           if cv2.waitKey(1) & 0xFF == ord('q'):
               break

       # Release the webcam
       cap.release()
       cv2.destroyAllWindows()

   def get_previous_month_dates(self):
       # Get today's date
       today = datetime.today()

       # Calculate the first day of the current month
       first_day_of_current_month = today.replace(day=1)

       # Calculate the first day of the previous month
       first_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
       first_day_of_previous_month = first_day_of_previous_month.replace(day=1)

       # Calculate the last day of the previous month
       last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)

       # Get the name of the previous month
       previous_month_name = first_day_of_previous_month.strftime("%B")

       return str(first_day_of_previous_month.strftime("%Y-%m-%d")), str(
           last_day_of_previous_month.strftime("%Y-%m-%d")), previous_month_name

   def auto_generate_report_logic(self, from_date, to_date, month):
       today = datetime.today()
       month_ = datetime.strptime(from_date, "%Y-%m-%d")
       month_ = month_.strftime("%B")
       _year = datetime.strptime(from_date, "%Y-%m-%d")
       _year = _year.strftime("%Y")
       file_name = f"Report_for_{month_}_{_year}.csv"

       # Check if today is 1st, 2nd, or 3rd and file does not exist
       if today.day in [1, 2, 3, 4, 5] and not os.path.exists(file_name):
           try:
               # Parse the input dates
               from_date = datetime.strptime(from_date, "%Y-%m-%d").date()
               to_date = datetime.strptime(to_date, "%Y-%m-%d").date()
               month = month + datetime.now().strftime("_%Y")
               # Generate a list of dates within the specified range
               date_range = [from_date + timedelta(days=x) for x in range((to_date - from_date).days + 1)]

               # Create the CSV file for the report
               report_csv_filename = f'Report_for_{month}.csv'
               with open(report_csv_filename, mode='w', newline='') as report_csv:
                   fieldnames = ['Names', 'Days Present', 'Days Absent', 'Days Late']
                   writer = csv.DictWriter(report_csv, fieldnames=fieldnames)
                   writer.writeheader()

                   # Iterate through each user
                   # Iterate through each user
                   all_user_names = self.get_all_user_names()  # Replace with your method to get user names
                   for user_name in all_user_names:
                       days_present = 0
                       days_absent = 0
                       days_late = 0

                       # Iterate through each date in the range
                       for current_date in date_range:
                           # Construct the CSV file names for attendance, late, and absent records
                           attendance_csv_filename = f'attendance_for:{current_date}.csv'
                           late_csv_filename = f'late_on_{current_date}.csv'
                           absent_csv_filename = f'absent_on_{current_date}.csv'

                           # Read and process the attendance CSV file
                           attendance_records = self.read_csv(attendance_csv_filename)
                           # Check if the user was present
                           if any(record['Name'] == user_name for record in attendance_records):
                               days_present += 1

                           # Read and process the late CSV file
                           late_records = self.read_csv(late_csv_filename)
                           # Check if the user was late
                           if any(record['Name'] == user_name for record in late_records):
                               days_late += 1

                           # Read and process the absent CSV file
                           absent_records = self.read_csv(absent_csv_filename)
                           # Check if the user was absent
                           if any(record['Name'] == user_name for record in absent_records):
                               days_absent += 1

                       # Write user details to the report CSV file
                       writer.writerow({'Names': user_name, 'Days Present': days_present,
                                        'Days Absent': days_absent, 'Days Late': days_late})

               print("Process complete.")
               self.send_late_query_emails()
           except ValueError:
               print("Invalid Input", "Invalid date format. Please use YYYY-MM-DD.")
       print("exist")

   def schedule_monthly_report(self):
       from_date, to_date, month = self.get_previous_month_dates()

       # Call function to perform scheduled task
       self.auto_generate_report_logic(from_date, to_date, month)
       self.auto_generate_clearance_file()
       self.process_word_document("template.docx")

       self.master.after(86400000, self.schedule_monthly_report)

   def collect_names_and_state_codes(self):
       # Connect to the database
       connection = self.connect_to_database()

       # Create a cursor object to execute SQL queries
       cursor = connection.cursor()

       names_and_state_codes = []  # List to store names and state codes

       try:
           # Execute SQL query to select names and state codes from the corper table
           cursor.execute("SELECT name, state_code FROM corpers")

           # Fetch all rows from the result set
           rows = cursor.fetchall()

           # Store the collected names and state codes
           for row in rows:
               names_and_state_codes.append(row)

       except sqlite3.Error as e:
           print("Error reading data from the database:", e)

       finally:
           # Close the cursor and connection
           cursor.close()
           connection.close()

       return names_and_state_codes  # Return the collected names and state codes


   def auto_generate_clearance_file(self):
       # Get current month and year
       # Get today's date
       today = datetime.today()

       # Calculate the first day of the current month
       first_day_of_current_month = today.replace(day=1)

       # Calculate the first day of the previous month
       first_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
       first_day_of_previous_month = first_day_of_previous_month.replace(day=1)

       # Get the name of the previous month and year
       previous_month_name = first_day_of_previous_month.strftime("%B")
       current_year = datetime.now().strftime("_%Y")

       # Construct the filename for the report of the previous month
       report_filename = f"Report_for_{previous_month_name}{current_year}.csv"

       #construct new name for clearance file
       current_month_year = datetime.now().strftime("%B_%Y")
       print(current_month_year)
       # Create a new CSV file for clearance
       clearance_filename = f"{current_month_year}_clearance.csv"

       if os.path.isfile(report_filename) and not os.path.isfile(clearance_filename):
           # Read the report file
           report_records = self.read_csv(report_filename)

           # Get names and state codes from the database
           names_and_state_codes = self.collect_names_and_state_codes()

           # Open the clearance file for writing
           with open(clearance_filename, mode='w', newline='') as csv_file:
               fieldnames = ['Name', 'State Code']
               writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
               writer.writeheader()

               # Iterate through report records
               for record in report_records:
                   name = record['Names']
                   days_absent = int(record['Days Absent'])

                   # Check if the user exists in the database and has less than or equal to three days absent
                   for user in names_and_state_codes:
                       if user[0] == name and days_absent <= self.max_days_absent:
                           writer.writerow({'Name': name, 'State Code': user[1]})
       else:
           print("report file and clearance file exist")

   def replace_placeholders(self, doc, replacements):
       # Loop through each paragraph in the document
       for paragraph in doc.paragraphs:
           # Loop through each placeholder and its replacement
           for placeholder, replacement in replacements.items():
               # Replace the placeholder with the replacement while preserving formatting
               for run in paragraph.runs:
                   if placeholder in run.text:
                       new_text = run.text.replace(placeholder, replacement)
                       run.text = new_text
                       # Preserve formatting
                       new_run = run._element
                       new_run.getparent().replace(new_run, new_run)

       return doc

   def process_word_document(self, doc_file_path):

       # Get current year
       current_year = datetime.now().strftime("%Y")

       # Get current month and year
       current_month_year = datetime.now().strftime("%B_%Y")

       # Construct CSV file name
       csv_file_name = f"{current_month_year}_clearance.csv"

       # Check if CSV file exists
       if os.path.exists(csv_file_name):
           # Read names and state codes from CSV file
           with open(csv_file_name, newline='') as csvfile:
               reader = csv.reader(csvfile)
               next(reader)  # Skip the header row
               for row in reader:
                   if len(row) >= 2:
                       name = row[0].upper()  # Convert name to uppercase
                       state_code = row[1].upper()  # Convert state code to uppercase

                       # Get current month and format it
                       current_month = datetime.now().strftime("%B").upper()  # Convert current month to uppercase

                       # Prepare replacements dictionary
                       replacements = {
                           '{Name}': name,
                           '{State Code}': state_code,
                           '{Date}': datetime.now().strftime("%B %d, %Y"),
                           '{Month}': current_month,
                           '{Year}': current_year
                       }

                       # Open the Word document
                       doc = Document(doc_file_path)

                       # Replace placeholders in the Word document
                       edited_doc = self.replace_placeholders(doc, replacements)

                       # Save the edited document with a new name
                       edited_doc.save(f'{name}_{current_month}_{current_year}.docx')
           current_month_year = datetime.now().strftime("%B_%Y")
           report_filename = f"{current_month_year}_clearance_report.txt"

       else:
           print(f"The CSV file {csv_file_name} does not exist.")

       current_month_year = datetime.now().strftime("%B_%Y")
       report_filename = f"{current_month_year}_clearance_report.txt"
       if not os.path.exists(report_filename):
           sender_email = f'{self.app_email}'
           sender_password = f'{self.app_password}'
           database_path = 'enrollment_data.db'

           # Get recipient data from the database
           data = self.get_recipient_data_from_database(database_path)

           # Send files to recipients
           self.send_files_to_recipients(sender_email, sender_password, data)

           self.generate_clearance_report()

   def send_email_with_attachment(self, sender_email, sender_password, recipient_email, subject, body, attachment_path,
                                  attachment_filename):
       server = None  # Initialize server variable
       try:
           # Set up the MIME
           message = MIMEMultipart()
           message['From'] = sender_email
           message['To'] = recipient_email
           message['Subject'] = subject

           # Attach the body to the email
           message.attach(MIMEText(body, 'plain'))

           # Attach the file using the attachment filename (basename)
           with open(attachment_path, 'rb') as attachment:
               part = MIMEBase('application', 'octet-stream')
               part.set_payload(attachment.read())

           # Encode file in ASCII characters to send by email
           encoders.encode_base64(part)

           # Add header as key/value pair to attachment part
           part.add_header('Content-Disposition',
                           f'attachment; filename= {attachment_filename}')  # Use the attachment filename

           # Add attachment to message and convert message to string
           message.attach(part)
           text = message.as_string()

           # Connect to the SMTP server
           server = smtplib.SMTP('smtp.gmail.com', 587)  # Example: Gmail SMTP server
           server.starttls()

           # Login to the email server
           server.login(sender_email, sender_password)

           # Send the email
           server.sendmail(sender_email, recipient_email, text)

           print("Email sent successfully to:", recipient_email)

       except Exception as e:
           print("An error occurred:", e)

       finally:
           # Quit the server if it was initialized
           if server:
               server.quit()

   def get_recipient_data_from_database(self, database_path):
       try:
           connection = sqlite3.connect(database_path)
           cursor = connection.cursor()

           # Query to select names and email addresses from the 'corpers' table
           cursor.execute("SELECT name, email FROM corpers")
           rows = cursor.fetchall()

           # Extract names and email addresses from the rows
           data = [(row[0], row[1]) for row in rows]

           return data

       except sqlite3.Error as e:
           print("Database error:", e)
           return []

       finally:
           if connection:
               connection.close()

   def send_files_to_recipients(self, sender_email, sender_password, data):
       cwd = os.getcwd()
       for name, email in data:
           # Construct the full filename
           current_month = datetime.now().strftime("%B").upper()  # Get the current month in uppercase
           current_year = datetime.now().strftime("%Y")  # Get the current year
           attachment_filename = f"{name.upper()}_{current_month}_{current_year}.docx"  # Ensure filename is in uppercase

           # Construct the full file path
           attachment_path = os.path.join(cwd, attachment_filename)

           # Check if the attachment file exists
           if os.path.isfile(attachment_path):
               print("Constructed file path:", attachment_path)  # Print the constructed file path
               self.send_email_with_attachment(sender_email, sender_password, email,
                                               f"Monthly Perfomance Clearance for {name}",
                                               f"Dear {name},\nPlease find attached the copy of your monthly perfomance clearance for the month of {current_month} {current_year}.",
                                               attachment_path, attachment_filename)
           else:
               print(f"Attachment file not found for {name}")

   def generate_clearance_report(self):
       current_month_year = datetime.now().strftime("%B_%Y")
       report_filename = f"{current_month_year}_clearance_report.txt"
       with open(report_filename, 'w') as report_file:
           report_file.write(f"{current_month_year} clearance was sent successfully.")

   def send_late_query_emails(self):
       from_date, to_date, previous_month = self.get_previous_month_dates()

       report_filename = f"Report_for_{previous_month}_{datetime.now().year}.csv"

       users_emails = self.get_all_users_emails_from_database('enrollment_data.db')

       if os.path.exists(report_filename):
           with open(report_filename, newline='') as csvfile:
               reader = csv.DictReader(csvfile)

               for row in reader:
                   if 'Names' in row:
                       name = row['Names']
                       days_late = int(
                           row.get('Days Late', 0))  # Use a default value of 0 if 'Days Late' column is missing
                       email = next((item[1] for item in users_emails if item[0] == name), None)

                       if days_late >= self.max_days_late:
                           subject = f"Late Attendance in {previous_month} {datetime.now().year}"
                           body = f"Dear {name},\n\nOur monthly attendance record shows that during {previous_month} {datetime.now().year}, you were late {days_late} times, which is beyond the specified threshold of maximum late attendance per month before query.\n\nPlease provide an explanation for the late attendance.\n\nBest regards,\nAttendance Managament Team"

                           self.send_query_email(self.app_email, self.app_password, email, subject, body)
                   else:
                       print("Warning: 'Names' column not found in the report CSV file. Skipping row.")

       else:
           print(f"Report file {report_filename} does not exist.")

   def send_query_email(self, app_email, app_password, recipient_email, subject, body):
       try:
           # Create the email message
           message = MIMEMultipart()
           message['From'] = app_email
           message['To'] = recipient_email
           message['Subject'] = subject

           # Attach the body of the email
           message.attach(MIMEText(body, 'plain'))

           # Connect to the SMTP server
           server = smtplib.SMTP('smtp.gmail.com', 587)
           server.starttls()

           # Create SSL context and disable certificate verification
           context = ssl.create_default_context()
           context.check_hostname = False
           context.verify_mode = ssl.CERT_NONE

           # Login to the email server
           server.login(app_email, app_password)

           # Send the email
           server.sendmail(app_email, recipient_email, message.as_string())

           print(f"Query email sent successfully to {recipient_email}")

       except Exception as e:
           print(f"Error sending query email to {recipient_email}: {e}")

       finally:
           # Quit the server
           server.quit()

   # Your existing method goes here, without any changes
if __name__ == "__main__":
   root = tk.Tk()
   app = WelcomeFrame(root)
   app.animate_welcome_message()
   app.create_cds_csv()
   root.mainloop()


