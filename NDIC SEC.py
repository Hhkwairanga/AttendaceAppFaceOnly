import os
import shutil
import tkinter as tk
from tkinter import PhotoImage, messagebox
from tkinter import ttk
from tkcalendar import DateEntry
import cv2
import threading
from PIL import Image, ImageTk
#from pyfingerprint.pyfingerprint import PyFingerprint
import sqlite3
import os
import sys


# Determine the absolute path to the directory containing the script
script_directory = os.path.dirname(os.path.abspath(__file__))

# Function to request camera permission using AppleScript
def request_camera_permission():
   # AppleScript command to request camera permission
   applescript_command = '''
   tell application "System Events"
       activate
       display dialog "This app requires access to your camera to capture images for enrollment." buttons {"OK"} default button "OK" with icon caution
   end tell
   '''

   # Execute the AppleScript command
   os.system(f"osascript -e '{applescript_command}'")

# Check if running on macOS and request camera permission
if sys.platform == "darwin":
   request_camera_permission()

# Determine the absolute path to the directory containing the script
script_directory = os.path.dirname(os.path.abspath(__file__))

# Specify the absolute path to the SQLite database file
db_file_path = os.path.join(script_directory, 'enrollment_data.db')

try:
   # Connect to the SQLite database (or create it if it doesn't exist)
   conn = sqlite3.connect(db_file_path)
   cursor = conn.cursor()

   # Create a table for storing Corper data
   cursor.execute('''
       CREATE TABLE IF NOT EXISTS corpers (
           id INTEGER PRIMARY KEY AUTOINCREMENT,
           name TEXT NOT NULL,
           state_code TEXT NOT NULL,
           cds_day TEXT NOT NULL,
           passing_out_date DATE NOT NULL,
           email TEXT NULL,  -- Add email column
           face_image_path TEXT NULL,
           fingerprint_image_path TEXT NULL
       )
   ''')

   # Create a table for storing Intern data
   cursor.execute('''
       CREATE TABLE IF NOT EXISTS interns (
           id INTEGER PRIMARY KEY AUTOINCREMENT,
           name TEXT NOT NULL,
           matric_number TEXT NOT NULL,
           it_duration INTEGER NOT NULL,
           school_resumption_date DATE NOT NULL,
           email TEXT NULL,  -- Add email column
           face_image_path TEXT NULL,
           fingerprint_image_path TEXT NULL
       )
   ''')

   # Commit the changes
   conn.commit()

except sqlite3.Error as e:
   print("Error:", e)

finally:
   # Close the connection
   if conn:
       conn.close()

'''class WelcomeFrame(tk.Frame):
   def __init__(self, master=None):
       super().__init__(master)
       self.master = master
       self.master.geometry("800x600")
       root.title("NDIC SEC")  # Set the title for the window
       #root.iconbitmap("icon.ico")  # Replace with the path to your icon file
       self.create_widgets()


   def create_widgets(self):
       self.welcome_label = tk.Label(self, text="Welcome to NDIC corpers/interns Enrollment System", font=("Helvetica", 20))
       self.welcome_label.pack(pady=20)'''

class WelcomeFrame(tk.Frame):
   def __init__(self, master=None):
       super().__init__(master)
       self.master = master
       self.master.geometry("800x600")  # Set fixed resolution
       self.pack()
       self.create_widgets()
       root.title("NDIC SEC")  # Set the title for the window

       # Store the original size of the background logo
       self.background_logo_original = self.background_logo

   def create_widgets(self):
       # Load background logo with expansion
       self.background_logo = PhotoImage(file=os.path.join(script_directory, 'bg_logo.png'))
       self.background_logo = self.background_logo.zoom(3)  # Optional: Adjust zoom factor
       self.logo_label = tk.Label(self.master, image=self.background_logo)
       self.logo_label.place(x=(800 - self.background_logo.width()) / 2,
                             y=(600 - self.background_logo.height()) / 2)  # Center the logo

       # Create a label for the welcome message with a larger font size and starting on a new line
       self.welcome_label = tk.Label(self, text="", font=("Helvetica", 36, "bold"), fg="dark blue")
       self.welcome_label.pack(pady=20)

       # Set a flag to track whether the welcome animation is complete
       self.animation_complete = False

       # Create "Start Enrollment" button after the animation is complete
       self.start_enrollment_button = ttk.Button(self, text="Start Enrollment", command=self.start_enrollment, style='TButton')
       # Do not pack the button initially
       # self.start_enrollment_button.pack(pady=10)

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
       welcome_text = "Welcome to\n NDIC corpers/interns Enrollment System"  # Use \n for a new line
       for i in range(1, len(welcome_text) + 1):
           self.welcome_label.config(text=welcome_text[:i])
           self.master.update_idletasks()  # Update the display
           self.master.after(100)  # Delay between characters (in milliseconds)

       # Set the animation complete flag to True after the animation is done
       self.animation_complete = True

       # Pack the "Start Enrollment" button after the animation is complete
       self.start_enrollment_button.pack(pady=10)

   def handle_resize(self, event):
       # Adjust the logo size proportionally to the window size
       new_width = event.width
       new_height = int(self.background_logo_original.height() * (new_width / self.background_logo_original.width()))
       self.background_logo = self.background_logo_original.subsample(
           round(self.background_logo_original.width() / new_width))
       self.background_logo = self.background_logo.zoom(round(self.background_logo_original.width() / new_width))
       self.logo_label.config(image=self.background_logo)
       self.logo_label.place(x=(new_width - self.background_logo.width()) / 2,
                             y=(new_height - self.background_logo.height()) / 2)

       # Adjust the welcome label font size
       font_size = max(14, round(new_width / 30))
       self.welcome_label.config(font=("Helvetica", font_size, "bold"))

       # Adjust the button padding
       button_padding = max(5, round(new_width / 80))
       self.master.style.configure('TButton', padding=button_padding)

   def start_enrollment(self):
       if self.animation_complete:
           self.destroy()  # Destroy the current frame
           enrollment_frame = EnrollmentSelectionFrame(self.master, previous_frame=self)
           enrollment_frame.pack()

class EnrollmentSelectionFrame(tk.Frame):
   def __init__(self, master=None, previous_frame=None):
       super().__init__(master)
       self.master = master
       self.master.geometry("800x600")  # Set fixed resolution
       self.master.resizable(True, True)  # Allow both width and height resizing
       self.previous_frame = previous_frame  # Store the reference to the previous frame
       self.create_widgets()

   def create_widgets(self):
       # Add a top-left logo and title
       self.top_left_logo = PhotoImage(file=os.path.join(script_directory, 'top_left_logo.png'))
       self.top_left_logo_label = tk.Label(self, image=self.top_left_logo)
       self.top_left_logo_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)

       self.title_label = tk.Label(self, text="Select Category: Corper or Intern", font=("Helvetica", 18, "bold"))
       self.title_label.grid(row=1, column=0, columnspan=2, pady=10)

       # Radio type selection
       self.category_var = tk.StringVar(value="corper")

       self.corper_radio = tk.Radiobutton(self, text="Corper", variable=self.category_var, value="corper",
                                          font=("Helvetica", 16))
       self.corper_radio.grid(row=2, column=0, pady=10, padx=10, sticky="w")

       self.intern_radio = tk.Radiobutton(self, text="Intern", variable=self.category_var, value="intern",
                                          font=("Helvetica", 16))
       self.intern_radio.grid(row=3, column=0, pady=10, padx=10, sticky="w")

       # Style the radio buttons
       self.master.style.configure('TRadiobutton',
                                   font=('Helvetica', 16),
                                   foreground='black',
                                   background='lightgray')

       # Proceed button
       self.proceed_button = ttk.Button(self, text="Proceed", command=self.proceed, style='TButton')
       self.proceed_button.grid(row=4, column=0, pady=20)

       # Configure button style with hovering effects
       self.master.style.configure('TButton',
                                   foreground='black',
                                   background='lightgray',
                                   font=('Helvetica', 18, 'bold'),
                                   padding=15)
       self.master.style.map('TButton',
                             foreground=[('active', 'green')],
                             background=[('active', 'lightgreen')])

   def proceed(self):
       selected_category = self.category_var.get()

       # Add logic to determine which form to show based on the selected category
       if selected_category == "corper":
           self.destroy()  # Destroy the current frame
           corper_form_frame = CorperFormFrame(self.master, previous_frame=self)
           corper_form_frame.pack()
       elif selected_category == "intern":
           self.destroy()  # Destroy the current frame
           intern_form_frame = InternFormFrame(self.master, previous_frame=self)
           intern_form_frame.pack()

   def go_back(self):
       self.destroy()  # Destroy the current frame
       previous_frame = WelcomeFrame(self.master)  # Create a new instance of the previous frame
       previous_frame.pack()  # Pack the new instance


class CorperFormFrame(tk.Frame):
   def __init__(self, master=None, previous_frame=None):
       super().__init__(master)
       self.master = master
       self.previous_frame = previous_frame
       self.master.geometry("800x600")  # Set fixed resolution
       self.create_widgets()
       self.name = ""

   def create_widgets(self):
       # Add a top-left logo and title
       self.top_left_logo = PhotoImage(file=os.path.join(script_directory, 'top_left_logo.png'))
       self.top_left_logo_label = tk.Label(self, image=self.top_left_logo)
       self.top_left_logo_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)

       # Corper form elements
       tk.Label(self, text="Name:").grid(row=1, column=0, pady=10)
       self.name_entry = tk.Entry(self)
       self.name_entry.grid(row=1, column=1, pady=10, sticky="ew")

       tk.Label(self, text="State Code:").grid(row=2, column=0, pady=10)
       self.state_code_entry = tk.Entry(self)
       self.state_code_entry.grid(row=2, column=1, pady=10, sticky="ew")

       tk.Label(self, text="CDs Day:").grid(row=3, column=0, pady=10)
       self.cds_day_var = tk.StringVar(value="Monday")  # Default value
       cds_day_options = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
       self.cds_day_combobox = ttk.Combobox(self, textvariable=self.cds_day_var, values=cds_day_options,
                                            state="readonly")
       self.cds_day_combobox.grid(row=3, column=1, pady=10, sticky="ew")

       tk.Label(self, text="Passing out Date:").grid(row=4, column=0, pady=10)
       self.passing_out_date_entry = DateEntry(
           self, date_pattern="yyyy-mm-dd", background='darkgray', foreground='black', selectbackground='darkred'
       )
       self.passing_out_date_entry.grid(row=4, column=1, pady=10, padx=10, sticky="ew")

       tk.Label(self, text="Email:").grid(row=5, column=0, pady=10)
       self.email_entry = tk.Entry(self)  # Add email entry
       self.email_entry.grid(row=5, column=1, pady=10, sticky="ew")

       ttk.Button(self, text="Proceed", command=self.submit, style='TButton').grid(row=6, column=0, columnspan=2,
                                                                                  pady=20)
       ttk.Button(self, text="Back", command=self.go_back, style='TButton').grid(row=7, column=0, columnspan=2,
                                                                                 pady=20)


   def go_back(self):
       self.destroy()  # Destroy the current frame
       previous_frame = EnrollmentSelectionFrame(self.master)  # Create a new instance of the previous frame
       previous_frame.pack()  # Pack the new instance

   def submit(self):
       # Retrieve values from the entries and combobox
       name = self.name_entry.get()
       state_code = self.state_code_entry.get()
       cds_day = self.cds_day_var.get()
       passing_out_date = self.passing_out_date_entry.get()
       email = self.email_entry.get()

       # Check if all fields are filled
       if not name or not state_code or not cds_day or not passing_out_date:
           messagebox.showerror("Error", "Please fill in all fields.")
           return
       # Store the entered name in the instance variable
       self.name = name

       # Add logic to store Corper data in the database
       conn = sqlite3.connect(db_file_path)
       cursor = conn.cursor()
       cursor.execute('''
           INSERT INTO corpers (name, state_code, cds_day, passing_out_date, email, face_image_path, fingerprint_image_path)
           VALUES (?, ?, ?, ?, ?, ?, ?)
       ''', (name, state_code, cds_day, passing_out_date, email, 'None', 'None'))  # Use None for optional (NULL) values
       conn.commit()
       conn.close()

       # Optionally, you can add logic to transition to the next frame or perform other actions
       # Proceed to the CaptureFrame
       self.destroy()
       capture_frame = CaptureFrame(self.master, previous_frame=self, name=self.name)
       capture_frame.pack()


class InternFormFrame(tk.Frame):
   def __init__(self, master=None, previous_frame=None):
       super().__init__(master)
       self.master = master
       self.previous_frame = previous_frame
       self.master.geometry("800x600")  # Set fixed resolution
       self.create_widgets()
       self.name = ""

   def create_widgets(self):
       # Add a top-left logo and title
       self.top_left_logo = PhotoImage(file=os.path.join(script_directory, 'top_left_logo.png'))
       self.top_left_logo_label = tk.Label(self, image=self.top_left_logo)
       self.top_left_logo_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)

       # Intern form elements
       tk.Label(self, text="Name:").grid(row=1, column=0, pady=10)
       self.name_entry = tk.Entry(self)
       self.name_entry.grid(row=1, column=1, pady=10)

       tk.Label(self, text="Matric Number:").grid(row=2, column=0, pady=10)
       self.matric_number_entry = tk.Entry(self)
       self.matric_number_entry.grid(row=2, column=1, pady=10)

       tk.Label(self, text="IT Duration (months)").grid(row=3, column=0, pady=10)

       it_duration_options = list(range(1, 13))

       self.it_duration_var = tk.StringVar()
       self.it_duration_combobox = ttk.Combobox(self, textvariable=self.it_duration_var, values=it_duration_options, width=15)
       self.it_duration_combobox.grid(row=3, column=1, pady=10)

       tk.Label(self, text="School Resumption Date:").grid(row=4, column=0, pady=10)
       self.school_resumption_date_cal = DateEntry(self, date_pattern="yyyy-mm-dd", background='darkgray', foreground='black', selectbackground='darkred')
       self.school_resumption_date_cal.grid(row=4, column=1, pady=10)

       tk.Label(self, text="Email:").grid(row=5, column=0, pady=10)
       self.email_entry = tk.Entry(self)  # Add email entry
       self.email_entry.grid(row=5, column=1, pady=10, sticky="ew")

       ttk.Button(self, text="Proceed", command=self.submit, style='TButton').grid(row=6, column=0, columnspan=2, pady=20)
       ttk.Button(self, text="Back", command=self.go_back, style='TButton').grid(row=7, column=0, columnspan=2, pady=20)

   def go_back(self):
       self.destroy()  # Destroy the current frame
       previous_frame = EnrollmentSelectionFrame(self.master)  # Create a new instance of the previous frame
       previous_frame.pack()  # Pack the new instance

   def submit(self):
       # Retrieve values from the entries and widgets
       name = self.name_entry.get()
       matric_number = self.matric_number_entry.get()
       it_duration_str = self.it_duration_var.get()
       email = self.email_entry.get()
       # Check if IT duration is selected
       if not it_duration_str:
           messagebox.showerror("Error", "Please select IT duration.")
           return

       # Convert IT duration to int
       try:
           it_duration = int(it_duration_str)
       except ValueError:
           messagebox.showerror("Error", "Invalid value for IT duration.")
           return

       school_resumption_date = self.school_resumption_date_cal.get_date()

       # Check if all fields are filled
       if not name or not matric_number or not school_resumption_date:
               messagebox.showerror("Error", "Please fill in all fields.")
               return

       # Add logic to store Intern data in the database
       conn = sqlite3.connect(db_file_path)
       cursor = conn.cursor()
       cursor.execute('''
                   INSERT INTO interns (name, matric_number, it_duration, school_resumption_date, email, face_image_path, fingerprint_image_path)
                   VALUES (?, ?, ?, ?, ?, ?, ?)
               ''', (name, matric_number, it_duration, school_resumption_date, email, None, None))
       conn.commit()
       conn.close()

       # Store the entered name in the instance variable
       self.name = name

       # Optionally, you can add logic to transition to the next frame or perform other actions
       # Proceed to the CaptureFrame
       self.destroy()
       capture_frame = CaptureFrame(self.master, previous_frame=self, name=self.name)
       capture_frame.pack()


class CaptureFrame(tk.Frame):
   def __init__(self, master=None, previous_frame=None, name=""):
       super().__init__(master)
       self.master = master
       self.master.geometry("800x600")  # Set fixed resolution
       self.master.resizable(True, True)  # Allow both width and height resizing
       self.previous_frame = previous_frame  # Store the reference to the previous frame
       self.name = name  # Store the name received as an argument
       self.create_widgets()

   def create_widgets(self):
       # Add a top-left logo and title
       self.top_left_logo = PhotoImage(file=os.path.join(script_directory, 'top_left_logo.png'))
       self.top_left_logo_label = tk.Label(self, image=self.top_left_logo)
       self.top_left_logo_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)

       # Create "Face Capture" button
       self.face_capture_button = ttk.Button(self, text="Face Capture", command=self.face_capture, style='TButton')
       self.face_capture_button.grid(row=1, column=0, pady=10, padx=10)

       # Create "Biometric Capture" button
       self.biometric_capture_button = ttk.Button(self, text="Biometric Capture", style='TButton')
       self.biometric_capture_button.grid(row=1, column=1, pady=20, padx=10)


       # Configure button style with hovering effects
       self.master.style.configure('TButton',
                                   foreground='black',
                                   background='lightgray',
                                   font=('Helvetica', 18, 'bold'),
                                   padding=15)
       self.master.style.map('TButton',
                             foreground=[('active', 'green')],
                             background=[('active', 'lightgreen')])

       # Create "Submit" button
       self.submit_button = ttk.Button(self, text="Submit", command=self.submit, style='TButton')
       self.submit_button.grid(row=2, column=0, columnspan=2, pady=20)
       self.submit_button.grid_remove()  # Initially hidden

   def show_submit_button(self):
       self.submit_button.grid()  # Make the button visible

   def start_face_capture(self):
       face_capture_thread = threading.Thread(target=self.face_capture)
       face_capture_thread.start()

   def face_capture(self):
       # Determine the directory containing the script
       script_directory = os.path.dirname(os.path.abspath(__file__))

       # Specify the path to the cascade file in the local directory
       cascade_file_path = os.path.join(script_directory, 'haarcascade_frontalface_default.xml')

       # Load the cascade classifier
       face_cascade = cv2.CascadeClassifier(cascade_file_path)

       # Initialize video capture
       video_capture = cv2.VideoCapture(0)

       count = 0  # Count to capture 3 pictures
       while count < 3:
           # Capture frame-by-frame
           ret, frame = video_capture.read()

           # Check if frame retrieval was successful
           if not ret:
               print("Error: Failed to retrieve frame from webcam.")
               break

           # Convert the frame to grayscale for face detection
           gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

           # Detect faces in the frame
           faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))

           # Draw rectangles around the faces
           for (x, y, w, h) in faces:
               cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)

           # Display the frame with face detection
           cv2.imshow('Face Capture', frame)

           # If faces are detected, capture the frame
           if len(faces) > 0:
               count += 1
               cv2.imwrite(f'capture_{count}.png', frame)
               messagebox.showinfo("Capture", f"Captured {count}/3")

           # Break the loop if 'q' key is pressed
           if cv2.waitKey(1) & 0xFF == ord('q'):
               break

       # Release the webcam and close the OpenCV window
       video_capture.release()
       cv2.destroyAllWindows()

       # Show a message when capturing is complete
       messagebox.showinfo("Capture Complete", "Face capture is complete. You can close the window.")

       self.show_submit_button()
       # ...
   def biometric_capture(self):
       # PyFingerprint code for fingerprint capture
       try:
           f = PyFingerprint('/dev/ttyUSB0', 57600, 0xFFFFFFFF, 0x00000000)

           if not f.verifyPassword():
               raise ValueError('The given fingerprint sensor password is wrong!')

       except Exception as e:
           messagebox.showerror("Fingerprint Sensor Error", f"Error: {e}")
           return

       count = 0
       while count < 2:  # Capture two fingerprints
           try:
               print('Waiting for finger...')
               while not f.readImage():
                   pass

               f.convertImage(0x01)

               # Save the fingerprint image
               image_path = f"capture_fingerprint_{count + 1}.bmp"
               f.saveImage(image_path)

               messagebox.showinfo("Capture", f"Fingerprint {count + 1}/2 captured successfully.")
               count += 1

           except Exception as e:
               messagebox.showerror("Fingerprint Capture Error", f"Error: {e}")
               break
       self.show_submit_button() # Show the submit button after capture completion

   def reset_enrollment(self):
       # Reset the enrollment process and navigate back to EnrollmentSelectionFrame
       self.destroy()
       enrollment_frame = EnrollmentSelectionFrame(self.master)
       enrollment_frame.pack()

   def submit(self):
       # Create a folder with the user's name
       folder_path = os.path.join(os.getcwd(), self.name)
       os.makedirs(folder_path, exist_ok=True)

       # Move facial images to the folder
       self.move_images("capture_", ".png", folder_path)

       # Check if any biometric images were captured
       biometric_images_captured = any(
           file.startswith("capture_fingerprint_") and file.endswith(".bmp") for file in os.listdir())

       # Move biometric images to the folder only if they exist
       if biometric_images_captured:
           self.move_images("capture_fingerprint_", ".bmp", folder_path)

       # Additional logic for handling the submission
       # Update the paths in the SQLite database
       conn = sqlite3.connect(db_file_path)
       cursor = conn.cursor()

       # Update face_image_path based on user's name
       cursor.execute('''
              UPDATE corpers
              SET face_image_path = ?
              WHERE name = ?
          ''', (folder_path, self.name))

       cursor.execute('''
              UPDATE interns
              SET face_image_path = ?
              WHERE name = ?
          ''', (folder_path, self.name))

       # Update fingerprint_image_path only if biometric images are captured
       if biometric_images_captured:
           cursor.execute('''
                  UPDATE corpers
                  SET fingerprint_image_path = ?
                  WHERE name = ?
              ''', (folder_path, self.name))

           cursor.execute('''
                  UPDATE interns
                  SET fingerprint_image_path = ?
                  WHERE name = ?
              ''', (folder_path, self.name))

       conn.commit()
       conn.close()

       # Change the text of the submit button
       self.submit_button.config(text="Enroll New User", command=self.reset_enrollment)

   def move_images(self, prefix, extension, destination_folder):
       # Get a list of image files with the specified prefix and extension
       image_files = [file for file in os.listdir() if file.startswith(prefix) and file.endswith(extension)]

       # Move each image file to the destination folder
       for image_file in image_files:
           source_path = os.path.join(os.getcwd(), image_file)
           destination_path = os.path.join(destination_folder, image_file)
           shutil.move(source_path, destination_path)


       # Optionally, you can add logic to handle cases when no images are found

       # Optionally, you can add logic to handle the submission

       # Show a message when the move operation is complete
       messagebox.showinfo(f"Successfull", f"Enrollment Complete for {self.name}.")

if __name__ == "__main__":
   root = tk.Tk()
   app = WelcomeFrame(root)
   root.title("NDIC SEC")  # Set your desired title
   img = tk.PhotoImage(file=os.path.join(script_directory, "log-in.png"))
   root.tk.call('wm', 'iconphoto', root._w, img)
   app.animate_welcome_message()
   root.mainloop()