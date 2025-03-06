import tkinter as tk
from tkinter import messagebox
import qrcode
import tempfile
import win32api
import pandas as pd
from PIL import ImageTk, Image
from tkinter import Toplevel
from tkinter import ttk
import os
import fnmatch
from datetime import datetime
import sys
import pyautogui
import time


class QRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Label Generator")
        self.root.geometry("1920x1080")
        self.root.resizable(True, True)  # Make the window resizable

        self.button_status1 = False
        self.button_status2 = False
        self.count_penguin = 0
        self.count_picard = 0
        self.num_penguin_box = 12
        self.num_picard_box = 6
        self.num_penguin_pallet = 144
        self.num_picard_pallet = 96
        self.marker_penguin = ''
        self.marker_picard = ''
        self.data_collection = '1'
        self.penguin_data = []
        self.picard_data = []
        self.penguin_file = self.find_latest_file(self.find_files()[0])
        self.picard_file = self.find_latest_file(self.find_files()[1])
        self.penguin_pallet_number = ''
        self.picard_pallet_number = ''

        # Button for setting
        self.setting_button = tk.Button(text="Setting", width=12, font=("Helvetica", 12, "bold"), command=self.setting_window)
        self.setting_button.grid(row=0, column=1, pady=5)

        #************* Penguin Label *************
        # Label for instructions
        self.instructions_label1 = tk.Label(self.root,
                                           text=f"Generate QR Label for Penguin Pallet #{self.penguin_pallet_number}:",
                                           font=("Monaco", 15, "bold"))
        self.instructions_label1.grid(row=0, column=0, sticky="w", padx=10, pady=10)

        # Marker for text area
        self.txt_label1 = tk.Label(self.root, text= self.marker_penguin, font=("Monaco", 15))
        self.txt_label1.grid(row=1, column=0, sticky="w", padx=10, pady=10)

        # Text area for QR data input
        self.text_area1 = tk.Text(self.root, height=self.num_penguin_box + 1, width=110, font=("Monaco", 15))
        self.text_area1.grid(row=1, column=0, padx=10, pady=10)

        # Button for enabling auto clear
        self.toggle_button1 = tk.Button(text="Auto Clear: Off", width=12, font=("Helvetica", 12, "bold"), command=self.toggle_switch1)
        self.toggle_button1.grid(row=2, column=1, pady=5)

        # Button to Generate and Print QR Code
        self.generate_button1 = tk.Button(self.root, text="Print", width=15, height=5, font=("Helvetica", 12, "bold"), command=lambda: self.generate_qr(self.text_area1))
        self.generate_button1.grid(row=1, column=1, pady=5)

        # Label for counting quantity
        self.count1 = tk.Label(self.root, text=f'Quantity:{self.count_penguin}\n File Path: {self.penguin_file}', font=("Helvetica", 18, "bold"))
        self.count1.grid(row=2, column=0, pady=5)

        # Bind KeyRelease and KeyPress events to update line count dynamically
        self.text_area1.bind("<KeyRelease>", lambda event: self.update_line_count(self.text_area1))
        self.text_area1.bind("<KeyPress>", lambda event: self.update_line_count(self.text_area1))

        # Button to Clear Data
        self.clear_button1 = tk.Button(self.root, text="Clear All", font=("Helvetica", 12), command=self.clear_penguin)
        self.clear_button1.grid(row=3, column=1, pady=5)

        #************* Line for separating two text boxes *************
        self.separator1 = ttk.Separator(self.root, orient='horizontal')
        self.separator2 = ttk.Separator(self.root, orient='horizontal')
        self.separator1.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(5, 5))
        self.separator2.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=(5, 5))

        #************* Picard Label *************
        # Label for instructions
        self.instructions_label2 = tk.Label(self.root,
                                           text=f"Generate QR Label for Picard Pallet #{self.picard_pallet_number}:",
                                           font=("Monaco", 15, "bold"))
        self.instructions_label2.grid(row=5, column=0, sticky="w", padx=10, pady=10)

        # Label for text area
        self.txt_label2 = tk.Label(self.root, text=self.marker_picard, font=("Monaco", 15))
        self.txt_label2.grid(row=6, column=0, sticky="w", padx=10, pady=10)

        # Text area for QR data input
        self.text_area2 = tk.Text(self.root, height=self.num_picard_box + 1, width=110, font=("Monaco", 15))
        self.text_area2.grid(row=6, column=0, padx=10, pady=10)

        # Button for enabling auto clear
        self.toggle_button2 = tk.Button(text="Auto Clear: Off", width=12, font=("Helvetica", 12, "bold"), command=self.toggle_switch2)
        self.toggle_button2.grid(row=7, column=1, pady=5)

        # Button to Generate and Print QR Code
        self.generate_button2 = tk.Button(self.root, text="Print", width=15, height=5, font=("Helvetica", 12, "bold"), command=lambda: self.generate_qr(self.text_area2))
        self.generate_button2.grid(row=6, column=1, pady=5)

        # Label for counting quantity
        self.count2 = tk.Label(self.root, text=f'Quantity:{self.count_picard}\n File Path: {self.picard_file}', font=("Helvetica", 18, "bold"))
        self.count2.grid(row=7, column=0, pady=5)

        # Bind KeyRelease and KeyPress events to update line count dynamically
        self.text_area2.bind("<KeyRelease>", lambda event: self.update_line_count(self.text_area2))
        self.text_area2.bind("<KeyPress>", lambda event: self.update_line_count(self.text_area2))

        # Button to Clear Data
        self.clear_button2 = tk.Button(self.root, text="Clear All", font=("Helvetica", 12), command=self.clear_picard)
        self.clear_button2.grid(row=8, column=1, pady=5)

        # Configure grid to expand dynamically when resizing window
        self.root.grid_rowconfigure(1, weight=1)  # Make row 1 resizable
        self.root.grid_columnconfigure(0, weight=1)  # Make column 0 resizable
        self.root.grid_columnconfigure(1, weight=1)  # Make column 1 resizable
        self.resize_textbox()

    def setting_window(self):
        # Get main window's position and size
        main_window_width = self.root.winfo_width()
        main_window_height = self.root.winfo_height()
        main_window_x = self.root.winfo_x()
        main_window_y = self.root.winfo_y()

        # Create the new Toplevel window
        new_window = Toplevel(self.root)
        new_window.title("Setting")
        self.new_window = new_window

        # Get the size of the new window
        window_width = 250  # Width of the new window
        window_height = 350  # Height of the new window

        # Calculate the position to center the new window relative to the main window
        position_top = main_window_y + int(main_window_height / 2 - window_height / 2)
        position_left = main_window_x + int(main_window_width / 2 - window_width / 2)

        # Set the new window position
        new_window.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

        # Setup for Num of Penguin per box
        tk.Label(new_window, text="Number of Penguin per Box:").pack()
        w1 = tk.Scale(new_window, from_=1, to=20, orient=tk.HORIZONTAL)
        w1.set(self.num_penguin_box)
        w1.pack()

        # Setup for Num of Picard per box
        tk.Label(new_window, text="Number of Picard per Box:").pack()
        w2 = tk.Scale(new_window, from_=1, to=20, orient=tk.HORIZONTAL)
        w2.set(self.num_picard_box)
        w2.pack()

        # Setup for Num of Penguin per pallet
        tk.Label(new_window, text="Number of Penguin per Pallet:").pack()
        w3 = tk.Entry(new_window, width=10)
        w3.insert(0, self.num_penguin_pallet)
        w3.pack()

        # Setup for Num of Picard per pallet
        tk.Label(new_window, text="Number of Picard per Pallet:").pack()
        w4 = tk.Entry(new_window, width=10)
        w4.insert(0, self.num_picard_pallet)
        w4.pack()

        checkbox_var = tk.StringVar(value=self.data_collection)
        w5 = tk.Checkbutton(new_window,text="Data Collection", variable=checkbox_var)
        w5.pack(padx=10, pady=10)

        tk.Button(new_window, text="Confirm",command=lambda:self.apply_setting(w1,w2,w3,w4,checkbox_var)).pack(padx=10, pady=10)

    def resize_textbox(self):
        self.marker_picard = ''
        self.marker_penguin = ''
        for i in range(1,self.num_penguin_box + 1):
            self.marker_penguin += f'{i}:\n'

        for i in range(1,self.num_picard_box + 1):
            self.marker_picard += f'{i}:\n'

        self.txt_label1.config(text=self.marker_penguin)
        self.text_area1.config(height=self.num_penguin_box + 1)
        self.txt_label2.config(text=self.marker_picard)
        self.text_area2.config(height=self.num_picard_box + 1)

    def apply_setting(self, w1, w2,w3,w4,checkbox_var):
        self.num_penguin_box = w1.get()
        self.num_picard_box = w2.get()
        self.data_collection = checkbox_var.get()
        self.num_penguin_pallet = int(w3.get())
        self.num_picard_pallet = int(w4.get())
        self.resize_textbox()
        if self.data_collection != 1:
            self.instructions_label1.config(text=f"Generate QR Label for Penguin")
            self.instructions_label2.config(text=f"Generate QR Label for Picard")
        self.new_window.destroy()

    def find_files(self):
        # List to store matching file paths
        penguin_dir = []
        picard_dir = []
        penguin_pattern = "*-*-* Penguin*.xlsx"
        picard_pattern = "*-*-* Picard*.xlsx"
        if getattr(sys, 'frozen', False):
            # If the application is bundled as an executable (using PyInstaller or similar)
            directory = os.path.dirname(sys.executable)
        else:
            # If running as a script
            directory = os.path.dirname(os.path.realpath(__file__))

        # Walk through the directory and its subdirectories
        for root, dirs, files in os.walk(directory):
            for file in files:
                # Check if the file matches the pattern
                if fnmatch.fnmatch(file, penguin_pattern):
                    # Add the full file path to the list
                    penguin_dir.append(os.path.join(root, file))

                if fnmatch.fnmatch(file, picard_pattern):
                    picard_dir.append(os.path.join(root, file))

        return penguin_dir, picard_dir

    def find_latest_file(self, directory):
        date_list = []
        file_list = []  # To store file paths alongside their dates

        # Process files in the local directory
        for file in directory:
            basename = os.path.basename(file)
            date_str = basename.split()[0]  # This assumes the date is the first part of the filename

            try:
                # Try to parse the date string to a datetime object
                date_info = datetime.strptime(date_str, "%m-%d-%y")
                date_list.append(date_info)
                file_list.append(file)  # Store corresponding file path

            except ValueError:
                continue

        if date_list:
            # Sort the date_list in descending order (latest date first)
            sorted_dates = sorted(zip(date_list, file_list), reverse=True)

            # Extract the file corresponding to the latest date
            latest_date, latest_file = sorted_dates[0]

            return latest_file
        else:
            return None

    def handle_qr_data(self, area):
        text_area = area
        qr_data = text_area.get("1.0", "end-1c").strip()  # Get all text input
        if not qr_data:
            messagebox.showwarning("No Data", "Please enter QR data.")
            return

        formatted_data = qr_data.replace('}{', '}\n{')
        data_list = formatted_data.split('\n')

        for item in data_list:
            if len(item) != len(data_list[0]):
                messagebox.showwarning("Invalid Format", "Inconsistent length.")
                return

        if len(data_list) != len(set(data_list)):
            messagebox.showwarning("Invalid Format", "Duplicate characters.")
            return

        if text_area == self.text_area1 and len(data_list) != self.num_penguin_box:
            messagebox.showwarning("Invalid Format", "Incorrect number of Penguins.")
            return

        if text_area == self.text_area2 and len(data_list) != self.num_picard_box:
            messagebox.showwarning("Invalid Format", "Incorrect number of Picards.")
            return

        return formatted_data

    def save_to_xlsx(self, area):
        text_area = area
        self.penguin_data = []
        self.picard_data = []
        penguin_file = self.penguin_file
        picard_file = self.picard_file

        # Ensure the file exists before reading
        try:
            # Read existing data
            if text_area == self.text_area1:
                penguin_df_existing = pd.read_excel(penguin_file)
                self.penguin_pallet_number = int(len(penguin_df_existing.index) / self.num_penguin_pallet) + 1
                print(f"Successfully loaded penguin data from: {penguin_file}")


            if text_area == self.text_area2:
                picard_df_existing = pd.read_excel(picard_file)
                self.picard_pallet_number = int(len(picard_df_existing.index) / self.num_picard_pallet) + 1
                print(f"Successfully loaded picard data from: {picard_file}")

        except Exception as e:
            messagebox.showwarning("Warning", f"Error reading excel file.{e}")
            return

        # Check the area and handle appropriately
        if text_area == self.text_area1:
            formatted_data_penguin = self.handle_qr_data(area)
            penguin_list = formatted_data_penguin.split('\n')

            # Ensure that formatted_data_penguin is not None before appending
            if formatted_data_penguin is not None:
                for i in range(len(penguin_list)):
                 self.penguin_data.append(penguin_list[i])

                try:
                    # Convert the new data to a DataFrame
                    new_penguin_df = pd.DataFrame(self.penguin_data, columns=["QR Data"])

                    # Concatenate the new data with the existing data
                    penguin_df_combined = pd.concat([penguin_df_existing, new_penguin_df], ignore_index=True)

                    # Find duplicated value
                    duplicates = penguin_df_combined[penguin_df_combined.duplicated()]
                    if not duplicates.empty:
                        messagebox.showwarning("Warning", "Duplicates found when trying to save data.")
                        return
                    else:
                        print("No duplicate entries found.")

                    # Save the combined data to Excel
                    penguin_df_combined.to_excel(penguin_file, index=False)
                    print(f"Data successfully saved to: {penguin_file}")
                except Exception as e:
                    print(f"Error appending data to Excel: {e}")

        if text_area == self.text_area2:
            formatted_data_picard = self.handle_qr_data(area)
            picard_list = formatted_data_picard.split('\n')

            # Ensure that formatted_data_picard is not None before appending
            if formatted_data_picard is not None:
                for i in range(len(picard_list)):
                 self.picard_data.append(picard_list[i])

                try:
                    # Convert the new data to a DataFrame
                    new_picard_df = pd.DataFrame(self.picard_data, columns=["QR Data"])

                    # Concatenate the new data with the existing data
                    picard_df_combined = pd.concat([picard_df_existing, new_picard_df], ignore_index=True)

                    # Find duplicated value
                    duplicates = picard_df_combined[picard_df_combined.duplicated()]

                    if not duplicates.empty:
                        messagebox.showwarning("Warning", "Duplicates found when trying to save data.")
                        return
                    else:
                        print("No duplicate entries found.")

                    # Save the combined data to Excel
                    picard_df_combined.to_excel(picard_file, index=False)
                    print(f"Data successfully saved to: {picard_file}")
                except Exception as e:
                    print(f"Error appending data to Excel: {e}")

        self.instructions_label1.config(text=f"Generate QR Label for Penguin Pallet #{self.penguin_pallet_number}:")
        self.instructions_label2.config(text=f"Generate QR Label for Picard Pallet #{self.picard_pallet_number}:")

    def generate_qr(self, area):
        text_area = area
        formatted_data = self.handle_qr_data(area)

        if not formatted_data:
            return

        # Generate QR Code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )

        qr.add_data(formatted_data)
        qr.make(fit=True)

        img = qr.make_image(fill='black', back_color='white')

        # Resize the image to a fixed size (e.g., 500x500)
        img = img.resize((500, 500))  # Change the size as needed

        # Convert to ImageTk format
        self.qr_image = ImageTk.PhotoImage(img)

        # Save the image for printing purposes
        self.img_for_penguin = img  # Save the original PIL image for printing

        # Automatically print the QR code after generation
        self.print_qr(self.img_for_penguin)

        if self.data_collection == '1' and text_area == self.text_area1:
            self.save_to_xlsx(area)

        if self.data_collection == '1' and text_area == self.text_area2:
            self.save_to_xlsx(area)

        if self.button_status1 and text_area == self.text_area1:
            self.clear_penguin()

        if self.button_status2 and text_area == self.text_area2:
            self.clear_picard()

    def count_num(self, text):
        # Count the element of items in the Text widget
        qr_data = text.get("1.0", "end-1c").strip()  # Get all text input
        if not qr_data:
            return  0
        formatted_data = qr_data.replace('}{', '}\n{')
        data_list = formatted_data.split('\n')
        num = len(data_list)
        self.count_penguin = num
        return self.count_penguin

    def update_line_count(self, text):
        # Update the label with live line count when user types or deletes
        new_count = self.count_num(text)
        if text == self.text_area1:
            self.count1.config(text=f'Quantity: {new_count}')
        if text == self.text_area2:
            self.count2.config(text=f'Quantity: {new_count}')

    def toggle_switch1(self):
        if self.toggle_button1.config('relief')[-1] == 'sunken':
            self.toggle_button1.config(relief="raised", bg="SystemButtonFace", text="Auto Clear: Off")
            self.button_status1 = False
        else:
            self.toggle_button1.config(relief="sunken", bg="green", text="Auto Clear: On")
            self.button_status1 = True

    def toggle_switch2(self):
        if self.toggle_button2.config('relief')[-1] == 'sunken':
            self.toggle_button2.config(relief="raised", bg="SystemButtonFace", text="Auto Clear: Off")
            self.button_status2 = False
        else:
            self.toggle_button2.config(relief="sunken", bg="green", text="Auto Clear: On")
            self.button_status2 = True

    def clear_penguin(self):
        self.text_area1.delete("1.0", "end")
        self.img_for_penguin = None
        self.count1.config(text=f'Quantity: 0')

    def clear_picard(self):
        self.text_area2.delete("1.0", "end")
        self.img_for_picard = None
        self.count2.config(text=f'Quantity: 0')

    def print_qr(self, image):
        file_to_print = image

        if file_to_print:
            # Save the image to a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_file:
                temp_file_path = temp_file.name
                file_to_print.save(temp_file_path)

            # Use the default Windows "print" command to send the image to the printer
            win32api.ShellExecute(0, "print", temp_file_path, None, ".", 0)
            time.sleep(0.5)  # Delay for 0.5 seconds
            pyautogui.hotkey('alt', 'f')


# Create the main window
root = tk.Tk()
app = QRApp(root)

# Run the application
root.mainloop()
