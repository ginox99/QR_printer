import tkinter as tk
from tkinter import messagebox
import qrcode
import tempfile
import win32api
from PIL import ImageTk, Image
from tkinter import Toplevel

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
        self.num_penguin = 12
        self.num_picard = 6
        self.marker_penguin = ''
        self.marker_picard = ''

        # Button for setting
        self.setting_button = tk.Button(text="Setting", width=12, font=("Helvetica", 12, "bold"), command=self.setting_window)
        self.setting_button.grid(row=0, column=1, pady=5)

        ########## Penguin Label ##############
        # Label for instructions
        self.instructions_label1 = tk.Label(self.root,
                                           text="Generate QR Label for Penguin:",
                                           font=("Helvetica", 15, "bold"))
        self.instructions_label1.grid(row=0, column=0, sticky="w", padx=10, pady=10)

        # Marker for text area
        self.txt_label1 = tk.Label(self.root, text= self.marker_penguin, font=("Helvetica", 15))
        self.txt_label1.grid(row=1, column=0, sticky="w", padx=10, pady=10)

        # Text area for QR data input
        self.text_area1 = tk.Text(self.root, height=self.num_penguin + 1, width=110, font=("Monaco", 15))
        self.text_area1.grid(row=1, column=0, padx=10, pady=10)

        # Button for enabling auto clear
        self.toggle_button1 = tk.Button(text="Auto Clear: Off", width=12, font=("Helvetica", 12, "bold"), command=self.toggle_switch1)
        self.toggle_button1.grid(row=2, column=1, pady=5)

        # Button to Generate and Print QR Code
        self.generate_button1 = tk.Button(self.root, text="Print", width=15, height=5, font=("Helvetica", 12, "bold"), command=lambda: self.generate_qr(self.text_area1))
        self.generate_button1.grid(row=1, column=1, pady=5)

        # Label for counting quantity
        self.count1 = tk.Label(self.root, text=f'Quantity:{self.count_penguin}', font=("Helvetica", 18, "bold"))
        self.count1.grid(row=2, column=0, pady=5)

        # Bind KeyRelease and KeyPress events to update line count dynamically
        self.text_area1.bind("<KeyRelease>", lambda event: self.update_line_count(self.text_area1))
        self.text_area1.bind("<KeyPress>", lambda event: self.update_line_count(self.text_area1))

        # Button to Clear Data
        self.clear_button1 = tk.Button(self.root, text="Clear All", font=("Helvetica", 12), command=self.clear_penguin)
        self.clear_button1.grid(row=3, column=1, pady=5)

        ########## Picard Label ##############
        # Label for instructions
        self.instructions_label2 = tk.Label(self.root,
                                           text="Generate QR Label for Picard:",
                                           font=("Helvetica", 15, "bold"))
        self.instructions_label2.grid(row=4, column=0, sticky="w", padx=10, pady=10)

        # Label for text area
        self.txt_label2 = tk.Label(self.root, text=self.marker_picard, font=("Helvetica", 15))
        self.txt_label2.grid(row=5, column=0, sticky="w", padx=10, pady=10)

        # Text area for QR data input
        self.text_area2 = tk.Text(self.root, height=self.num_picard + 1, width=110, font=("Monaco", 15))
        self.text_area2.grid(row=5, column=0, padx=10, pady=10)

        # Button for enabling auto clear
        self.toggle_button2 = tk.Button(text="Auto Clear: Off", width=12, font=("Helvetica", 12, "bold"), command=self.toggle_switch2)
        self.toggle_button2.grid(row=6, column=1, pady=5)

        # Button to Generate and Print QR Code
        self.generate_button2 = tk.Button(self.root, text="Print", width=15, height=5, font=("Helvetica", 12, "bold"), command=lambda: self.generate_qr(self.text_area2))
        self.generate_button2.grid(row=5, column=1, pady=5)

        # Label for counting quantity
        self.count2 = tk.Label(self.root, text=f'Quantity:{self.count_picard}', font=("Helvetica", 18, "bold"))
        self.count2.grid(row=6, column=0, pady=5)

        # Bind KeyRelease and KeyPress events to update line count dynamically
        self.text_area2.bind("<KeyRelease>", lambda event: self.update_line_count(self.text_area2))
        self.text_area2.bind("<KeyPress>", lambda event: self.update_line_count(self.text_area2))

        # Button to Clear Data
        self.clear_button2 = tk.Button(self.root, text="Clear All", font=("Helvetica", 12), command=self.clear_picard)
        self.clear_button2.grid(row=7, column=1, pady=5)

        # Configure grid to expand dynamically when resizing window
        self.root.grid_rowconfigure(1, weight=1)  # Make row 1 resizable
        self.root.grid_columnconfigure(0, weight=1)  # Make column 0 resizable
        self.root.grid_columnconfigure(1, weight=1)  # Make column 1 resizable
        self.apply_setting()

    def setting_window(self):
        # Get main window's position and size
        main_window_width = self.root.winfo_width()
        main_window_height = self.root.winfo_height()
        main_window_x = self.root.winfo_x()
        main_window_y = self.root.winfo_y()

        # Create the new Toplevel window
        new_window = Toplevel(self.root)
        new_window.title("Setting")
        new_window.geometry("250x250")
        self.new_window = new_window

        # Get the size of the new window
        window_width = 250  # Width of the new window
        window_height = 250  # Height of the new window

        # Calculate the position to center the new window relative to the main window
        position_top = main_window_y + int(main_window_height / 2 - window_height / 2)
        position_left = main_window_x + int(main_window_width / 2 - window_width / 2)

        # Set the new window position
        new_window.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

        # Label and Scale for Penguin
        tk.Label(new_window, text="Number of Penguin per Box:").pack()
        w1 = tk.Scale(new_window, from_=1, to=20, orient=tk.HORIZONTAL)
        w1.set(self.num_penguin)
        w1.pack()

        # Label and Scale for Picard
        tk.Label(new_window, text="Number of Picard per Box:").pack()
        w2 = tk.Scale(new_window, from_=1, to=20, orient=tk.HORIZONTAL)
        w2.set(self.num_picard)
        w2.pack()

        tk.Button(new_window, text="Confirm",command=lambda:self.get_setting_value(w1,w2)).pack(padx=10, pady=20)

    def apply_setting(self):
        self.marker_picard = ''
        self.marker_penguin = ''
        for i in range(1,self.num_penguin + 1):
            self.marker_penguin += f'{i}:\n'

        for i in range(1,self.num_picard + 1):
            self.marker_picard += f'{i}:\n'

        self.txt_label1.config(text=self.marker_penguin)
        self.text_area1.config(height=self.num_penguin + 1)
        self.txt_label2.config(text=self.marker_picard)
        self.text_area2.config(height=self.num_picard + 1)

    def get_setting_value(self, w1, w2):
        self.num_penguin = w1.get()
        self.num_picard = w2.get()
        self.apply_setting()
        self.new_window.destroy()

    def generate_qr(self, area):
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

        if text_area == self.text_area1 and len(data_list) != self.num_penguin:
            messagebox.showwarning("Invalid Format", "Incorrect number of Penguins.")
            return

        if text_area == self.text_area2 and len(data_list) != self.num_picard:
            messagebox.showwarning("Invalid Format", "Incorrect number of Picards.")
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

        if self.button_status1 and text_area == self.text_area1:
            self.clear_penguin()

        if self.button_status2 and text_area == self.text_area2:
            self.clear_picard()

    def count_num(self, text):
        # Get the number of lines in the Text widget
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


# Create the main window
root = tk.Tk()
app = QRApp(root)

# Run the application
root.mainloop()
