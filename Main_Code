import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter import filedialog, messagebox, simpledialog

# Default email templates
templates = {
    "Template 1": "Hello {Client_Name},\n\nWe are excited to share with you our new tool – the **Automatic Email Sender**.\nThis application reads an Excel file containing client details (like your name and email) and automatically sends personalized messages.  ",
    "Template 2": "Dear {Client_Name},\n\nWe are pleased to inform you about your job title: {job_title}.",
    "Template 3": "Hi {Client_Name},\n\nJust a quick note to confirm your job title as {job_title}.",
}

# Global variable to store the file path
file_path = None

# Function to login to the email server
def login():
    global your_email, your_password

    your_email = simpledialog.askstring("Login", "Enter your email:")
    your_password = simpledialog.askstring("Login", "Enter your password:", show='*')

    if not your_email or not your_password:
        messagebox.showwarning("Input Error", "Please enter both email and password.")
        return
    
    try:
        # Establish connection with Gmail server
        global server
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(your_email, your_password)
        login_status_var.set("Login successful")
    except Exception as e:
        login_status_var.set(f"Login failed: {str(e)}")
        messagebox.showerror("Login Error", f"Failed to login: {str(e)}")

# Function to upload the Excel file
def upload_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_status_var.set(f"File loaded: {file_path}")
    else:
        file_status_var.set("No file selected")

# Function to send emails
def send_emails():
    if not file_path:
        messagebox.showwarning("No File", "Please upload an Excel file before sending emails.")
        return

    try:
        # Load the Excel file
        email_list = pd.read_excel(file_path)

        # Ensure columns are strings and handle missing values
        email_list['Client_Name'] = email_list['Client_Name'].fillna('')
        email_list['Emails'] = email_list['Emails'].fillna('')
        #email_list['job_title'] = email_list['job_title'].fillna('')

        # Add a new column for email sent status if not already present
        if 'email_sent' not in email_list.columns:
            email_list['email_sent'] = ''

        # Get selected template
        selected_template = template_var.get()

        # Initialize counters
        sent_count = 0
        failed_count = 0

        # Iterate through the records and send emails
        for i in range(len(email_list)):
            Client_Name = email_list['Client_Name'][i]
            email = email_list['Emails'][i]
            #job_title = email_list['job_title'][i]

            # Skip if the email is missing or invalid
            if pd.isna(email) or email.strip() == "":
                continue

            # Create the email message
            message = MIMEMultipart()
            message['From'] = your_email
            message['To'] = email
            message['Subject'] = "Mail from Harshal"
            
            try:
                body = templates[selected_template].format(
                    Client_Name = Client_Name)
                    # ,job_title=job_title   # enable later if needed)
                
            except KeyError as e:
                messagebox.showerror("Error", f"Missing placeholder in template: {e}")
                continue

            message.attach(MIMEText(body, 'plain'))

            try:
                # Send the email
                server.sendmail(your_email, email, message.as_string())
                email_list.at[i, 'email_sent'] = 'email sent'
                sent_count += 1
            except Exception as e:
                email_list.at[i, 'email_sent'] = f'failed: {str(e)}'
                failed_count += 1

        # Save the updated DataFrame back to the Excel file
        email_list.to_excel(file_path, index=False)

        # Update status on the GUI
        status_var.set(f"Sent: {sent_count}, Failed: {failed_count}")

        # Show success message
        messagebox.showinfo("Success", f"Emails sent: {sent_count}\nFailed: {failed_count}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to update the template text area when a template is selected
def update_template_text(*args):
    selected_template = template_var.get()
    template_text.delete(1.0, tk.END)
    template_text.insert(tk.END, templates[selected_template])

# Function to save the edited template
def save_template():
    selected_template = template_var.get()
    templates[selected_template] = template_text.get(1.0, tk.END).strip()
    messagebox.showinfo("Template Saved", f"The template '{selected_template}' has been updated.")

# Function to add a new custom template
def add_template():
    new_template_name = simpledialog.askstring("New Template", "Enter the name for the new template:")

   
    if new_template_name:
        # Create a popup window for template content
        popup = tk.Toplevel(root)
        popup.title(f"Enter content for {new_template_name}")
        popup.geometry("500x400")  # size similar to main text box

        # Text area for template content
        content_text = tk.Text(popup, height=15, width=60)
        content_text.pack(pady=10)
        # Function to save the new template
        def save_and_close():
            new_template_content = content_text.get(1.0, tk.END).strip()
            if new_template_content:
                # Save to templates dictionary
                templates[new_template_name] = new_template_content
                # Update the dropdown menu
                template_menu['menu'].add_command(
                    label=new_template_name, 
                    command=tk._setit(template_var, new_template_name)
                )
                template_var.set(new_template_name)
                update_template_text()
                messagebox.showinfo("Template Added", f"Template '{new_template_name}' has been added.")
                popup.destroy()
            else:
                messagebox.showwarning("Empty Content", "Please enter template content before saving.")

        # Save button in the popup
        tk.Button(popup, text="Save Template", command=save_and_close).pack(pady=5)        

# Function to delete the selected template
def delete_template():
    selected_template = template_var.get()
    
    # Prevent deletion if there's only one template left
    if len(templates) == 1:
        messagebox.showwarning("Cannot Delete", "You must have at least one template.")
        return
    
    # Ask for confirmation
    confirm = messagebox.askyesno("Delete Template", f"Are you sure you want to delete the template '{selected_template}'?")
    
    if confirm:
        # Delete the selected template
        del templates[selected_template]
        
        # Update the dropdown menu by rebuilding it
        menu = template_menu['menu']
        menu.delete(0, 'end')
        for template_name in templates.keys():
            menu.add_command(label=template_name, command=tk._setit(template_var, template_name))
        
        # Set the dropdown to the first available template
        new_selection = list(templates.keys())[0]
        template_var.set(new_selection)
        
        # Update the template text area
        update_template_text()

        messagebox.showinfo("Template Deleted", f"The template '{selected_template}' has been deleted.")




# Setting up the GUI
root = tk.Tk()
root.title("Automatic Email Sender")
root.option_add("*Font", "Arial 10")

style = ttk.Style()
style.theme_use("clam")

# === 1. LOGIN & FILE SECTION ===
login_frame = ttk.LabelFrame(root, text="Login & File", padding=10)
login_frame.pack(fill="x", padx=10, pady=5)

# Button to log in to the email server

login_button = ttk.Button(login_frame, text="Login", command=login)
login_button.pack(side=tk.LEFT, padx=5)

# Add a label for the login status
login_status_var = tk.StringVar(value="Not logged in")
login_status_label = ttk.Label(login_frame, textvariable=login_status_var, foreground="blue")
login_status_label.pack(side=tk.LEFT, padx=10)

upload_button = ttk.Button(login_frame, text="Upload Excel File", command=upload_file)
upload_button.pack(side=tk.LEFT, padx=5)

# Add a label to show the file upload status
file_status_var = tk.StringVar(value="No file uploaded")
file_status_label = ttk.Label(login_frame, textvariable=file_status_var, foreground="green")
file_status_label.pack(side=tk.LEFT, padx=10)

# === 2. TEMPLATE SETTINGS SECTION ===
template_frame = ttk.LabelFrame(root, text="Template Settings", padding=10)
template_frame.pack(fill="both", expand=True, padx=10, pady=5)

# FRAME for formatting buttons
format_frame = ttk.Frame(template_frame)
format_frame.pack(pady=2)

# Bold button
bold_button = ttk.Button(format_frame, text="B", command=lambda: apply_tag("bold"))
bold_button.pack(side=tk.LEFT, padx=2)

# Italic button
italic_button = ttk.Button(format_frame, text="I", command=lambda: apply_tag("italic"))
italic_button.pack(side=tk.LEFT, padx=2)

# Underline button
underline_button = ttk.Button(format_frame, text="U", command=lambda: apply_tag("underline"))
underline_button.pack(side=tk.LEFT, padx=2)

# Red color button
red_button = ttk.Button(format_frame, text="Red", command=lambda: apply_color("red"))
red_button.pack(side=tk.LEFT, padx=2)

# Create a dropdown menu for selecting the email template
template_var = tk.StringVar(root)
template_var.set("Template 1") #Default Template
template_var.trace("w", update_template_text)  # Update text area on selection

template_menu = ttk.Combobox(template_frame, textvariable=template_var, values=list(templates.keys()))
template_menu.pack(pady=5)


# Text area to display and edit the selected template
template_text = tk.Text(template_frame, height=10, width=60,bg="#ffffff", font=("Arial", 11))
template_text.pack(pady=5)
update_template_text()  # Load the default template into the text area

# Configure tags for template_text
template_text.tag_configure("bold", font=("Arial", 10, "bold"))
template_text.tag_configure("italic", font=("Arial", 10, "italic"))
template_text.tag_configure("underline", font=("Arial", 10, "underline"))
template_text.tag_configure("red", foreground="red")

def apply_tag(tag_name):
    try:
        start = template_text.index("sel.first")
        end = template_text.index("sel.last")
        template_text.tag_add(tag_name, start, end)
    except tk.TclError:
        pass  # No text selected

def apply_color(color):
    try:
        start = template_text.index("sel.first")
        end = template_text.index("sel.last")
        template_text.tag_configure(color, foreground=color)
        template_text.tag_add(color, start, end)
    except tk.TclError:
        pass  # No text selected

# === 3. TEMPLATE MANAGEMENT SECTION ===
manage_frame = ttk.LabelFrame(root, text="Manage Templates", padding=10)
manage_frame.pack(fill="x", padx=10, pady=5)


# Button to add a new template
add_template_button = ttk.Button(manage_frame, text="Add New Template", command=add_template)
add_template_button.pack(side=tk.LEFT, padx=5)

save_button = ttk.Button(manage_frame, text="Save Template", command=save_template)
save_button.pack(side=tk.LEFT, padx=5)

# Button to delete the selected template
delete_template_button = ttk.Button(manage_frame, text="Delete Template", command=delete_template)
delete_template_button.pack(side=tk.LEFT, padx=5)


# === 4. SEND SECTION ===
send_frame = ttk.LabelFrame(root, text="Send Emails", padding=10)
send_frame.pack(fill="x", padx=10, pady=5)

#ttk doesn’t support bg directly → need style
style.configure("Send.TButton", background="#4CAF50", foreground="white")

send_button = ttk.Button(send_frame, text="Send Emails", style="Send.TButton",command=send_emails)
send_button.pack(side=tk.LEFT, padx=5)

status_var = tk.StringVar(value="No emails sent yet.")
status_label = tk.Label(send_frame, textvariable=status_var, fg="green")
status_label.pack(side=tk.LEFT, padx=10)

# === 5. FORMAT UI ===
style.configure("Send.TButton",
                background="#4CAF50",     # green bg
                foreground="white",       # white text
                bordercolor="#4CAF50",    # green border
                focusthickness=3,
                focuscolor="#4CAF50")

style.map("Send.TButton",
          background=[("active", "#45a049"), ("pressed", "#3e8e41")],  # darker green on hover/press
          foreground=[("disabled", "gray")],
          bordercolor=[("active", "#45a049"), ("pressed", "#3e8e41")])

# General frame style
style.configure("TLabelframe", background="#ffffff", borderwidth=2, relief="groove")
style.configure("TLabelframe.Label", background="#ffffff", foreground="#333333", font=("Arial", 10, "bold"))

# General button style
style.configure("TButton",
                background="#e6e6e6",
                foreground="#000000",
                padding=5)
style.map("TButton",
          background=[("active", "#d9d9d9"), ("pressed", "#cccccc")])

# Text widget background
template_text.config(bg="#ffffff", fg="#111111", insertbackground="black")

file_paths = []
output_dir = ""

root.mainloop()