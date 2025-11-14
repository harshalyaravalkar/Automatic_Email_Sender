# 📧 Automatic Email Sender (Tkinter GUI)

A simple and efficient **Python desktop application** that automates sending personalized bulk emails using data from an Excel file.  
It provides a **clean Tkinter GUI** to log in, upload data, create or edit templates, and send emails — all without needing to touch code.

---

## 🌟 Features

✅ **Easy-to-use GUI** — built using Tkinter, no coding needed.  
✅ **Upload Excel File** — automatically reads client names and email IDs.  
✅ **Dynamic Email Templates** — personalize each email using placeholders like `{Client_Name}`.  
✅ **Template Management** — add, edit, or delete templates from within the app.    
✅ **Real-time Status Display** — view sent/failed count in the GUI.  
✅ **Automatic Excel Update** — updates each row with “email sent” or error details.  
✅ **Secure Login System** — authenticate once and send multiple emails.

---

## 🧩 How It Works

1. **Login** using your sender email credentials (or Gmail app password).  
2. **Upload Excel File** containing your client list.  
3. **Choose or Edit Template** from the dropdown menu.  
4. **Use Format Buttons** — apply **bold**, *italic*, underline, or text color.  
5. **Click “Send Emails”** — emails are sent automatically to each address.  
6. The Excel file gets updated with the delivery status for each row.

---

## 🧾 Example Excel File Format

| Client_Name | Emails             |
|--------------|--------------------|
| John Doe     | johndoe@gmail.com |
| Jane Smith   | janesmith@yahoo.com |

Make sure your Excel file includes these exact columns:  
- `Client_Name`  
- `Emails`

---

## ✉️ Example Email Template

You can include placeholders in your templates (inside the text editor):
<img width="660" height="744" alt="Screenshot 2025-10-17 171801" src="https://github.com/user-attachments/assets/d29699b3-380a-46f6-8fda-e0e6a1fc0043" />

## Email Login Credentials
You can use these credentials to log in when you install and try to use the AutoEmail.exe file
**Gmail ID** — demouser012001@gmail.com
**Login Password** — zccp bewk emvp bvwf

## ⚙️ Installation

### Step 1. Clone the Repository

```bash
git clone https://github.com/harshalyaravalkar/Automatic_Email_Sender.git
cd Atomatic_Email_Sender

