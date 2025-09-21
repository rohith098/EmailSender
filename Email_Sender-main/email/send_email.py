import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import re
import os
def validate_email(email):
    """Validate the email format."""
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(email_regex, email))

def clean_email(value):
    """Clean and validate email values."""
    try:
        if pd.isna(value):  # Handle NaN or None
            return None
        email = str(value).strip()  # Convert to string and strip whitespace
        return email if validate_email(email) else None
    except Exception as e:
        print(f"Error processing email value '{value}': {e}")
        return None

def read_emails_from_file(file_name):
    """Read and validate email addresses from a file."""
    to_emails = set()
    try:
        if file_name.endswith(".csv"):
            df = pd.read_csv(file_name)
        elif file_name.endswith((".xls", ".xlsx")):
            df = pd.read_excel(file_name)
        else:
            messagebox.showerror("Error", "Unsupported file format. Please provide a CSV or Excel file.")
            return to_emails

        for _, row in df.iterrows():
            to_email = clean_email(row.get("To"))
            if to_email:
                to_emails.add(to_email)
        return to_emails

    except FileNotFoundError:
        messagebox.showerror("Error", f"The file '{file_name}' was not found.")
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the file: {e}")
    return to_emails

def send_email(sender_email, sender_password, subject, body, file_name, to_emails):
    """Send emails with an attachment."""
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)

            for recipient in to_emails:
                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = recipient
                message["Subject"] = subject

                message.attach(MIMEText(body, "plain"))

                if file_name:
                    with open(file_name, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            "Content-Disposition",
                            f"attachment; filename={os.path.basename(file_name)}",
                        )
                        message.attach(part)

                server.sendmail(sender_email, recipient, message.as_string())

            messagebox.showinfo("Success", f"Emails sent successfully to {len(to_emails)} recipients.")

    except Exception as e:
        messagebox.showerror("Error", f"Error sending email: {e}")

def select_file(entry):
    """Open a file dialog and allow the user to select a file."""
    file_name = filedialog.askopenfilename(
        title="Select a File",
        filetypes=(("CSV Files", ".csv"), ("Excel Files", ".xls;*.xlsx"))
    )
    if file_name:
        entry.delete(0, tk.END)
        entry.insert(0, file_name)

def preview_recipients(file_entry):
    """Preview the list of recipients."""
    file_name = file_entry.get()
    if not file_name:
        messagebox.showerror("Error", "Please select a file.")
        return
    recipients = read_emails_from_file(file_name)
    if recipients:
        messagebox.showinfo("Recipients", f"Valid email addresses:\n\n{', '.join(recipients)}")
    else:
        messagebox.showinfo("Recipients", "No valid email addresses found.")

def main():
    def on_send():
        sender_email = sender_email_entry.get()
        sender_password = sender_password_entry.get()
        subject = subject_entry.get()
        body = body_text.get("1.0", tk.END).strip()
        file_name = attachment_entry.get()
        recipient_file = recipient_file_entry.get()

        if not sender_email or not sender_password:
            messagebox.showerror("Error", "Please enter your email credentials.")
            return
        if not recipient_file:
            messagebox.showerror("Error", "Please select a recipient list file.")
            return

        to_emails = read_emails_from_file(recipient_file)
        if not to_emails:
            messagebox.showerror("Error", "No valid email addresses found in the file.")
            return

        send_email(sender_email, sender_password, subject, body, file_name, to_emails)

    root = tk.Tk()
    root.title("Email Distribution System")

    # Sender Email
    tk.Label(root, text="Sender_Email: ").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    sender_email_entry = tk.Entry(root, width=40)
    sender_email_entry.grid(row=0, column=1, padx=10, pady=5)

    # Sender Password
    tk.Label(root, text="Password: ").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    sender_password_entry = tk.Entry(root, width=40, show="*")
    sender_password_entry.grid(row=1, column=1, padx=10, pady=5)

    # Email Subject
    tk.Label(root, text="Subject:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    subject_entry = tk.Entry(root, width=40)
    subject_entry.grid(row=2, column=1, padx=10, pady=5)

    # Email Body
    tk.Label(root, text="Body:").grid(row=3, column=0, padx=10, pady=5, sticky="nw")
    body_text = tk.Text(root, width=40, height=10)
    body_text.grid(row=3, column=1, padx=10, pady=5)

    # Recipient File
    tk.Label(root, text="Recipient List File:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
    recipient_file_entry = tk.Entry(root, width=40)
    recipient_file_entry.grid(row=4, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: select_file(recipient_file_entry)).grid(row=4, column=2, padx=10, pady=5)

    # Preview Recipients
    tk.Button(root, text="Preview Recipients", command=lambda: preview_recipients(recipient_file_entry)).grid(row=5, column=1, padx=10, pady=5)

    # Attachment File
    tk.Label(root, text="Attachment File (optional):").grid(row=6, column=0, padx=10, pady=5, sticky="w")
    attachment_entry = tk.Entry(root, width=40)
    attachment_entry.grid(row=6, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda: select_file(attachment_entry)).grid(row=6, column=2, padx=10, pady=5)

    # Send Button
    tk.Button(root, text="Send Emails", command=on_send, bg="green", fg="white").grid(row=7, column=1, padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
