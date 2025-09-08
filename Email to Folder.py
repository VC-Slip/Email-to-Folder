import win32com.client
import os
import re
import datetime

# -----------------------------
# 1️⃣ Set your OneDrive Desktop base folder path
# -----------------------------
one_drive_desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
base_folder = os.path.join(one_drive_desktop, "Bids Pending 2016", "Bids Pending", "Bids Pending")

if not os.path.exists(base_folder):
    print(f"⚠️ Base folder does not exist: {base_folder}")
    exit(1)

# -----------------------------
# 2️⃣ Connect to Outlook
# -----------------------------
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
sent_items = outlook.GetDefaultFolder(5)  # 5 = olFolderSentMail

# -----------------------------
# 3️⃣ Get the newest sent email
# -----------------------------
sent_emails = sent_items.Items
sent_emails.Sort("[ReceivedTime]", True)  # newest first
message = sent_emails.GetFirst()

if message is None:
    print("⚠️ No emails found in Sent Items.")
    exit(1)

subject = message.Subject or "No_Subject"
print(f"Subject of newest email: {subject}")

# -----------------------------
# 4️⃣ Extract quote number dynamically
# -----------------------------
quote_match = re.search(r'(Q\d+)', subject, re.IGNORECASE)
if quote_match:
    quote_number = quote_match.group(1).upper()
    print(f"Quote number found: {quote_number}")

    # -----------------------------
    # Search recursively for folder starting with Q#
    # -----------------------------
    matching_folder = None
    for root, dirs, files in os.walk(base_folder):
        for folder_name in dirs:
            if folder_name.upper().startswith(quote_number):
                matching_folder = os.path.join(root, folder_name)
                break
        if matching_folder:
            break

    if matching_folder:
        folder_path = os.path.join(matching_folder, "Correspondence")
        print(f"Matched folder: {folder_path}")
    else:
        folder_path = os.path.join(base_folder, "No_Quote_Found", "Correspondence")
        print(f"⚠️ No matching folder found, using fallback: {folder_path}")
else:
    folder_path = os.path.join(base_folder, "No_Quote_Found", "Correspondence")
    print("No quote number found in subject")

# -----------------------------
# 5️⃣ Make sure Correspondence folder exists
# -----------------------------
os.makedirs(folder_path, exist_ok=True)

# -----------------------------
# 6️⃣ Build safe filename
# -----------------------------
safe_subject = re.sub(r'[\\/*?:"<>|]', "", subject)
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
filename = f"{safe_subject}_{timestamp}.msg"
full_path = os.path.join(folder_path, filename)

# -----------------------------
# 7️⃣ Save the email
# -----------------------------
message.SaveAs(full_path, 3)  # 3 = olMSGUnicode
print(f"✅ Email saved successfully: {full_path}")
