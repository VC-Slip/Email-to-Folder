import win32com.client
import os
import re
import datetime
import shutil

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
# -----------------------------
# 4️⃣ Extract quote number dynamically
# -----------------------------
matching_folder = None
quote_match = re.search(r'(Q\d+)', subject, re.IGNORECASE)
if quote_match:
    quote_number = quote_match.group(1).upper()
    print(f"Quote number found: {quote_number}")

    # -----------------------------
    # Search recursively for folder containing the quote number
    # -----------------------------
    for root, dirs, files in os.walk(base_folder):
        for folder_name in dirs:
            # Match even if folder has initials like "VC Q24024 ..."
            if re.search(rf'\b{quote_number}\b', folder_name, re.IGNORECASE):
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

# -----------------------------
# 8️⃣ Move the quote folder into the client’s folder (if found)
# -----------------------------
if matching_folder and os.path.exists(matching_folder):
    quote_folder_name = os.path.basename(matching_folder)

    # Try to extract client name (word after Q####)
    client_name_match = re.search(r'Q\d+\s+([A-Za-z]+)', quote_folder_name)
    if client_name_match:
        client_name = client_name_match.group(1)
        print(f"Detected client name: {client_name}")

        # Search for a matching client folder under "Bids Pending 2016"
        client_root = os.path.join(one_drive_desktop, "Bids Pending 2016")
        target_folder = None
        for folder in os.listdir(client_root):
            if client_name.lower() in folder.lower() and os.path.isdir(os.path.join(client_root, folder)):
                target_folder = os.path.join(client_root, folder)
                break

        if target_folder:
            destination = os.path.join(target_folder, quote_folder_name)
            if not os.path.exists(destination):
                print(f"📦 Moving folder:\n  From: {matching_folder}\n  To:   {target_folder}")
                shutil.move(matching_folder, destination)
                print(f"✅ Folder moved successfully to: {destination}")

                # -----------------------------
                # 9️⃣ Rename after move (remove initials like 'VC ')
                # -----------------------------
                folder_name_only = os.path.basename(destination)
                parent_folder = os.path.dirname(destination)
                clean_folder_name = re.sub(r'^[A-Z]{2,3}\s+', '', folder_name_only)
                new_path = os.path.join(parent_folder, clean_folder_name)

                if clean_folder_name != folder_name_only:
                    try:
                        os.rename(destination, new_path)
                        print(f"🧹 Renamed folder to: {new_path}")
                    except Exception as e:
                        print(f"⚠️ Could not rename folder: {e}")
            else:
                print(f"⚠️ Destination already exists: {destination}, skipping move.")
        else:
            print(f"⚠️ No matching client folder found for '{client_name}', skipping move.")
    else:
        print("⚠️ Could not detect client name from folder name.")
else:
    print("⚠️ No matching quote folder found, skipping move.")
