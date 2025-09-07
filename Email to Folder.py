import win32com.client
import os
import re
import datetime

# Set your base "Bids" folder path
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
base_folder = os.path.join(desktop_path, "Bids pending 2016", "Bids Pending", "Bids Pending")

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
sent_items = outlook.GetDefaultFolder(5)  # 5 = olFolderSentMail

# Get the most recent sent email
message = sent_items.Items.GetLast()

# Try to extract the quote number (e.g., Q24024) from the subject
subject = message.Subject or "No_Subject"
quote_match = re.search(r'(Q\d+)', subject, re.IGNORECASE)

if quote_match:
    quote_number = quote_match.group(1).upper()  # normalize to uppercase

    # Look for an existing folder starting with the quote number
    matching_folder = None
    for folder_name in os.listdir(base_folder):
        if folder_name.upper().startswith(quote_number):
            matching_folder = folder_name
            break

    if matching_folder:
        folder_path = os.path.join(base_folder, matching_folder, "Correspondence")
    else:
        # If no folder exists, create one with just the quote number
        matching_folder = quote_number
        folder_path = os.path.join(base_folder, matching_folder, "Correspondence")

else:
    # If no quote number found, dump into a general folder
    folder_path = os.path.join(base_folder, "No_Quote_Found", "Correspondence")

# Make sure Correspondence folder exists
os.makedirs(folder_path, exist_ok=True)

# Clean subject for filename
safe_subject = re.sub(r'[\\/*?:"<>|]', "", subject)
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
filename = f"{safe_subject}_{timestamp}.msg"
full_path = os.path.join(folder_path, filename)

# Save the email
message.SaveAs(full_path)
print(f"Email saved to {full_path}")
