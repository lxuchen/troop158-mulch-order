import imaplib
import email
import re
import csv
import quopri
from bs4 import BeautifulSoup
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from config import EMAIL_ACCOUNT, EMAIL_PASSWORD, SENDER_EMAIL, SEARCH_YEAR
from datetime import datetime


# Function to split name into first and last name
def split_name(name):
    if name.isupper():  # Convert uppercase names to title case
        name = name.title()
    
    parts = name.split()
    if len(parts) > 2:
        first_name = " ".join(parts[:-1])  # Keep middle name with first name
        last_name = parts[-1]
    else:
        first_name, last_name = parts

    return first_name, last_name

# Function to split address into house number, street name (title case), city, and ZIP code
def split_address(address):
    # Remove "VA" from the address
    address = re.sub(r",?\s*VA\s*", "", address, flags=re.IGNORECASE)

    # Extract house number (first numeric sequence)
    house_number_match = re.match(r"(\d+)", address)
    house_number = house_number_match.group(1) if house_number_match else "Unknown"

    # Extract ZIP code (last numeric sequence)
    zip_code_match = re.search(r"(\d{5})$", address)
    zip_code = zip_code_match.group(1) if zip_code_match else "Unknown"

    # Extract street name and city
    street_city_part = re.sub(r"^\d+\s*", "", address)  # Remove house number
    street_city_part = re.sub(r"\s*\d{5}$", "", street_city_part)  # Remove ZIP code

    # Split street name and city
    parts = street_city_part.split(",")
    street_name = parts[0].strip().title() if len(parts) > 0 else "Unknown"  # Convert to title case
    city = parts[1].strip() if len(parts) > 1 else "Unknown"

    return house_number, street_name, city, zip_code

# Function to extract delivery address, delivery instructions, and scout name
def extract_delivery_info(text):
    delivery_address = "Unknown"
    delivery_instructions = "Unknown"
    scout_name = "Unknown"

    # Extract Delivery Address (text after "Delivery Address" until "Delivery Instructions")
    delivery_address_match = re.search(r"Delivery Address:\s*(.*?)\s*Delivery Instructions:", text, re.IGNORECASE | re.DOTALL)
    if delivery_address_match:
        delivery_address = delivery_address_match.group(1).strip()

    # Extract Delivery Instructions (text after "Delivery Instructions" until "Scout Name" or "Quantity")
    delivery_instructions_match = re.search(r"Delivery Instructions:\s*(.*?)(?:\s*Scout Name:|\s*Quantity:)", text, re.IGNORECASE | re.DOTALL)
    if delivery_instructions_match:
        delivery_instructions = delivery_instructions_match.group(1).strip()

    # Extract Scout Name (text after "Scout Name" until "Quantity")
    scout_name_match = re.search(r"Scout Name:\s*(.*?)\s*Quantity:", text, re.IGNORECASE | re.DOTALL)
    if scout_name_match:
        scout_name = scout_name_match.group(1).strip()

    return delivery_address, delivery_instructions, scout_name


# Connect to Gmail
print("🔄 Connecting to Gmail IMAP server...")
try:
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    print("✅ Connected to Gmail IMAP.")
    
    print("🔑 Logging in...")
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    print("✅ Login successful!")
except imaplib.IMAP4.error as e:
    print(f"❌ Login failed: {e}")
    exit()

import re

print("📂 Selecting inbox...")
mail.select("inbox")  # Select the inbox

# 🔎 Search for emails from 2025, with subject "Order #" and sender "no-reply@editmysite.com"
search_criteria = f'(SINCE "01-Jan-{SEARCH_YEAR}" FROM "{SENDER_EMAIL}" SUBJECT "Order #")'
print(f"🔎 Searching for order emails from {SEARCH_YEAR}, sent by {SENDER_EMAIL}...")
status, email_ids = mail.search(None, search_criteria)

if status != "OK" or not email_ids[0]:
    print(f"❌ No matching emails found for {SEARCH_YEAR}!")
    exit()

email_ids = email_ids[0].split()
print(f"📬 Found {len(email_ids)} order emails from {SEARCH_YEAR}.")

order_data = []

for num in email_ids:
    print(f"📥 Fetching email {num.decode()} without marking as read...")
    status, data = mail.fetch(num, "(BODY.PEEK[])")  # 👈 Prevents marking as read

    if status != "OK":
        print(f"❌ Failed to fetch email {num.decode()}. Skipping...")
        continue

    raw_email = data[0][1]
    msg = BytesParser(policy=policy.default).parsebytes(raw_email)

    # Extract subject
    subject = msg["subject"]
    print(f"📩 Email subject: {subject}")

    # Extract order number from subject
    order_match = re.search(r"Order #(\d{1,20})", subject)
    if not order_match:
        print("⚠️ No valid order number found in subject. Skipping...")
        continue
    order_number = order_match.group(1)

    # Convert email date to MM/DD/YYYY format
    raw_date = msg["Date"]
    parsed_date = email.utils.parsedate_to_datetime(raw_date)
    date_received = parsed_date.strftime("%m/%d/%Y")  # Convert to MM/DD/YYYY format

    # Extract email body (HTML content, quoted-printable decoding)
    email_html = ""
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            raw_payload = part.get_payload(decode=True)
            email_html = quopri.decodestring(raw_payload).decode(errors="ignore")
            break

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(email_html, "html.parser")
    text_content = soup.get_text("\n", strip=True)  # Extract text while preserving formatting

    # Extract customer details
    name_match = re.search(r"Customer Information\s*([A-Za-z\s]+)\n", text_content, re.DOTALL)
    email_match = re.search(r"Customer Information.*?(\S+@\S+)", text_content, re.DOTALL)
    phone_match = re.search(r"Customer Information.*?(\+?\d{10,})", text_content, re.DOTALL)

    order_quantity_match = re.search(r"Quantity:\s*(\d+)", text_content)
    total_amount_match = re.search(r"Total:\s*\$([\d,.]+)", text_content)

    # Apply formatting if a name is found
    if name_match:
        first_name, last_name = split_name(name_match.group(1).strip())
    else:
        first_name, last_name = "Unknown", "Unknown"

    # Extract delivery information
    delivery_address, delivery_instructions, scout_name = extract_delivery_info(text_content)

    if delivery_address:
        house_number, street_name, city, zip_code = split_address(delivery_address)
    else:
        house_number, street_name, city, zip_code = "Unknown", "Unknown", "Unknown", "Unknown"

    # Extract phone number, strip "+1", and format as (XXX)XXX-XXXX
    if phone_match:
        raw_phone = phone_match.group(1).strip()
        raw_phone = re.sub(r"^\+1\s*", "", raw_phone)  # Remove "+1" if present
        digits = re.sub(r"\D", "", raw_phone)  # Keep only numbers

        if len(digits) == 10:  # Ensure it's a valid 10-digit US phone number
            customer_phone = f"({digits[:3]}){digits[3:6]}-{digits[6:]}"
        else:
            customer_phone = raw_phone  # Keep original format if not 10 digits
    else:
        customer_phone = "Unknown"

    # Assign extracted values if found, otherwise set to "Unknown"
    extracted_info = {
        "First Name": first_name,
        "Last Name": last_name,
        "Order No": order_number,
        "Date Ordered": date_received,
        "Customer Phone": customer_phone,
        "Street Number": house_number,
        "Street Name": street_name,
        "City": city,
        "ZIP Code": zip_code,
        "Pipestem?": "No",
        "Customer Email": email_match.group(1).strip() if email_match else "Unknown",
        "Donation": "",
        "Order Quantity": order_quantity_match.group(1) if order_quantity_match else "Unknown",
        "Deliver": "",
        "Delivery Instructions": delivery_instructions,
        "Scout Name": scout_name,
        "Payment": "",
        "Total Amount": total_amount_match.group(1) if total_amount_match else "Unknown",
    }

    print(f"✅ Extracted Order {order_number}: {extracted_info['Last Name']}, {extracted_info['Total Amount']}")

    # Save extracted data
    order_data.append(extracted_info)

print("📤 Logging out...")
mail.logout()
print("✅ Done!")

# Save results to CSV file
csv_filename = "orders_2025.csv"
with open(csv_filename, mode="w", newline="", encoding="utf-8") as file:
    writer = csv.DictWriter(file, fieldnames=[
        "First Name", "Last Name", "Order No", "Date Ordered", "Customer Phone", "Street Number",
        "Street Name", "City", "ZIP Code", "Pipestem?",
        "Customer Email", "Donation" , "Order Quantity", "Deliver", 
        "Delivery Instructions", "Scout Name", "Payment", "Total Amount"
    ])
    writer.writeheader()
    # Ensure all fields are single-line strings before writing to CSV
    for order in order_data:
        for key in order:
            if isinstance(order[key], str):
                # Remove newlines, tabs, non-breaking spaces, and excess spaces
                order[key] = re.sub(r"[\n\r\t\xa0]+", " ", order[key]).strip()

                # Ensure no trailing spaces or weird hidden characters
                order[key] = re.sub(r"\s{2,}", " ", order[key])  # Replace multiple spaces with a single space
    # Write cleaned data to CSV
    writer.writerows(order_data)

print(f"📁 Data saved to {csv_filename}")
