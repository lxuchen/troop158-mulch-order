# 📩 Troop158 Mulch Order Extractor

This Python script automates extracting order details from Gmail order confirmation emails. It parses structured information from emails and saves it into a CSV file.

## 🚀 Features
- Retrieves **order emails** from Gmail.
- Extracts **customer details, delivery instructions, and scout name**.
- Properly **formats phone numbers and addresses**.
- Saves data to **CSV** in **chronological order** (oldest email first).

---

## 🔑 **1. Generate Google App Password**
Since Gmail blocks less secure apps, you **must enable 2FA and generate an App Password**.

### **Steps to Generate an App Password:**
1. **Enable 2-Step Verification** on your Google Account:
   - Go to [Google My Account](https://myaccount.google.com/security).
   - Scroll down to **"Signing in to Google"** and enable **2-Step Verification**.
2. **Generate an App Password**:
   - In **"Signing in to Google"**, click **"App passwords"**.
   - Select **Mail** as the app and **Other (Custom Name)**.
   - Enter a name like `"Gmail Order Extractor"` and click **Generate**.
   - Copy the **16-character password** (remove spaces when pasting into `config.py`).

---

## 💻 **2. Setup Virtual Environment & Install Dependencies**
To isolate dependencies, use a **Python virtual environment**.

### **Steps:**
1. **Create & activate a virtual environment**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On macOS/Linux
   venv\Scripts\activate     # On Windows
2. **Install required dependencies**
   ```bash
   pip install -r requirements.txt
3. **Configure config.py**
   The script reads email credentials and search filters from config.py
   Edit config.py and fill in:
   ```python
   # Gmail credentials (Use App Password, not regular password)
   EMAIL_ACCOUNT = "your-email@gmail.com"
   EMAIL_PASSWORD = "your-app-password"  # Remove spaces from the app password

   # Email filtering criteria
   SENDER_EMAIL = "no-reply@example.com"  # Only process emails from this sender
   SEARCH_YEAR = "2025"  # Only process emails sent in this year

4. **Run the Script**
   Once configured, run:
   ```bash
   python read.py
   ```
   This will  
   ✅ Connect to Gmail  
   ✅ Fetch only 2025 orders from no-reply@example.com  
   ✅ Extract order details, customer information, and delivery instructions  
   ✅ Save the output as orders_2025.csv

 5. **Output: CSV Format**
    After running, the script creates orders_2025.csv, with the following columns:
    | Order No  | Date Ordered | First Name | Last Name | House Number | Street Name  | City      | ZIP Code | Order Quantity | Total Amount | Delivery Instructions | Scout Name  |
    |-----------|-------------|------------|-----------|--------------|--------------|-----------|----------|----------------|--------------|-----------------------|-------------|
    | 123456789 | 03/12/2025  | John       | Doe       | 1234         | Elm Street  | Fairfax   | 22030    | 2              | $45.00       | Leave by front door   | Jane Smith  |
    | 987654321 | 03/15/2025  | Alice      | Johnson   | 5678         | Maple Ave   | Herndon   | 20170    | 1              | $20.00       | Use side entrance     | Bob Taylor  |

   
