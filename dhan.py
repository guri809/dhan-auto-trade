from dhanhq import dhanhq
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import sys

# -----------------------------
# GOOGLE SHEET CONNECTION
# -----------------------------
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    "advance-maker-388815-a10824329a8d.json", scope)

client = gspread.authorize(creds)

sheet_url = "https://docs.google.com/spreadsheets/d/1yAOcAdgVQ-Qbj_2-zLbZqk_l4niY49MTiZ_FvBx5F24/edit"
sheet = client.open_by_url(sheet_url).sheet1

print("Connected to Google Sheet ✔")

# -----------------------------
# READ CLIENT ID + TOKEN FROM SHEET
# -----------------------------
client_id = "1108518198"   # FIXED - stays in script always

access_token = sheet.acell("M2").value  # Read token from Column M

if not access_token:
    print("ERROR: No token found in Column M2 ❌")
    sys.exit()

print("Token Loaded Successfully ✔")

# -----------------------------
# DHAN LOGIN
# -----------------------------
dhan = dhanhq(client_id, access_token)

# -----------------------------
# READ SECURITY ID + QTY
# -----------------------------
security_ids = sheet.col_values(15)[1:]  # O column
quantities = sheet.col_values(14)[1:]    # N column

# -----------------------------
# RUN ONLY FRIDAY 3:28 PM
# -----------------------------
now = datetime.datetime.now()

if now.weekday() != 4:  # Only Friday
    print("Not Friday → No Trades")
    sys.exit()

if not (now.hour == 15 and now.minute >= 28):
    print("Not 3:28 PM → No Trades")
    sys.exit()

print("Friday 3:28 PM → Starting Auto Orders...")

# -----------------------------
# PLACE ORDERS
# -----------------------------
for sec, qty in zip(security_ids, quantities):
    if not sec or not qty:
        continue

    try:
        qty = int(float(qty))

        print(f"BUY ORDER → Security ID={sec} | Qty={qty}")

        order = dhan.place_order(
            security_id=sec,
            exchange_segment=dhan.NSE,
            transaction_type=dhan.BUY,
            order_type=dhan.MARKET,
            product_type=dhan.CNC,
            quantity=qty
        )

        print("SUCCESS:", order)

    except Exception as e:
        print("FAILED:", sec, e)

print("All Orders Completed ✔")
