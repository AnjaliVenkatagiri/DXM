import requests
import traceback
from datetime import datetime
import tkinter as tk
from threading import Thread
import json
from main import *

def run_bot():
    try:
        date = datetime.now().strftime("%d/%m/%Y")
        internal = entry.get()
        response = requests.get(
            f"https://careers.shahi.co.in/shahiapiprods/apiTest/getOrder?fdate=04-02-2024&tdate={date}")
        response = response.json()

        for i, row in enumerate(response):
            if f"{internal}" in str(row["dxmInternalOrderNo"]):
                print("Bot initiated")
                data = {
                    "SlNo": str(i),
                    "ReqNo": str(row["sl_NO"]),
                    "DMXOrdNo": row["dxmordno"],
                    "DMXInternal": row["dxmInternalOrderNo"],
                    "DMXDate": row["tdate"],
                    "UpdateStatus": row["dxmupdsts"],
                    "ItemNo": row["m_ITEM_NO"],
                    "CONo": row["m_ORNO"],
                    "ItemError": row["m_ITEM_ERR"],
                    "COError": row["m_ORNO_ERR"],
                    "SalesPrice": row["salesPrice"],
                    "FreightCharge": row["freightCharge"],
                }
                if row["content"] and row["styleDescription"]:
                    data["Desc"] = row["content"] + " " + row["styleDescription"]
                    print("Style Description is done!")
                print(data)
                if data["ItemNo"] and data["FreightCharge"] and data["SalesPrice"] and data["Desc"] and data["CONo"] and data["DMXOrdNo"]:
                    try:
                        print("HSN Update Started!")
                        update_hsn(data["ItemNo"])
                        print("HSN Update Over!")
                    except:
                        print("Error Updating HSN")
                        print(f"Error While Updating HSN: \n{traceback.format_exc()}")
                        return
                    try:
                        print("Update CO Started")
                        update_CO(data["CONo"], data["DMXOrdNo"], data["SalesPrice"])
                        print("Update CO Ended")
                    except:
                        print("Error Updating CO")
                        print(f"Error Updating CO Details: \n{traceback.format_exc()}")
                        return
                    try:
                        print("Plan Creation Started")
                        data["Plan"] = plan_creation(data["CONo"])
                        print("Plan Creation Ended")
                    except:
                        print("Error Creating Plan")
                        print(f"Error Generating Plan Number: \n{traceback.format_exc()}")
                        return
                    try:
                        print("Invoice Generation Started")
                        upcharge = float(data["FreightCharge"]) * 100.0 / float(data["SalesPrice"])
                        data["InvoiceDate"], data["CINumber"], data["InvoiceNo"], rosino = invoice(data["Plan"], upcharge, data["SalesPrice"], data["Desc"])
                        print("Invoice Generation Ended")
                    except:
                        print(f"Error Generating Invoice, Plan: {data['Plan']}")
                        print(f"Error Generating Invoice, The plan number is: {data['Plan']}\nexception: {traceback.format_exc()}")
                        return

                    upload_data = {
                        "invoiceM3": data["InvoiceNo"],
                        "m3InvoiceDate": str(datetime.now().strftime("%Y-%m-%dT%H:%M:%S.000")),
                        "slNo": int(data["ReqNo"]),
                        "packType": "LSE-LOOSE PACK",
                        "discriptionOfGoods": data["Desc"],
                        "rosiNumber": rosino,
                        "cifValue": 111.00,
                        "serviceNumber": str(data["Plan"])
                    }
                    bod = json.dumps(upload_data)
                    print(bod)
                    resp = requests.post(url="https://careers.shahi.co.in/shahiapiprods/apiTest/updateInvoiceNumber", data=bod, headers={"Content-Type": "application/json"})
                    print("Updated To the API")
                    print(resp.json())
                    add_data(data)
                    print("Updated To Google Sheet")
                break
        print("Ended")
    except:
        print(f"Error: {traceback.format_exc()}")

# GUI Setup
app = tk.Tk()
app.geometry("400x100")
app.title("Single Entry DXM")
label = tk.Label(app, text="Enter DXM Internal Order Number")
label.grid(row=0, column=0)
entry = tk.Entry(app)
entry.grid(row=0, column=1)
button = tk.Button(app, text="Run Bot", command=lambda: Thread(target=run_bot).start())
button.grid(row=1, column=0, columnspan=2)
app.mainloop()
