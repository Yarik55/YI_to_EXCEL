#Load Python's built-in JSON toolbox.
import json
#"Load the requests toolbox so we can talk to Xero's API."
import requests

#This function does one job — builds the authentication headers that every single Xero API call needs. Instead of repeating this code in every function, we write it once here and call it whenever we need it.
def get_headers():
    with open("token.json") as f:
        tokens = json.load(f)
    
    access_token = tokens["access_token"]
    
    connections = requests.get(
        "https://api.xero.com/connections",
        headers={"Authorization": f"Bearer {access_token}"}
    ).json()
    
    tenant_id = connections[0]["tenantId"]
    
    return {
        "Authorization": f"Bearer {access_token}",
        "Xero-tenant-id": tenant_id,
        "Accept": "application/json"
    }

#"Store Xero's base API address in a variable called BASE."

BASE = "https://api.xero.com/api.xro/2.0/"

#"Define a function called get_profit_and_loss."

def get_profit_and_loss ():

#"Send a GET request to Xero's P&L endpoint with our auth headers and store the response in r."

    r = requests.get(
        f"{BASE}/reports/ProfitAndLoss",
        headers=get_headers()

    )
    return r.json()["Reports"][0]

def get_balance_sheet ():
    r = requests.get(
        f"{BASE}/reports/BalanceSheet",
        headers=get_headers()

    )
    return r.json()["Reports"][0]

def get_invoices():
    r = requests.get(
        f"{BASE}/invoices?order=Date DESC",
        headers=get_headers()

    )
    return r.json()["Invoices"]

def get_bank_transactions():
    r = requests.get(
        f"{BASE}/BankTransactions?order=Date DESC",
        headers=get_headers()

    )
    return r.json()["BankTransactions"]
