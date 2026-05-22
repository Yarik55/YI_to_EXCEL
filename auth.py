import os # "Give me access to the operating system."
import requests #"Give me the toolbox for talking to the internet."
import webbrowser #"Give me the ability to open a browser window."

from flask import Flask, request, redirect #"From the Flask toolbox, specifically give me just these 3 tools: Flask, request, and redirect."
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__) #"Open the .env file and load everything inside it into memory right now."

CLIENT_ID = os.getenv("XERO_CLIENT_ID")
CLIENT_SECRET = os.getenv("XERO_CLIENT_SECRET")
REDIRECT_URI = os.getenv("XERO_REDIRECT_URI")
SCOPES = "openid profile email offline_access accounting.invoices.read accounting.banktransactions.read accounting.reports.profitandloss.read accounting.reports.balancesheet.read accounting.settings"
#"When someone visits the home page of our mini web server, run the function below this line."

@app.route("/")
def login():
    url = (
        "https://login.xero.com/identity/connect/authorize"
        "?response_type=code"
        f"&client_id={CLIENT_ID}"
        f"&redirect_uri={REDIRECT_URI}"
        f"&scope={SCOPES}"
    )
    return redirect(url)

#"When Xero sends the browser back to /callback, run the function below."

@app.route("/callback")

#"Define the function that handles Xero's response."

def callback():

#"Grab the authorisation code that Xero attached to the URL."

    code = request.args.get("code")
    
    token_response = requests.post(
        "https://identity.xero.com/connect/token",
        data={
            "grant_type":   "authorization_code",
            "code":         code,
            "redirect_uri": REDIRECT_URI,
        },
        auth=(CLIENT_ID, CLIENT_SECRET)
    )
    
    tokens = token_response.json()
    
    import json
    with open("token.json", "w") as f:
        json.dump(tokens, f)
    
    return "✅ Logged in! Close this tab and run xero_to_excel.py"


if __name__ == "__main__":
    webbrowser.open("http://localhost:5000")
    app.run(port=5000, ssl_context="adhoc")
