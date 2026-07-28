from flask import Flask, request
import requests
import os

app = Flask(__name__)

# ---------------------------
# CONFIGURATION VIA RENDER
# ---------------------------
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
# ---------------------------

def get_ms_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    response = requests.post(url, data=data).json()
    return response["access_token"]

@app.route("/brevo-sync", methods=["POST"])
def brevo_sync():
    data = request.json

    email = data.get("email")
    first_name = data.get("attributes", {}).get("FIRSTNAME", "")
    last_name = data.get("attributes", {}).get("LASTNAME", "")

    token = get_ms_token()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "givenName": first_name,
        "surname": last_name,
        "emailAddresses": [{"address": email}]
    }

    response = requests.post(
        "https://graph.microsoft.com/v1.0/me/contacts",
        headers=headers,
        json=payload
    )

    if response.status_code == 201:
        return "Contact ajouté dans Outlook", 201
    else:
        return f"Erreur Outlook: {response.text}", 400

@app.route("/", methods=["GET"])
def home():
    return "Brevo → Outlook Sync actif"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
