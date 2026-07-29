from flask import Flask, request, jsonify
import requests
import os

app = Flask(__name__)

# ============================================================
# CONFIGURATION
# ============================================================

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
USER_EMAIL = os.getenv("USER_EMAIL")

# ============================================================
# DÉMARRAGE
# ============================================================

print("\n===================================")
print(" Brevo → Microsoft Sync v3.1")
print("===================================")

for variable in [
    "TENANT_ID",
    "CLIENT_ID",
    "CLIENT_SECRET",
    "USER_EMAIL"
]:
    valeur = os.getenv(variable)
    print(f"{'✓' if valeur else '✗'} {variable}")

print("===================================\n")


# ============================================================
# MICROSOFT GRAPH TOKEN
# ============================================================

def get_token():

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    response = requests.post(
        url,
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials"
        }
    )

    print("\n========== AZURE ==========")
    print("Status :", response.status_code)

    response.raise_for_status()

    token = response.json()["access_token"]

    print("Jeton obtenu.")

    return token


# ============================================================
# WEBHOOK BREVO
# ============================================================

@app.route("/brevo-sync", methods=["POST"])
def brevo():

    data = request.get_json(force=True)

    print("\n==============================")
    print("WEBHOOK BREVO")
    print("==============================")
    print(data)

    event = data.get("event", "")

    print("Événement :", event)

    # --------------------------------------------------------

    if event == "contact_deleted":

        print("Événement ignoré.")

        return jsonify({
            "ignored": True,
            "reason": "contact_deleted"
        }), 200

    if event not in [
        "contact_created",
        "contact_updated",
        "list_addition"
    ]:

        print("Événement inconnu.")

        return jsonify({
            "ignored": True,
            "event": event
        }), 200

    # --------------------------------------------------------
    # Email
    # --------------------------------------------------------

    email = data.get("email")

    if isinstance(email, list):
        email = email[0] if email else None

    if not email:

        print("Aucun email.")

        return jsonify({
            "error": "Email manquant."
        }), 400

    # list_addition n'a pas de bloc "attributes" (pas de FIRSTNAME/LASTNAME)
    attributes = data.get("attributes", {})

    firstname = attributes.get("FIRSTNAME", "")
    lastname = attributes.get("LASTNAME", "")

    print("\nCONTACT")
    print("Email :", email)
    print("Prénom :", firstname)
    print("Nom :", lastname)

    # --------------------------------------------------------

    if not USER_EMAIL:

        print("USER_EMAIL absent.")

        return jsonify({
            "error": "Variable USER_EMAIL absente."
        }), 500

    # --------------------------------------------------------

    token = get_token()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "givenName": firstname,
        "surname": lastname,
        "emailAddresses": [
            {
                "address": email,
                "name": f"{firstname} {lastname}".strip() or email
            }
        ]
    }

    print("\nPAYLOAD")
    print(payload)

    url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/contacts"

    response = requests.post(
        url,
        headers=headers,
        json=payload
    )

    print("\n========== GRAPH ==========")
    print("Status :", response.status_code)
    print(response.text)

    if response.status_code == 201:

        print("\n✓ Contact créé.")

        return jsonify({
            "success": True
        }), 201

    print("\n✗ Erreur Graph.")

    return jsonify({
        "success": False,
        "status": response.status_code,
        "graph": response.text
    }), response.status_code


# ============================================================
# HOME
# ============================================================

@app.route("/", methods=["GET"])
def home():
    return "Brevo Sync v3.1 OK"


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
