from flask import Flask, request, jsonify
import requests
import os

app = Flask(__name__)

# ==========================================
# CONFIGURATION (Variables Render)
# ==========================================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
USER_EMAIL = os.getenv("USER_EMAIL")   # <- À AJOUTER dans Render

# ==========================================
# Vérification au démarrage
# ==========================================
print("========== DÉMARRAGE ==========")

for var in ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "USER_EMAIL"]:
    if os.getenv(var):
        print(f"✓ {var} chargé")
    else:
        print(f"❌ {var} MANQUANT")

print("===============================\n")


# ==========================================
# Obtenir un jeton Microsoft Graph
# ==========================================
def get_ms_token():

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    response = requests.post(url, data=data)

    print("\n===== AZURE TOKEN =====")
    print("Status :", response.status_code)

    if response.status_code != 200:
        print(response.text)
        response.raise_for_status()

    token = response.json()["access_token"]

    print("Token obtenu avec succès.\n")

    return token


# ==========================================
# WEBHOOK BREVO
# ==========================================
@app.route("/brevo-sync", methods=["POST"])
def brevo_sync():

    data = request.get_json(force=True)

    print("\n==============================")
    print("WEBHOOK BREVO REÇU")
    print("==============================")
    print(data)

    email = data.get("email")

    if not email:
        return jsonify({
            "error": "Aucune adresse email reçue."
        }), 400

    attributes = data.get("attributes", {})

    first_name = attributes.get("FIRSTNAME", "")
    last_name = attributes.get("LASTNAME", "")

    print("\nContact reçu :")
    print("Email :", email)
    print("Prénom :", first_name)
    print("Nom :", last_name)

    # ----------------------
    # Jeton Microsoft
    # ----------------------

    token = get_ms_token()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {

        "givenName": first_name,

        "surname": last_name,

        "emailAddresses": [
            {
                "address": email,
                "name": f"{first_name} {last_name}".strip()
            }
        ]
    }

    print("\nPayload envoyé à Microsoft :")
    print(payload)

    url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/contacts"

    response = requests.post(
        url,
        headers=headers,
        json=payload
    )

    print("\n===== MICROSOFT GRAPH =====")
    print("URL :", url)
    print("Status :", response.status_code)
    print("Réponse :")
    print(response.text)

    if response.status_code == 201:

        print("\n✓ Contact créé avec succès.\n")

        return jsonify({
            "success": True,
            "message": "Contact créé dans Outlook."
        }), 201

    return jsonify({
        "success": False,
        "status": response.status_code,
        "graph_error": response.text
    }), response.status_code


# ==========================================
# Accueil
# ==========================================
@app.route("/", methods=["GET"])
def home():
    return "Brevo → Microsoft 365 Sync (v2.0)"


# ==========================================
# Lancement
# ==========================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
