import requests
import json

# ==========================================
# CDISC Library API Key
# ==========================================
API_KEY = "c09319188b59468dafe2c7d62d374d7e"

# ==========================================
# API Endpoint
# ==========================================
url = "https://library.cdisc.org/api/mdr/ct/packages"

# ==========================================
# Headers
# ==========================================
headers = {
    "api-key": API_KEY,
    "Accept": "application/json"
}

# ==========================================
# Request
# ==========================================
response = requests.get(url, headers=headers)

# ==========================================
# Output
# ==========================================
print("Status Code:", response.status_code)

if response.status_code == 200:
    data = response.json()

    print("\nSUCCESS - Connected to CDISC Library API\n")

    print(json.dumps(data, indent=2))

else:
    print("\nFAILED")
    print(response.text)