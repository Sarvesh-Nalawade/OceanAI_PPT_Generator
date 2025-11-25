import requests

API_KEY = "AIzaSyAmkFv8Kr0vMyDnDCOtVs-FmUecC4AzipE"

url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={API_KEY}"

payload = {
    "contents": [
        {"parts": [{"text": "Say hello in one short sentence."}]}
    ]
}

response = requests.post(url, json=payload)

print("\n✅ RESPONSE:\n")
print(response.text)
