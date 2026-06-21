"""Start ngrok tunnel for port 8501 and keep it alive."""
import os
import sys
import time
from pyngrok import ngrok
from dotenv import load_dotenv

load_dotenv()
token = os.getenv("NGROK_AUTHTOKEN", "").strip()
port = int(os.getenv("STREAMLIT_PORT", "8501"))

if not token:
    print("ERROR: NGROK_AUTHTOKEN missing", flush=True)
    sys.exit(1)

ngrok.set_auth_token(token)
# clean existing
for t in ngrok.get_tunnels():
    try:
        ngrok.disconnect(t.public_url)
    except Exception:
        pass

tunnel = ngrok.connect(port, "http")
url = tunnel.public_url
print(f"NGROK_URL={url}", flush=True)
print(f"LOCAL_URL=http://localhost:{port}", flush=True)
# sidecar file for app.py to read
with open("ngrok_url.txt", "w", encoding="utf-8") as f:
    f.write(url)
sys.stdout.flush()

# keep alive
try:
    while True:
        time.sleep(60)
except KeyboardInterrupt:
    ngrok.kill()
