#!/usr/bin/env bash
# Generate a self-signed HTTPS cert valid for localhost AND your LAN IP.
# Required so the PWA's camera works on a phone (browsers block getUserMedia on http://<lan-ip>).
#
# Usage:
#   ./scripts/make_cert.sh           # auto-detects LAN IP, makes cert in ./data/
#   ./scripts/make_cert.sh 192.168.1.42   # specify the IP yourself
#
# After running:
#   uvicorn backend.app:app --host 0.0.0.0 --port 8443 \
#       --ssl-keyfile data/cert.key --ssl-certfile data/cert.crt
# Then open https://<lan-ip>:8443/ on the phone and TAP-THROUGH the
# "not trusted" warning once (Self-signed certs always warn).

set -euo pipefail

cd "$(dirname "$0")/.."

LAN_IP="${1:-}"
if [[ -z "$LAN_IP" ]]; then
  # macOS / Linux fallback to find the primary LAN IP.
  LAN_IP=$(ipconfig getifaddr en0 2>/dev/null || \
           ipconfig getifaddr en1 2>/dev/null || \
           hostname -I 2>/dev/null | awk '{print $1}' || true)
fi
if [[ -z "$LAN_IP" ]]; then
  echo "Could not detect LAN IP. Pass it explicitly: $0 192.168.x.x"
  exit 1
fi

mkdir -p data
KEY="data/cert.key"
CRT="data/cert.crt"
CFG="$(mktemp)"

cat > "$CFG" <<EOF
[req]
default_bits        = 2048
prompt              = no
default_md          = sha256
distinguished_name  = dn
x509_extensions     = san

[dn]
C   = AU
ST  = QLD
L   = Toowoomba
O   = Pacific Seeds
CN  = ${LAN_IP}

[san]
subjectAltName = @alt_names
basicConstraints = critical, CA:FALSE
keyUsage = digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth

[alt_names]
DNS.1 = localhost
DNS.2 = nursery.local
IP.1  = 127.0.0.1
IP.2  = ${LAN_IP}
EOF

openssl req -x509 -newkey rsa:2048 -sha256 -days 365 -nodes \
  -keyout "$KEY" -out "$CRT" -config "$CFG"

rm -f "$CFG"

echo
echo "✅ Cert generated:"
echo "   key: $KEY"
echo "   crt: $CRT"
echo
echo "Now start the server with TLS:"
echo
echo "  .venv/bin/uvicorn backend.app:app --host 0.0.0.0 --port 8443 \\"
echo "      --ssl-keyfile $KEY --ssl-certfile $CRT"
echo
echo "Then open on your phone:"
echo "  https://${LAN_IP}:8443/"
echo
echo "Your phone will warn 'not trusted'. Tap Advanced → Proceed once;"
echo "the camera will work after that."
