#!/bin/bash
PYTHON=/usr/local/bin/python

TENANT_ID="0bfce6b5-cf85-4dfd-98ac-813dd7557140"

# Récupérer et convertir toutes les clés en PEM
curl -s "https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys" | $PYTHON -c "
import sys, json, base64
from cryptography.hazmat.primitives.asymmetric.rsa import RSAPublicNumbers
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.backends import default_backend

def b64_to_int(s):
    s += '=' * (4 - len(s) % 4)
    return int.from_bytes(base64.urlsafe_b64decode(s), 'big')

data = json.load(sys.stdin)
for key in data['keys']:
    n = b64_to_int(key['n'])
    e = b64_to_int(key['e'])
    pub_key = RSAPublicNumbers(e, n).public_key(default_backend())
    pem = pub_key.public_bytes(
        serialization.Encoding.PEM,
        serialization.PublicFormat.SubjectPublicKeyInfo
    )
    print(f\"kid: {key['kid']}\")
    print(pem.decode())
    print('---')
"
