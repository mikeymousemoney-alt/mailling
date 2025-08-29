import ctypes
from ctypes import wintypes
import binascii
import datetime
import json
import base64
import requests
import hashlib
import time

crypt32 = ctypes.WinDLL('crypt32.dll')
ncrypt = ctypes.WinDLL("ncrypt.dll")

# Konstanten
CERT_STORE_PROV_SYSTEM = 10
CERT_SYSTEM_STORE_CURRENT_USER = 0x00010000
CERT_STORE_READONLY_FLAG = 0x00008000
CERT_FIND_HASH = 0x10000
X509_ASN_ENCODING = 0x00000001

CRYPT_ACQUIRE_ONLY_NCRYPT_KEY_FLAG = 0x00040000
NCRYPT_SILENT_FLAG = 0x00000040

CALG_SHA_256 = 0x0000800c


# Typen
PCCERT_CONTEXT = ctypes.c_void_p
HCRYPTPROV_OR_NCRYPT_KEY_HANDLE = wintypes.HANDLE
HCRYPTPROV = wintypes.HANDLE
HCRYPTHASH = wintypes.HANDLE
DWORD = wintypes.DWORD
BOOL = wintypes.BOOL
LPBYTE = ctypes.POINTER(ctypes.c_ubyte)
LPVOID = ctypes.c_void_p

# Struktur für Thumbprint-Blob
class CRYPT_HASH_BLOB(ctypes.Structure):
    _fields_ = [("cbData", DWORD), ("pbData", LPBYTE)]

# Funktionen

def open_cert_store(store_name="MY"):
    store = crypt32.CertOpenStore(b"System", 0, None, CERT_SYSTEM_STORE_CURRENT_USER | CERT_STORE_READONLY_FLAG, store_name)
    if not store:
        raise ctypes.WinError()
    return store

def close_cert_store(store):
    if not crypt32.CertCloseStore(store, 0):
        raise ctypes.WinError()

def find_cert_by_thumbprint(store, thumbprint_hex):
    thumbprint_bytes = binascii.unhexlify(thumbprint_hex.replace(" ", ""))
    cbData = len(thumbprint_bytes)
    data_array = (ctypes.c_ubyte * cbData)(*thumbprint_bytes)
    hash_blob = CRYPT_HASH_BLOB(cbData, data_array)
    pCertContext = crypt32.CertFindCertificateInStore(store, X509_ASN_ENCODING, 0, CERT_FIND_HASH, ctypes.byref(hash_blob), None)
    if not pCertContext:
        return None
    return pCertContext

NCRYPT_KEY_HANDLE = wintypes.HANDLE
CRYPT_ACQUIRE_SILENT_FLAG = 0x00000040
CRYPT_ACQUIRE_ONLY_NCRYPT_KEY_FLAG = 0x00040000

def acquire_cng_key_handle(cert_ctx: PCCERT_CONTEXT) -> NCRYPT_KEY_HANDLE:
    key_handle = NCRYPT_KEY_HANDLE()
    key_spec = wintypes.DWORD()
    caller_free = wintypes.BOOL()

    res = crypt32.CryptAcquireCertificatePrivateKey(
        cert_ctx,
        CRYPT_ACQUIRE_SILENT_FLAG | CRYPT_ACQUIRE_ONLY_NCRYPT_KEY_FLAG,
        None,
        ctypes.byref(key_handle),
        ctypes.byref(key_spec),
        ctypes.byref(caller_free)
    )

    if not res:
        raise ctypes.WinError()

    return key_handle

def get_x5t_thumbprint(cert_ctx):
    class CERT_CONTEXT(ctypes.Structure):
        _fields_ = [
            ("dwCertEncodingType", wintypes.DWORD),
            ("pbCertEncoded", ctypes.POINTER(ctypes.c_ubyte)),
            ("cbCertEncoded", wintypes.DWORD),
            ("pCertInfo", wintypes.LPVOID),
            ("hCertStore", wintypes.HANDLE)
        ]

    cert = ctypes.cast(cert_ctx, ctypes.POINTER(CERT_CONTEXT)).contents
    cert_bytes = ctypes.string_at(cert.pbCertEncoded, cert.cbCertEncoded)
    sha1 = hashlib.sha1(cert_bytes).digest()
    return base64.urlsafe_b64encode(sha1).rstrip(b'=').decode()

def ncrypt_sign_hash(key_handle, jwt_payload: bytes) -> bytes:
    # Hash vorbereiten (SHA-256 über JWT header.payload)
    hash_val = hashlib.sha256(jwt_payload.encode("ascii")).digest()
    hash_ptr = ctypes.create_string_buffer(hash_val)
    hash_len = len(hash_val)

    sig_len = wintypes.DWORD()
    #Padding info für SHA256
    class BCRYPT_PKCS1_PADDING_INFO(ctypes.Structure):
         _fields_ = [("pszAlgId", wintypes.LPCWSTR)]

    padding_info = BCRYPT_PKCS1_PADDING_INFO("SHA256")
    BCRYPT_PAD_PKCS1 = 0x00000002
    flags = 0  # PKCS1 padding

    # Schritt 1: Länge der Signatur ermitteln
    status = ncrypt.NCryptSignHash(
        key_handle,
        None,
        hash_ptr,
        hash_len,
        None,
        0,
        ctypes.byref(sig_len),
        flags
    )
    if status != 0:
        raise ctypes.WinError()

    # Schritt 2: Signatur erzeugen
    sig_buf = ctypes.create_string_buffer(sig_len.value)
    status = ncrypt.NCryptSignHash(
        key_handle,
        ctypes.byref(padding_info),
        hash_ptr,
        hash_len,
        sig_buf,
        sig_len,
        ctypes.byref(sig_len),
        BCRYPT_PAD_PKCS1
    )
    if status != 0:
        raise ctypes.WinError()

    return sig_buf.raw[:sig_len.value]

def base64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")

def create_jwt_header_payload(client_id, tenant_id, x5t):
    header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
    now = now = int(time.time())
    payload = {
        "aud": f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        "iss": client_id,
        "sub": client_id,
        "jti": "unique-jti-1234567890",
        "nbf": now,
        "exp": now + 1800
    }

    # header_b64 = base64url_encode(json.dumps(header).encode("utf-8"))
    # payload_b64 = base64url_encode(json.dumps(payload).encode("utf-8"))
    header_b64 = base64url_encode(json.dumps(header).encode())
    payload_b64 = base64url_encode(json.dumps(payload).encode())
    return header_b64, payload_b64

########################################################################################################################
#  Get Folder ID
########################################################################################################################
def get_folder_id(folder_name, mailbox):
    #hier müsste eine Überprüfung eingebaut werden, ob der Access Token noch gültig ist ...
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/Inbox/childFolders"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    folders = response.json().get("value", [])
    for folder in folders:
        if folder["displayName"] == folder_name:
            print("Folder ID:", folder["id"])
            return folder["id"]
    return None

########################################################################################################################
#  Move message
########################################################################################################################

def move_message(message_id, mailbox, destination_folder_id):
    print("Moving mail now:")
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}/move"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    data = {
        "destinationId": destination_folder_id
    }
    response = requests.post(url, headers=headers, json=data)
    response.raise_for_status()
    return response.json()

########################################################################################################################
#  Mark Message as Read
########################################################################################################################

def mark_message(message_id, mailbox, is_read):
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    if is_read :
        data = {
            "isRead": True
        }
    else:
        data = {
            "isRead": False
        }
    response = requests.patch(url, headers=headers, json=data)
    
    if response.status_code == 200:
        print(f"Nachricht {message_id} wurde als gelesen markiert.")
    else:
        print(f"Fehler beim Aktualisieren der Nachricht {message_id}: {response.status_code}")
        print(response.text)


########################################################################################################################
#  Send Mail
########################################################################################################################
def send_mail(subject, body_html, recipient_emails, mailbox):
    """
    Sendet eine E-Mail über Microsoft Graph API.

    :param subject: Betreff der E-Mail
    :param body_html: HTML-Inhalt der E-Mail
    :param recipient_emails: Liste von E-Mail-Adressen (Empfänger)
    :param sender: Benutzer-E-Mail (z.B. für Shared Mailbox)
    """
    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    message = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body_html
            },
            "toRecipients": [
                {"emailAddress": {"address": email}} for email in recipient_emails
            ]
        },
        "saveToSentItems": "true"
    }

    response = requests.post(url, headers=headers, json=message)
    
    if response.status_code == 202:
        print("E-Mail erfolgreich gesendet.")
    else:
        print(f"Fehler beim Senden: {response.status_code}")
        print(response.text)

########################################################################################################################
#  Get Attachments
########################################################################################################################

def get_email_attachments(mailbox, message_id):
    """
    Fetches and downloads attachments from a specified email message in a Microsoft Graph API mailbox.

    Args:
        mailbox: Email address or user ID of the mailbox.
        message_id (str): ID of the email message.
        download_folder (str): Folder to save attachments.

    Returns:
        List of saved file paths.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json'
    }

    url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}/attachments"
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        # Nicht erfolgreich, aber kein harter Abbruch
        print(f"Info: Keine Attachments (Status {response.status_code})")
        return[]
        #raise Exception(f"Failed to get attachments: {response.status_code} - {response.text}")

    attachments = response.json().get('value', [])
    results = []

    for attachment in attachments:
        if attachment.get('@odata.type') == '#microsoft.graph.fileAttachment':
            filename = attachment['name']
            content_bytes = base64.b64decode(attachment['contentBytes'])
            results.append({
                'filename': filename,
                'content': content_bytes
            })

    return results

########################################################################################################################
#  Main
########################################################################################################################

def main(start_date, thumbprint, client_id, tenant_id, mailbox):

    store = open_cert_store("MY")
    try:
        cert_ctx = find_cert_by_thumbprint(store, thumbprint)
        if not cert_ctx:
            raise Exception("Zertifikat nicht gefunden!")
        print("Zertifikat gefunden.")

        hKey = acquire_cng_key_handle(cert_ctx)
        print(f"Privater Schlüssel Handle: {hKey}")
        #acquire x5t for JWT header
        x5t = get_x5t_thumbprint(cert_ctx)
        header_b64, payload_b64 = create_jwt_header_payload(client_id, tenant_id, x5t)

        # signature for header and payload
        jwt_header_payload = f"{header_b64}.{payload_b64}"
        signature = ncrypt_sign_hash(hKey, jwt_header_payload)
        signature_b64 = base64url_encode(signature)
        jwt_token = f"{header_b64}.{payload_b64}.{signature_b64}"

        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": client_id,
            "scope": "https://graph.microsoft.com/.default",
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": jwt_token,
            "grant_type": "client_credentials"
        }
        

        resp = requests.post(token_url, data=data)
        if not resp.ok:
            print("Fehler beim Token-Request:")
            print("Status:", resp.status_code)
            print("Antwort:", resp.text)
            resp.raise_for_status()

        global access_token
        access_token = resp.json()["access_token"]

        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/Inbox/messages"
        headers = {"Authorization": f"Bearer {access_token}"}
        # Format für Microsoft Graph filter: yyyy-MM-ddTHH:mm:ssZ
        start_date_str = start_date.strftime("%Y-%m-%dT%H:%M:%SZ")

        params = {
            "$filter": f"receivedDateTime ge {start_date_str} and isRead eq false",
            "$top": 100
        }
        result = requests.get(url, headers=headers, params=params)
        result.raise_for_status()

        messages = result.json().get("value", [])
        return messages


    finally:
        if cert_ctx:
            crypt32.CertFreeCertificateContext(cert_ctx)
        close_cert_store(store)