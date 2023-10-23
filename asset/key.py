from cryptography.fernet import Fernet

def save_key():
    key = Fernet.generate_key()
    with open("secret.key", "wb") as key_file:
        key_file.write(key)

save_key()
