import getpass
import hashlib

senha = getpass.getpass("Digite a senha: ")
print(hashlib.sha256(senha.encode("utf-8")).hexdigest())
