# create_president.py
from app import create_president_account
import getpass

if __name__ == "__main__":
    username = input("Username: ").strip()
    fullname = input("Nom complet (optionnel): ").strip()
    departements = input("Départements (séparés par virgule): ").strip()
    email = input("Email (optionnel): ").strip()
    pwd = getpass.getpass("Mot de passe: ")
    pwd2 = getpass.getpass("Confirmer mot de passe: ")
    if pwd != pwd2:
        print("Les mots de passe ne correspondent pas.")
        exit(1)
    try:
        create_president_account(username, pwd, fullname=fullname, departements=departements, email=email)
        print("Compte président créé avec succès.")
    except Exception as e:
        print("Erreur :", e)
