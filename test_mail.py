"""
Script de diagnostic e-mail — à exécuter sur le poste de travail.
Usage : python test_mail.py
"""
import smtplib, ssl, sys

HOST = "barnabe.oziolab.fr"
USER = "j.cosmao@actiajuris.fr"
PWD  = "tkq6b3zq"
TO   = "c.sevilla@actiajuris.fr"

def tester(port: int, mode: str):
    print(f"\n--- Test port {port} ({mode}) ---")
    try:
        ctx = ssl.create_default_context()
        if mode == "SSL":
            srv = smtplib.SMTP_SSL(HOST, port, context=ctx, timeout=10)
        else:
            srv = smtplib.SMTP(HOST, port, timeout=10)
            srv.ehlo()
            srv.starttls(context=ctx)
            srv.ehlo()

        srv.login(USER, PWD)
        print("  Connexion + login : OK")

        from email.mime.text import MIMEText
        msg = MIMEText("Test de notification — Registre Timbres Fiscaux", "plain", "utf-8")
        msg["Subject"] = "[TEST] Notification timbres"
        msg["From"]    = USER
        msg["To"]      = TO
        srv.sendmail(USER, [TO], msg.as_bytes())
        srv.quit()
        print(f"  E-mail envoyé à {TO} : OK")
        return True
    except Exception as e:
        print(f"  ÉCHEC : {type(e).__name__}: {e}")
        return False

if __name__ == "__main__":
    print(f"Serveur cible : {HOST}")
    ok = False
    ok = ok or tester(587, "STARTTLS")
    ok = ok or tester(465, "SSL")
    ok = ok or tester(25,  "STARTTLS")
    if not ok:
        print("\n⚠  Aucune combinaison port/mode n'a fonctionné.")
        print("   → Vérifiez que barnabe.oziolab.fr est joignable depuis ce poste.")
        print("   → Vérifiez les identifiants avec votre hébergeur mail.")
        sys.exit(1)
    else:
        print("\n✓ Un e-mail de test a été envoyé avec succès.")
