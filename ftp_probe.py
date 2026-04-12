# -*- coding: utf-8 -*-
import os
import ftplib


def main():
    # load .env if python-dotenv installed
    try:
        from dotenv import load_dotenv  # type: ignore

        p = (os.environ.get("AIPRINTWORD_DOTENV_PATH") or "").strip()
        if p and os.path.isfile(p):
            load_dotenv(p, override=True, encoding="utf-8-sig")
        else:
            d = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
            if os.path.isfile(d):
                load_dotenv(d, override=True, encoding="utf-8-sig")
    except Exception:
        pass
    host = os.environ.get("FTP_HOST", "10.26.1.221")
    port = int(os.environ.get("FTP_PORT", "2121"))
    user = os.environ.get("FTP_USER", "")
    pwd = os.environ.get("FTP_PASSWORD", "")
    print("Host", host, "Port", port, "User", user)

    print("\n== Plain FTP ==")
    try:
        ftp = ftplib.FTP()
        ftp.connect(host, port, timeout=10)
        print("welcome:", ftp.getwelcome())
        try:
            print("FEAT:", ftp.sendcmd("FEAT"))
        except Exception as e:
            print("FEAT err:", type(e).__name__, e)
        try:
            ftp.login(user=user, passwd=pwd)
            print("login ok")
            try:
                print("PWD:", ftp.pwd())
            except Exception as e:
                print("PWD err:", type(e).__name__, e)
            ftp.quit()
        except Exception as e:
            print("login err:", type(e).__name__, e)
            try:
                ftp.close()
            except Exception:
                pass
    except Exception as e:
        print("connect err:", type(e).__name__, e)

    print("\n== Explicit FTPS (FTP_TLS) ==")
    try:
        ftps = ftplib.FTP_TLS()
        ftps.connect(host, port, timeout=10)
        print("welcome:", ftps.getwelcome())
        try:
            print("FEAT:", ftps.sendcmd("FEAT"))
        except Exception as e:
            print("FEAT err:", type(e).__name__, e)
        try:
            ftps.auth()
            ftps.prot_p()
            ftps.login(user=user, passwd=pwd)
            print("login ok")
            try:
                print("PWD:", ftps.pwd())
            except Exception as e:
                print("PWD err:", type(e).__name__, e)
            ftps.quit()
        except Exception as e:
            print("auth/login err:", type(e).__name__, e)
            try:
                ftps.close()
            except Exception:
                pass
    except Exception as e:
        print("connect err:", type(e).__name__, e)


if __name__ == "__main__":
    main()

