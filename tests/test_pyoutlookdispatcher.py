from pyoutlookdispatcher import __version__
from pyoutlookdispatcher import Outlook, Mail

def test_version():
    assert __version__ == '0.1.0'


def test_mail_dispatch():
    mail = Mail(
        Subject="Subject here",
        To="example@example.com",
        HTMLBody="<h1>Teste email</h1>"
    )
    outlook = Outlook()
    could_send = outlook.send(mail)
    assert could_send == True