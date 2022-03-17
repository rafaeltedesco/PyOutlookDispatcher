# Outlook Email Dispatcher

### A Simple Email Dispatcher based on top of win32Api

## Examples of Usage: 

### Send Email with Attachments

```
import os
from pyoutlookdispatcher import Outlook, Mail

FILES_TO_ATTACH_FOLDER = os.path.join(os.getcwd(), 'files_to_attach')
ATTACHMENTS = [os.path.join(FILES_TO_ATTACH_FOLDER, f) for f in os.listdir(FILES_TO_ATTACH_FOLDER)]

mail = Mail(
    Subject="Teste",
    To="example@example.com",
    HTMLBody="Teste",
    CC="example@example.com",
    Attachments=ATTACHMENTS
)

outlook = Outlook()
outlook.send(mail)
```

### Initialize Outlook

Instanciate an Object from Outlook Class

```
outlook = Outlook()
```

### Preview Mail:
```
outlook.preview(mail)
```

### Send Mail:
```
outlook.send(mail)
```


