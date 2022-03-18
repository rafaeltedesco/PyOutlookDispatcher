# Outlook Email Dispatcher

### A Simple Email Dispatcher based on top of win32Api

## Installation

```
pip install pyoutlookdispatcher
```

## Examples of Usage: 

### Send Email with Attachments

By default it adds your signature if you have one.

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

### Send Email without Signature

```
from pyoutlookdispatcher import Outlook, Mail

mail = Mail(
    Subject="Teste",
    To="example@example.com",
    HTMLBody="Teste",
    CC="example@example.com",
    Attachments=ATTACHMENTS,
    Signature=False
)

outlook = Outlook()
outlook.send(mail)
```

### Preview an Email

```
from pyoutlookdispatcher import Outlook, Mail

mail = Mail(
    Subject="Teste",
    To="example@example.com",
    HTMLBody="Teste",
    CC="example@example.com",
    Attachments=ATTACHMENTS,
)

outlook = Outlook()
outlook.preview(mail)
```

## Object Mail Params:
```
Subject: str
To: str
HTMLBody: str
CC: Optional[str] = None
Attachments: Optional[List[str]] = field(default_factory=list)
Signature: Optional[bool] = True
```

## Short Use Cases:

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


