import os
import win32com.client as win32

from dataclasses import dataclass, field
from typing import List, Optional
    
@dataclass
class Mail:
    Subject: str
    To: str
    HTMLBody: str
    CC: Optional[str] = None
    Attachments: Optional[List[str]] = field(default_factory=list)
    Signature: Optional[bool] = True

class Outlook:
        
    def __init__(self):
        self._outlook = win32.dynamic.Dispatch('Outlook.Application')
        self._mail = None

    def send(self, mail: Mail) -> bool:
        self._create_new_mail(mail)
        try:
            self._mail.Send() 
            return True
        except Exception as e:
            print(e, 'erro send')
            return False

    def preview(self, mail: Mail):
        self._create_new_mail(mail)
        self._mail.Display(True)
       
    def _add_attachments(self, mail: Mail):
        for attach in mail.Attachments:
            if not os.path.isfile(attach):
                raise Exception(f'{attach} It\'s not a valid file')
            self._mail.Attachments.Add(Source=attach)
            
    def _add_copies(self, mail: Mail):
        self._mail.CC = mail.CC

    def _add_signature(self):
        self._mail.GetInspector.Activate()
        
    def _create_new_mail(self, mail: Mail):
        self._mail = self._outlook.CreateItem(0)
        self._mail.Subject = mail.Subject
        self._mail.To = mail.To
        if mail.CC:
            self._add_copies(mail)
        if mail.Signature:
            self._add_signature()
        self._mail.HTMLBody = mail.HTMLBody + self._mail.HTMLBody
        if mail.Attachments:
            self._add_attachments(mail)
        
        
