from appscript import app, k
from mactypes import Alias
from pathlib import Path


class Outlook(object):
    def __init__(self):
        self.client = app("Microsoft Outlook")


class Message(object):
    def __init__(
        self,
        parent=None,
        subject="",
        body="",
        to_recipient=[],
        cc_recipient=[],
        show_=True,
    ):

        if parent is None:
            parent = Outlook()
        client = parent.client

        self.msg = client.make(
            new=k.outgoing_message,
            with_properties={k.subject: subject, k.content: body},
        )

        self.add_recipients(emails=to_recipient, type_="to")
        self.add_recipients(emails=cc_recipient, type_="cc")

        if show_:
            self.show()

    def show(self):
        self.msg.open()
        self.msg.activate()

    def add_attachment(self, path_file):
        # path_file is a Path() obj, could also pass string

        path_file = Alias(
            str(path_file)
        )  # convert string/path obj to POSIX/mactypes path

        attach = self.msg.make(new=k.attachment, with_properties={k.file: path_file})

        return attach

    def add_recipients(self, emails, type_="to"):
        if not isinstance(emails, list):
            emails = [emails]
        for email in emails:
            self.add_recipient(email=email, type_=type_)

    def add_recipient(self, email, type_="to"):
        msg = self.msg

        if type_ == "to":
            recipient = k.to_recipient
        elif type_ == "cc":
            recipient = k.cc_recipient

        msg.make(new=recipient, with_properties={k.email_address: {k.address: email}})

