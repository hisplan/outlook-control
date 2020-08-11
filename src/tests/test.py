from outlook.mac import Message

msg = Message(
    subject="Hello, World!",
    body="Test!",
    to_recipient=["chunj@mskcc.org"],
)
msg.show()
