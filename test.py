from outlook_app import Message

msg = Message(
    subject="Hello, World!",
    body="Test!",
    to_recip=["chunj@mskcc.org"],
)
msg.show()
