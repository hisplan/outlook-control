# outlook-control

```python
to_recipient = ["chunj@mskcc.org"]

msg = Message(subject=subject, body=body, to_recipient=to_recipient)

path_file = Path("/Users/chunj/hello.txt")
msg.add_attachment(path_file)

msg.show()
```
