# outlook-control

## Example

```bash
from outlook.mac import Message

msg = Message(
    subject="Hello, World!",
    body="Test!",
    to_recipient=["chunj@mskcc.org"],
)
msg.show()
```

## Dependencies

### AppScript

Appscript is a high-level, user-friendly Apple event bridge that allows you to control scriptable Mac OS X applications from Python:

- http://appscript.sourceforge.net/
- https://pypi.org/project/appscript/
