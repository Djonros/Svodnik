import base64, io, sys
from PIL import Image
d = base64.b64decode(open(sys.argv[1], 'rb').read().strip())
im = Image.open(io.BytesIO(d))
print(im.size, file=sys.stderr)
f = float(sys.argv[2]) if len(sys.argv) > 2 else 0.4
im = im.convert('RGB').resize((int(im.width * f), int(im.height * f)))
b = io.BytesIO(); im.save(b, 'JPEG', quality=80)
print(base64.b64encode(b.getvalue()).decode())
