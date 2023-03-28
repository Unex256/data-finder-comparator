import base64

with open('images/ui/cogwheel25x25.png', 'rb') as f:
    image_data = f.read()
    cogwheel_image_data = base64.b64encode(image_data).decode('utf-8')