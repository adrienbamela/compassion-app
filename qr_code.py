import qrcode


# Remplace par ton URL Render (après déploiement)
base = 'https://tonapp.onrender.com'


events = {
'presence': f'{base}/presence?event=Conf%C3%A9rence%20%C3%89tudiants',
'questions': f'{base}/questions?event=Conf%C3%A9rence%20%C3%89tudiants'
}


for name, url in events.items():
img = qrcode.make(url)
img.save(f'{name}.png')
print('QR généré :', name + '.png ->', url)
