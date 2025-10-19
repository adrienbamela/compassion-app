import qrcode

# URLs de ton application déployée
events = {
    'presence': 'https://compassion-app.onrender.com/presence',
    'questions': 'https://compassion-app.onrender.com/questions'
}

for name, url in events.items():
    img = qrcode.make(url)         # Génère le QR code
    img.save(f'{name}.png')        # Sauvegarde sous "presence.png" et "questions.png"
    print('QR généré :', name + '.png ->', url)
