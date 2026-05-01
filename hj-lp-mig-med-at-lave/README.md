# Outlook frokost add-in

Et lille Office.js add-in til Outlook-aftaler, hvor brugeren kan bestille frokost til en godkendt kantine eller levering til Boardroom på adressen Industrivej 8.

## Regler

- Add-in'et er lavet til Outlook-aftaler.
- Mødets lokation bruges kun som reference i bestillingsmailen.
- Bestillingssted vælges manuelt i formularen.
- Der kan kun bestilles frokost.
- Godkendte kantiner: `Kantinen på Industrivej`, `Kantinen på Savværksvej` og `Kantinen på Gammelgårdsvej`.
- Kantinerne viser kun `Frokost i kantinen`.
- Boardroom viser `Frokost i kantinen`, `Smørrebrød`, `Sandwich` og `Poke bowl`.
- `Smørrebrød`, `Sandwich` og `Poke bowl` kan kun vælges, når bestillingsstedet er `Boardroom`.
- Brugeren kan kun vælge én frokostmulighed pr. bestilling.
- Bestilling under 2 dage før viser en advarsel, men blokerer ikke.
- Bestillingen klargøres som en Outlook-mail med alle udfyldte oplysninger.

## Skift modtager

Ret modtageradressen i [src/taskpane/config.js](src/taskpane/config.js):

```js
recipientEmail: "kwl@jual.dk"
```

## Kør lokalt

```powershell
npm start
```

Serveren åbner som udgangspunkt på `http://localhost:3000`.

Outlook sideloading kræver normalt HTTPS. Hvis du lægger et lokalt certifikat i:

- `certs/localhost.pem`
- `certs/localhost-key.pem`

starter serveren automatisk som `https://localhost:3000`, som matcher [manifest.xml](manifest.xml).

## Sideload i Outlook

1. Start serveren med HTTPS.
2. Åbn Outlook på web.
3. Gå til `My add-ins` / `Mine tilføjelsesprogrammer`.
4. Vælg `Add a custom add-in` og derefter `Add from file`.
5. Vælg [manifest.xml](manifest.xml).

## Automatisk afsendelse

Denne version åbner en færdigudfyldt mail, som brugeren sender. Hvis mailen skal sendes helt automatisk uden brugerhandling, skal add-in'et have en backend med Microsoft Graph og de nødvendige administratorgodkendelser.
