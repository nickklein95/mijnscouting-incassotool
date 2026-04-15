# SOL-incasso

Statische browser-app om een Excel-export met incassoregels om te zetten naar een SEPA `pain.008.001.02` XML-bestand.

Live versie:

- [mijnscouting-incassotool.pages.dev](https://mijnscouting-incassotool.pages.dev/)

## Bestanden

- [index.html](/Users/nickklein/Documents/SOL-incasso/index.html): pagina-structuur
- [styles.css](/Users/nickklein/Documents/SOL-incasso/styles.css): styling
- [app.js](/Users/nickklein/Documents/SOL-incasso/app.js): parsing, validatie en XML-generatie
- [bic-list.js](/Users/nickklein/Documents/SOL-incasso/bic-list.js): lokale BIC-lookup voor Nederlandse bankcodes

## Gebruik

1. Open `index.html` in een browser of deploy de map als statische site.
2. Upload het Excel-bestand met incassoregels.
3. Controleer de preview en meldingen.
4. Klik op `XML genereren`.
5. Klik op `XML downloaden`.

## Cloudflare Pages

Deze app kan direct op Cloudflare Pages worden gedeployed.

- Build command: geen
- Output directory: project root

Upload of publiceer deze bestanden samen:

- `index.html`
- `styles.css`
- `app.js`
- `bic-list.js`

## Afhankelijkheden

De Excel-parser wordt geladen vanaf een CDN:

- `https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js`

Er is geen backend nodig. Excel-bestanden worden in de browser verwerkt.

## Opmerking

Omdat `xlsx` vanaf een CDN wordt geladen, moet de browser internettoegang hebben om die dependency op te halen.
