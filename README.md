# 🌿 Andreas' Gemüsekisterl – à la carte

Wöchentliche Gemüse-Bestellplattform mit React + Supabase.

---

## Schritt 1 – Supabase Datenbank einrichten

1. Gehe zu [supabase.com](https://supabase.com) und öffne dein Projekt
2. Klicke links auf **SQL Editor**
3. Kopiere den gesamten Inhalt von `supabase_setup.sql` und füge ihn ein
4. Klicke auf **Run** – die Tabellen `products` und `orders` werden angelegt

---

## Schritt 2 – Projekt auf GitHub hochladen

1. Gehe zu [github.com](https://github.com) → **New repository**
2. Name z. B. `gemuese-kisterl`, Sichtbarkeit: **Private**
3. Öffne ein Terminal in diesem Projektordner und führe aus:

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/DEIN-USERNAME/gemuese-kisterl.git
git push -u origin main
```

> **Wichtig:** Die `.env`-Datei wird durch `.gitignore` automatisch **nicht** hochgeladen.
> Die Zugangsdaten kommen in Schritt 3 als Umgebungsvariablen zu Vercel.

---

## Schritt 3 – Auf Vercel deployen

1. Gehe zu [vercel.com](https://vercel.com) und melde dich mit GitHub an
2. Klicke auf **Add New Project** → wähle dein Repository aus
3. Vercel erkennt React automatisch – keine Build-Einstellungen nötig
4. Vor dem Deploy: Klicke auf **Environment Variables** und füge hinzu:

| Name | Wert |
|---|---|
| `REACT_APP_SUPABASE_URL` | `https://ncgxtvxcgryglyrdgsys.supabase.co` |
| `REACT_APP_SUPABASE_ANON_KEY` | `sb_publishable_eiTk-TNMWY5p_GrSgr5MXw_zFPPQnW1` |

5. Klicke auf **Deploy** – nach ca. 1–2 Minuten ist die App live

Die App ist dann erreichbar unter einer URL wie:
`https://gemuese-kisterl.vercel.app`

---

## Lokale Entwicklung

```bash
npm install
npm start
```

Die App läuft dann auf [http://localhost:3000](http://localhost:3000).

---

## Benutzung

### Kunden
- Einfach Namen eingeben, kein Passwort erforderlich
- Produkte in den Warenkorb legen (auch halbe Einheiten möglich)
- Bestellung speichern – wird direkt in Supabase gespeichert
- Bestellung jederzeit erneut aufrufen und anpassen

### Administrator
- Name + Passwort eingeben (Standard: `admin123`)
- Alle Bestellungen einsehen und aktualisieren
- Neue Excel-Datei hochladen → ersetzt Produktliste in der Datenbank

---

## Passwort ändern

In `src/App.js` Zeile 1:
```js
const ADMIN_PASSWORD = "admin123";
```
Ersetze `admin123` durch ein sicheres Passwort, dann neu deployen.

---

## Projektstruktur

```
gemuese-app/
├── public/
│   └── index.html
├── src/
│   ├── App.js            ← Gesamte Anwendung
│   ├── index.js          ← React-Einstiegspunkt
│   └── supabaseClient.js ← Datenbankverbindung
├── .env                  ← Zugangsdaten (nicht in Git!)
├── .env.example          ← Vorlage für andere Entwickler
├── .gitignore
├── package.json
├── supabase_setup.sql    ← Datenbank-Setup
└── README.md
```

---

## Datenbank-Tabellen

### `products`
| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | bigint | Automatische ID |
| `name` | text | Produktname |
| `price` | numeric | Preis in Euro |
| `unit` | text | Einheit (Kilo, Stück, …) |
| `category` | text | Kategorie |
| `special` | boolean | Als Neuheit hervorheben |

### `orders`
| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | bigint | Automatische ID |
| `user_name` | text | Name des Bestellers (eindeutig) |
| `cart` | jsonb | `{ "produktId": menge, … }` |
| `updated_at` | timestamptz | Letztes Speicherdatum |
