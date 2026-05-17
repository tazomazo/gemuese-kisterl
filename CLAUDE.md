# Andreas' Gemüsekisterl – Claude Code Kontext

## Projektübersicht
Wöchentliche Gemüse-Bestellplattform für einen Wiener Gemüsehändler.
Kunden bestellen wöchentlich Gemüse online, ohne zu bezahlen.
Der Administrator verwaltet Produkte per Excel-Upload und sieht alle Bestellungen.

## Tech Stack
- **Frontend:** React 18 (Create React App), Inline-Styles (kein Tailwind trotz ursprünglicher Planung)
- **Datenbank:** Supabase (PostgreSQL), direkt aus dem Browser angesprochen
- **Deployment:** Vercel (Auto-Deploy bei Git Push auf `main`)
- **Excel-Import:** SheetJS (`xlsx`)

## Projektstruktur
```
gemuese-app/
├── public/
│   └── index.html
├── src/
│   ├── App.js              ← Gesamte Anwendung (eine Datei)
│   ├── index.js            ← React-Einstiegspunkt
│   └── supabaseClient.js   ← Supabase-Instanz
├── .env                    ← Supabase-Zugangsdaten (nicht in Git)
├── .env.example
├── .gitignore
├── package.json
├── supabase_setup.sql      ← Datenbank-Schema
└── README.md
```

## Supabase-Datenbank

**Projekt-URL:** `https://ncgxtvxcgryglyrdgsys.supabase.co`  
**Anon Key:** in `.env` als `REACT_APP_SUPABASE_ANON_KEY`

### Tabellen

#### `products`
| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | bigint (auto) | Primary Key |
| `name` | text | Produktname |
| `price` | numeric | Preis in Euro |
| `unit` | text | Einheit (Kilo, Stück, 500g, …) |
| `category` | text | Kategorie (z. B. „Gemüse", „Obst & Früchte") |
| `special` | boolean | Als Frühlingsneuheit/NEU hervorheben |
| `created_at` | timestamptz | Automatisch |

#### `users`
| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | bigint (auto) | Primary Key |
| `name` | text (unique) | Anzeigename des Benutzers |
| `password` | text | Optionales Klartext-Passwort (leer = kein Passwort) |
| `created_at` | timestamptz | Automatisch |

#### `orders`
| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | bigint (auto) | Primary Key |
| `user_id` | bigint (FK → users.id) | Zugehöriger Benutzer |
| `cart` | jsonb | `{ "produktId": menge, … }` – Mengen als Zahlen |
| `updated_at` | timestamptz | Letztes Speicherdatum |

### Supabase-Hilfsfunktionen in App.js
```
dbLoadProducts()         → alle Produkte laden
dbSeedProducts(arr)      → Initialprodukte einfügen
dbReplaceProducts(arr)   → alle Produkte ersetzen (Excel-Import)
dbLoadUsers()            → alle Benutzer laden (für Login-Liste)
dbCreateUser(name, pw)   → neuen Benutzer anlegen
dbLoadOrder(userId)      → Warenkorb eines Users laden
dbSaveOrder(userId, cart)→ Warenkorb eines Users speichern (upsert)
dbLoadAllOrders()        → alle Bestellungen mit username (für Admin)
```

## App-Struktur (Screens)

```
"landing"      → Startseite mit Bestellen/Admin-Button
"login-user"   → Benutzer-Login (Liste + Neu anlegen)
"login-admin"  → Admin-Login (Name + Passwort)
"shop"         → Produktliste + Warenkorb (für eingeloggte User)
"admin"        → Admin-Panel (Bestellungen, Produkte, Excel-Import)
```

Screen-Wechsel über `setScreen("...")` im Root-State.

## Wichtige Konventionen

- **Kein Tailwind** – alle Styles sind Inline-Style-Objekte in JavaScript
- **Grünes Farbschema:** Primär `#16a34a`, Dunkel `#14532d`, Hell `#f0fdf4`, Border `#bbf7d0`
- **Shared Style-Objekte:** `btnPrimary`, `btnSecondary`, `backBtn`, `inputStyle`, `qBtn` am Ende von App.js
- **Mengen in 0,5er-Schritten** (auch halbe Kilos möglich)
- **Admin-Passwort:** `admin123` (Konstante `ADMIN_PASSWORD` oben in App.js)
- **Wochenlabel:** `WEEK_LABEL` Konstante oben in App.js — wöchentlich anpassen
- **ESLint:** Vercel bricht bei Warnings ab → `// eslint-disable-line` verwenden wenn nötig
- **Keine `<form>`-Tags** verwenden (React-Kompatibilität)

## Produktkategorien (aktuell)
- Frühlingsneuheiten 🌱
- Kartoffeln & Wurzeln 🥔
- Obst & Früchte 🍎
- Blatt & Salat 🥬
- Gemüse 🥦
- Pilze 🍄
- Zwiebeln & Knoblauch 🧅
- Kräuter 🌿
- Nüsse & Extras 🌰
- Öle & Essige 🫙
- Säfte & Honig 🍯
- Eier 🥚

## Deployment

```bash
# Lokal entwickeln
npm install
npm start        # → http://localhost:3000

# Deployen
git add .
git commit -m "Beschreibung"
git push         # → Vercel deployed automatisch
```

**Vercel Umgebungsvariablen** (unter vercel.com → Project Settings → Environment Variables):
- `REACT_APP_SUPABASE_URL`
- `REACT_APP_SUPABASE_ANON_KEY`

## Häufige Aufgaben

**Woche wechseln:** `WEEK_LABEL` in App.js Zeile 4 anpassen, z. B. `"KW 17 · 2026"`

**Neues Produkt als Sonderangebot markieren:** `special: true` in der Datenbank oder in `INITIAL_PRODUCTS`

**Admin-Passwort ändern:** `ADMIN_PASSWORD` Konstante in App.js

**Neue Kategorie hinzufügen:** In `categoryIcons` Objekt eintragen (Icon + Name)

**Datenbank zurücksetzen:** `supabase_setup.sql` erneut im Supabase SQL Editor ausführen
