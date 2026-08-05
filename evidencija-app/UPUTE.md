# 📱 Evidencija rada — upute za postavljanje

Prava aplikacija za tvoju firmu: prijava e-mailom, **samo tvoji zaposlenici** (nitko se ne može sam registrirati),
uloge admin/zaposlenik, sve na sigurnom serveru, radi na svim mobitelima kao instalirana aplikacija.

**Trošak: 0 € mjesečno** za malu firmu (Supabase i Vercel imaju besplatne pakete koji su ti više nego dovoljni).

---

## KORAK 1: Supabase (server i baza) — 10 minuta

1. Otvori **https://supabase.com** → *Start your project* → registriraj se (može Google račun)
2. **New project** → ime npr. `evidencija`, izmisli **Database password** (spremi ga negdje), regija: *Central EU (Frankfurt)* → Create
3. Pričekaj 1–2 min da se projekt kreira
4. Lijevi izbornik → **SQL Editor** → *New query* → zalijepi **CIJELI sadržaj datoteke `supabase/schema.sql`** → **Run**
   - Mora pisati "Success" — time su kreirane sve tablice i sigurnosna pravila

### Isključi samostalnu registraciju (ključno!)
5. Lijevi izbornik → **Authentication** → **Providers** → **Email**:
   - **Isključi** "Allow new users to sign up" (ugasi prekidač)
   - Isključi i "Confirm email" (da zaposlenici ne moraju potvrđivati mail)
   - Save

Time se **nitko izvana ne može registrirati** — prijaviti se mogu samo računi koje TI kreiraš.

### Kreiraj svoj (admin) račun
6. **Authentication** → **Users** → **Add user** → *Create new user*:
   - upiši svoj e-mail i lozinku → Create
7. Vrati se u **SQL Editor** i pokreni (zamijeni e-mail svojim):
   ```sql
   update profiles set role = 'admin', name = 'Tvoje Ime'
   where id = (select id from auth.users where email = 'tvoj@email.com');
   ```
   Sad si ti admin. 👑

### Uzmi ključeve
8. **Settings** (zupčanik) → **API**:
   - kopiraj **Project URL** i **anon public** ključ (trebat će u koraku 2)

---

## KORAK 2: Pokreni aplikaciju lokalno — 5 minuta

Treba ti [Node.js](https://nodejs.org) (LTS verzija) na računalu.

```bash
cd evidencija-app
cp .env.example .env
# otvori .env i upiši svoj Project URL i anon ključ iz koraka 1.8

npm install
npm run dev
```

Otvori adresu koju ispiše (obično http://localhost:5173) i prijavi se svojim e-mailom i lozinkom.

---

## KORAK 3: Objavi na internet (Vercel) — 5 minuta

1. Otvori **https://vercel.com** → registriraj se
2. Najlakše: instaliraj Vercel CLI i objavi iz mape projekta:
   ```bash
   npm install -g vercel
   vercel
   ```
   (odgovori Enter na sva pitanja)
3. U Vercel nadzornoj ploči → tvoj projekt → **Settings → Environment Variables** → dodaj:
   - `VITE_SUPABASE_URL` = tvoj Project URL
   - `VITE_SUPABASE_ANON_KEY` = tvoj anon ključ
4. **Deployments** → Redeploy

Dobit ćeš adresu tipa `https://evidencija.vercel.app` — to je tvoja aplikacija, dostupna sa svakog mobitela.

---

## KORAK 4: Dodaj zaposlenike

Za svakog zaposlenika:
1. Supabase → **Authentication** → **Users** → **Add user** → njegov e-mail + lozinka koju mu ti odrediš
2. Pošalji mu link aplikacije + e-mail i lozinku (npr. WhatsAppom)
3. On otvori link, prijavi se i u pregledniku odabere **"Dodaj na početni zaslon"** — dobiva ikonu kao prava aplikacija
4. (Po želji) U aplikaciji, panel **👥 Korisnici**, možeš nekome dati admin ulogu

**Zaposlenik vidi:** SAMO svoje unose — radnike, objekte, sate i isplate koje je sam kreirao.
Tuđe objekte, radnike i podatke uopće ne vidi, kao da ne postoje.
**Admin vidi SVE:** sve radnike i objekte svih zaposlenika, naplatu, dobit, dnevnik aktivnosti i koš — i samo admin briše.
Ovo NIJE samo sakriveno u aplikaciji — server (Row Level Security) odbija zaposlenika koji bi pokušao
pročitati tuđe podatke, naplatu ili dobit, čak i da zaobiđe aplikaciju.

**Važno zbog toga:** radnika treba unijeti onaj zaposlenik koji će s njim raditi.
**Objekti idu na dva načina:** zaposlenik ga unese sam, ILI ga ti kao admin uneseš i **dodijeliš mu ga**
(otvori objekt → admin kartica → "Vidljiv zaposlenicima" → kvačica kraj imena).
Dodijeljeni objekt zaposlenik vidi i upisuje sate na njemu — ali i dalje vidi samo vlastite unose.

---

## Kasnije: App Store i Google Play (kad poželiš)

Ista aplikacija se zapakira pomoću **Capacitor**-a:
```bash
npm install @capacitor/core @capacitor/cli
npx cap init "Evidencija rada" com.tvojafirma.evidencija
npm run build && npx cap add android && npx cap add ios
```
- **Google Play**: račun 25 USD jednokratno → Android Studio → build → objava (recenzija par dana)
- **App Store**: Apple Developer 99 USD/god → Xcode (treba Mac) → objava (recenzija 1–3 dana)

Ali iskreno — za internu firmu, web aplikacija s "Dodaj na početni zaslon" radi identično,
bez naknada i bez čekanja recenzija. Trgovine imaju smisla tek ako aplikaciju želiš nuditi drugima.

---

## Česta pitanja

**Zaposlenik je zaboravio lozinku?** Supabase → Authentication → Users → tri točkice kraj korisnika → Reset password (ili mu upiši novu).

**Zaposlenik je otišao iz firme?** Isto mjesto → Delete user (ili Ban) — više se ne može prijaviti. Podaci koje je upisao ostaju.

**Gdje su moji podaci?** U tvom Supabase projektu (EU server, Frankfurt). Samo ti imaš pristup nadzornoj ploči.

**Backup?** Supabase besplatni paket čuva dnevne sigurnosne kopije 7 dana. Dodatno, gumb "Prebaci u Excel" u aplikaciji ti je ručni backup obračuna kad god želiš.
