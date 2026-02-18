# Versionshantering i app.json -- Steg för steg

Den här guiden beskriver hur du hanterar versionsnummer i ett Business
Central AL-projekt tillsammans med Git.

------------------------------------------------------------------------

## 🔹 När ska du ändra versionen?

Du ska uppdatera versionen i `app.json` varje gång du gör en ändring som
ska:

-   Publiceras till sandbox
-   Delas med testare
-   Levereras till produktion
-   Taggas i Git

------------------------------------------------------------------------

## 🔹 Steg 1 -- Uppdatera version i app.json

Öppna `app.json` och ändra:

``` json
"version": "1.0.0.0"
```

Exempel på uppdateringar:

  Typ av ändring   Ny version
  ---------------- ------------
  Mindre fix       1.0.1.0
  Ny funktion      1.1.0.0
  Större ändring   2.0.0.0

### Rekommenderad struktur

-   **Major.Minor.Patch.Build**
-   1.0.0.0 → första stabila version
-   1.0.1.0 → mindre förbättring
-   1.1.0.0 → ny funktion
-   2.0.0.0 → större förändring

Spara filen.

------------------------------------------------------------------------

## 🔹 Steg 2 -- Commit ändringen

I terminalen:

``` bash
git add .
git commit -m "PIA v1.0.1:
- Beskriv kort vad som ändrats
- T.ex. förbättrad inköpslogik
- Justerad statusfiltrering"
```

------------------------------------------------------------------------

## 🔹 Steg 3 -- Pusha till GitHub

``` bash
git push origin main
```

------------------------------------------------------------------------

## 🔹 Steg 4 -- Skapa Git-tag (checkpoint)

Detta gör att du kan gå tillbaka till exakt denna version senare.

``` bash
git tag v1.0.1
git push origin v1.0.1
```

Nu finns versionen sparad både i kod och som Git-tag.

------------------------------------------------------------------------

## 🔹 Steg 5 -- Publicera till Business Central

Efter versionsändring måste appen publiceras igen:

1.  Kör **AL: Publish**
2.  Kontrollera att ny version visas i Extension Management

------------------------------------------------------------------------

## 🔹 Bra arbetsflöde framåt

1.  Gör ändring
2.  Uppdatera version i app.json
3.  Publish till sandbox
4.  Testa
5.  Commit + Tagga
6.  Skicka för granskning

------------------------------------------------------------------------

## 🔹 Varför detta är viktigt

-   Ekonomi kan hänvisa till specifik version
-   Du kan alltid backa till en tidigare tagg
-   Det blir tydlig spårbarhet
-   Produktionsuppdateringar blir kontrollerade

------------------------------------------------------------------------

Skapad för PIA-projektet -- Taberg Media Group.
