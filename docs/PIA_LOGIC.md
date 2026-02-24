```markdown
# PIA-LOGIC.md

# PTE: TMG-PIA-PTE

## Översikt

Appen genererar en **PIA-rapport (Pågående arbete)** och bryter ner kostnader i följande delar:

- Prepress PIA  
- Tryck PIA  
- Papper PIA  
- Efterbehandling PIA  
- Inköp PIA  
- Total PIA  

Rapporten:

- Hämtar data direkt från **PrintVis kalkyl- och inköpstabeller**
- Är helt **regelstyrd**
- Exporteras till en **Excel-fil**

---

# Prepress PIA

## Ingående kalkylenheter

- 210-PE10  
- 210-PE20  

## Statuskrav

Kostnader inkluderas för uppdrag med status:

- PRODUKTION  
- EFTERBEHANDLING  
- LEVERANS  

## Regel

Så snart uppdraget har passerat Prepress och gått vidare i flödet inkluderas Prepress-kostnaden i PIA.

---

# Tryck PIA

## Ingående kalkylenheter

- 340-PE10  
- 350-PE10  
- 360-PE10  
- 380-PE10  
- 510-PE10  
- 510-PE20  
- 550-PE10  
- 560-PE10  
- 570-PE10  
- 580-PE10  
- 770-PE20  

## Statuskrav

Kostnader inkluderas för uppdrag med status:

- EFTERBEHANDLING  
- LEVERANS  

## Regel

Tryck måste vara avslutat innan kostnaden inkluderas i PIA.

## Viktigt

Tryck PIA inkluderar **inte** papper.  
Papperskostnaden bryts ut separat.

---

# Papper PIA

## Scope

Samma kalkylenheter som Tryck PIA:

- 340-PE10  
- 350-PE10  
- 360-PE10  
- 380-PE10  
- 510-PE10  
- 510-PE20  
- 550-PE10  
- 560-PE10  
- 570-PE10  
- 580-PE10  
- 770-PE20  

## Urvalskriterier

Endast kalkylrader där:

- `Item Type = Paper`
- Belopp hämtas från `Cost Amount`

## Statuskrav

- EFTERBEHANDLING  
- LEVERANS  

## Syfte

Separera papperskostnaden från övrig tryckkostnad:

- Tryck PIA = maskin-/produktionskostnad  
- Papper PIA = materialkostnad  

## Exempel

| Post                     | Belopp |
|--------------------------|--------|
| Ursprunglig tryckkostnad | 3000   |
| Papper                   | 1000   |
| **Tryck PIA**            | 2000   |
| **Papper PIA**           | 1000   |

---

# Efterbehandling PIA

## Ingående kalkylenheter

- 610-PE10  
- 610-PE20  
- 610-PE30  
- 620-PE10  
- 630-PE10  
- 640-PE10  
- 645-PE10  
- 650-PE10  
- 655-PE10  
- 660-PE10  
- 670-PE10  
- 680-PE10  
- 690-PE10  
- 750-PE10  
- 750-PE15  
- 750-PE20  
- 750-PE25  
- 760-PE10  
- 780-PE10  
- 790-PE10  

## Statuskrav

Endast för uppdrag med status:

- LEVERANS  

## Regel

Efterbehandling måste vara slutförd innan kostnaden inkluderas i PIA.

---

# Inköp PIA

## Datakälla

- Tabell: `Purch. Inv. Line`
- Fält: `Amount`
- Koppling via: `PVS Order No.`

## Statuskrav

Gäller Case med status:

- ORDER  
- PROD.FÖRB  
- PREPRESS  
- KORREKTUR  
- PRODUKTION  
- EFTERBEHANDLING  
- LEVERANS  
- EFTERKALKYL  

## Regel – Undvik dubbelräkning

Om en inköpsrad har ett artikelnummer som:

- Matchar ett artikelnummer som redan räknats som Papper PIA  

→ Då exkluderas den inköpsraden.

### Syfte

Förhindra att papperskostnad räknas två gånger:

- En gång i kalkylen  
- En gång i inköp  

---

# Total PIA

Total PIA beräknas som summan av:

- Prepress PIA  
- Tryck PIA  
- Papper PIA  
- Efterbehandling PIA  
- Inköp PIA  
```
