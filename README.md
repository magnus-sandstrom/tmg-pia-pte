# TMG-PIA-PTE

## Översikt

TMG-PIA-PTE genererar en PIA-rapport (Pågående arbete) baserad på data från PrintVis kalkyl- och inköpstabeller.  
Rapporten är helt regelstyrd och exporteras till Excel.

PIA bryts ner i följande delar:

- Prepress PIA  
- Tryck PIA  
- Papper PIA  
- Efterbehandling PIA  
- Inköp PIA  
- Total PIA  

Detaljerad beräkningslogik finns i:  
👉 `docs/PIA-LOGIC.md`

---

## Grundprincip

Kostnader inkluderas i PIA beroende på:

- Kalkylenhet
- Uppdragsstatus
- Typ av kostnad (t.ex. Paper)
- Regler för att undvika dubbelräkning

---

## Viktiga regler

- Prepress inkluderas när jobbet passerat Prepress.
- Tryck inkluderas först när tryck är avslutat.
- Papper separeras från tryckkostnad.
- Efterbehandling inkluderas endast vid leverans.
- Inköp kopplas via `PVS Order No.`.
- Pappersartiklar exkluderas från Inköp PIA om de redan räknats i Papper PIA.

---

## Output

Rapporten genererar en Excel-fil med uppdelad PIA per kostnadstyp samt Total PIA.