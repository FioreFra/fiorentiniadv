# GTM Audit Parser — Guida rapida
**Brandalize Hub** · strumento interno

---

## Cos'è

Uno script Python che legge il file JSON esportato da GTM e genera automaticamente un Excel con la mappatura completa del tracciamento: GA4, Meta Ads, Google Ads, variabili e checklist CMP.

---

## Requisiti

- Python 3.8 o superiore → [python.org/downloads](https://www.python.org/downloads/)
- Libreria openpyxl (si installa una volta sola)

Per verificare che Python sia installato, apri il terminale e digita:
```
python --version
```

---

## Installazione (una volta sola)

```
pip install openpyxl
```

---

## Come esportare il container GTM

1. Entra in **GTM → Amministrazione → Esporta container**
2. Seleziona il workspace (di solito *Default Workspace* o quello attivo)
3. Clicca **Esporta** → scarica il file `.json`
4. Metti il file JSON nella stessa cartella dello script

---

## Utilizzo

Apri il terminale, naviga nella cartella e lancia:

```
python gtm_audit_parser.py NomeFile.json "Nome Cliente"
```

**Esempio reale:**
```
python gtm_audit_parser.py GTM-W2L4QV2_workspace80.json "Don Gnocchi"
```

Il nome cliente viene usato nel titolo del file Excel e nel foglio Overview.
Se ometti il nome cliente, verrà usato "Cliente" come default.

**Output:** l'Excel viene salvato automaticamente nella stessa cartella del file JSON.

---

## Struttura dell'Excel generato

| Foglio | Contenuto |
|---|---|
| **Overview** | Riepilogo container: ID, data, conteggi per piattaforma |
| **GA4** | Tutti gli event tag GA4 con trigger, tipo, pagina, parametri, auto/custom |
| **Meta Ads** | Tutti gli eventi Pixel con trigger, parametri, Pixel ID, flag deduplicazione EventID |
| **Google Ads** | Tutte le conversioni con etichetta, Conversion ID, trigger, pagina, valore |
| **CMP** | Checklist da compilare manualmente sulla configurazione del consenso |
| **Variabili GTM** | Tutte le variabili con tipo, descrizione, valore/fonte e tag che le usano |

---

## Cosa compila lo script / cosa resta manuale

| | Automatico | Manuale |
|---|---|---|
| Nomi eventi e tag | ✅ | |
| Trigger e tipo trigger | ✅ | |
| Pagine/URL (da filtri trigger) | ✅ | |
| Parametri e dimensioni | ✅ | |
| Auto vs personalizzato (GA4) | ✅ | |
| Deduplicazione EventID (Meta) | ✅ | |
| Variabili: tipo, fonte, utilizzo | ✅ | |
| Colonna **Note** e **Stato** | | ✅ |
| Script diretti su sito (no GTM) | | ✅ |
| Configurazione CMP | | ✅ |
| Verifica live (DebugView, browser) | | ✅ |

---

## Workflow consigliato

1. **Esporta** il container GTM del cliente
2. **Lancia lo script** → hai l'80% dell'audit pre-compilato in 5 secondi
3. **Apri l'Excel** e aggiungi la colonna Stato (tracciato / da verificare / da rimuovere)
4. **Testa il sito** con DevTools / GA4 DebugView per verificare gli eventi live e integrare gli eventi non GTM
5. **Compila il foglio CMP** navigando il sito del cliente
6. **Consegna** al cliente

---

## Risoluzione problemi comuni

**`ModuleNotFoundError: No module named 'openpyxl'`**
→ Esegui `pip install openpyxl` e riprova.

**`python` non riconosciuto**
→ Prova con `python3` al posto di `python`.

**Il file JSON non viene trovato**
→ Assicurati che lo script e il JSON siano nella stessa cartella, oppure specifica il percorso completo:
```
python gtm_audit_parser.py "C:\Users\franc\Downloads\GTM-XXXXX.json" "Cliente"
```

**L'Excel viene sovrascritto se rilanci lo script lo stesso giorno**
→ Il nome file include la data (`_YYYYMMDD`). Se vuoi conservare versioni multiple, rinomina il file prima di rilanciarli.
