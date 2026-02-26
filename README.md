# DocGenerator
FPAnalysis Documentation Generator from Excel to Docx



# 📄 Doc Generator Pro

**Doc Generator** è un tool automatico sviluppato in Python per trasformare i dati estratti da file Excel (modello Conteggio Funzioni) in documentazione Word formattata.

---

## 🚀 Guida all'avvio rapido

1. **Scarica il pacchetto**: Scarica il file `.zip` dall'ultima [Release] e estrai il contenuto in una cartella.
2. **Avvia l'applicazione**: Fai doppio clic su `Doc_Generator.exe`.
3. **Sicurezza di Windows**: Se appare l'avviso "Windows ha protetto il PC", clicca su *Ulteriori informazioni* e poi su **Esegui comunque**.

---

## 🛠️ Come si usa

1. **Seleziona il File**: Clicca su "Seleziona File Excel" e scegli il tuo file `.xlsx`.
2. **Configura il Foglio**: Assicurati che il nome del foglio nel campo di testo corrisponda a quello presente nell'Excel (es. `ConteggioFunzioni`).
3. **Genera**: Clicca sul tasto verde **GENERA DOCUMENTO**.
4. **Risultato**: Il file Word verrà creato istantaneamente nella stessa cartella dove si trova il file Excel di origine.

---

## ⚙️ Personalizzazione (config.json)

Il programma è dinamico. Se le colonne del tuo file Excel cambiano, non serve modificare il codice. Puoi agire in due modi:
1. **Dall'App**: Vai nella scheda **Impostazioni**, modifica gli indici delle colonne o le frasi e clicca su "Salva".
2. **Dal File**: Apri il file `config.json` con il Blocco Note e modifica i valori.



### Mappatura Colonne predefinita:
* **Tipo (B)**: Colonna 1
* **Titolo (D)**: Colonna 3
* **Numero (E)**: Colonna 4
* **FTR/RET (P)**: Colonna 15
* **DET (Q)**: Colonna 16

---

## 🔄 Aggiornamenti

Il programma controlla automaticamente la presenza di nuove versioni su GitHub all'avvio. Se viene rilevato un aggiornamento, ti verrà chiesto se desideri scaricare l'ultima versione disponibile.

---

## ⚠️ Note Tecniche
* Assicurati che il file Excel sia **chiuso** prima di avviare la generazione.
* Non rinominare o eliminare il file `config.json` se desideri mantenere le tue impostazioni personalizzate.

---
Per segnalazioni bug o richieste di nuove funzionalità contatta email: gianlucadip1302@gmail.com
