Questo ZIP contiene SOLO i file da sostituire nel repo:
- app.py
- requirements.txt
- Dockerfile

Per usare OCR gratis su Render serve Docker, perche' il servizio Python standard non ha Tesseract.
Passi:
1. sostituisci questi file nel repo GitHub
2. fai commit
3. su Render crea un NUOVO Web Service di tipo Docker collegando lo stesso repo
4. usa il nuovo link

La tua UI attuale resta uguale.
Cambia solo il backend della scansione scontrino.
