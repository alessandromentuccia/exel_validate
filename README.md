# exel_validate

A partire dal file di mapping compilato dalla struttura (che di solito contiene migliaia di righe), sarebbe utile avere un automatismo in grado di verificare:

### AMBITO QD
- per ogni Agenda sono inseriti gli stessi QD
- tutti i QD di una determinata Agenda appartengono alla stessa disciplina
- tutti i QD hanno lo stesso separatore (la , )
- se non ci sono le descrizioni corrispondenti vengono inserite

### AMBITO METODICHE
- per ogni prestazione sono inserite le Metodiche di pertinenza
- tutte le metodiche hanno lo stesso separatore
- se non ci sono le descrizioni corrispondenti vengono inserite

### AMBITO DISTRETTI
- per ogni prestazione sono inserite i Distretti di pertinenza
- tutti i distretti hanno lo stesso separatore
- se non ci sono le descrizioni corrispondenti vengono inserite

### AMBITO PRIORITA' e TIPOLOGIE DI ACCESSO
- prime visite - controllare se le priorità ci sono
- controlli - controllare se Accesso Programmabile ZP
- esami strumentali

aggiungerei per QD metodiche e distretti anche l'assenza di spazi dopo il separatore e l'aggiunta dei codici qualora l'ente inserisse le descrizioni.

a me è capitato di trovare anche codici o descrizioni duplicati nella stessa cella.
aggiungerei il controlli dei casi 1:n con abilitazione esposizione siss "S"

controllo colonna 1:n e filtrare su una agenda e una prestazione e controllare che ci sia sono una prestazione abilitata al SISS
### domande da fare:

- che ordinamento ha il mapping? au che campo è ordinato? 
- l'ordinamento è sempre lo stesso?
- il campo agenda da considerare è Codice SISS Agenda?
- i campi codice distretto e descrizione distretto 
- esami strumentali come faccio a riconoscerli? 
- cosa devo controllare che ci sia negli esami strumentali?


data = OrderedDict()
if row["question_number"] not in data_list["question_number"]:
    data[row[1]] = [row["question_template"]]
    self.file_data.update({ row["question_number"]: [row["question_type"], row["question_number"], row["question_template"], row["answer"], row["note"]]})
    data_list["question_number"].update(data)
    print("data_list: %s", data_list)
else: 
    data_list["question_number"][row["question_number"]].append(row["question_template"])
    self.file_data.update({ row["question_number"]: [row["question_type"], row["question_number"], row["question_template"], row["answer"], row["note"]]})


- COSE DA FARE:
1. check ck_esami_strumentali
2. controllare se posso separare dei check in ck_QD_disciplina_agenda
3. check ck_casi_1n
4. studiare come riportare i risultati in modo ottimale
5. verifica dello script con file di mapping nuovo
6. verifica che tutto funzioni anche se i campi sono vuoti
7. verifica che tutto funzioni e gestisca le eccezioni nel caso non si trovi le colonne
8. creare iniziazione dove definire l'indice delle colonne da analizzare
9. studiare come rendere accessibile il prodotto da server. Es. FLASK
10. creare una interfaccia grafica per semplificare la fase iniziale
11. definire un check per verificare che le descrizioni abbiano il codice
12. check per descrizioni duplicati nella stessa cella
13. stampare il risultato in un file creato ex novo