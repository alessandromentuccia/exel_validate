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
- prime visite
- controlli
- esami strumentali

aggiungerei per QD metodiche e distretti anche l'assenza di spazi dopo il separatore e l'aggiunta dei codici qualora l'ente inserisse le descrizioni.

a me è capitato di trovare anche codici o descrizioni duplicati nella stessa cella.
aggiungerei il controlli dei casi 1:n con abilitazione esposizione siss "S"

### domande da fare:

- che ordinamento ha il mapping? au che campo è ordinato? 
- l'ordinamento è sempre lo stesso?
- il campo agenda da considerare è Codice SISS Agenda?
- i campi codice distretto e descrizione distretto 


data = OrderedDict()
if row["question_number"] not in data_list["question_number"]:
    data[row[1]] = [row["question_template"]]
    self.file_data.update({ row["question_number"]: [row["question_type"], row["question_number"], row["question_template"], row["answer"], row["note"]]})
    data_list["question_number"].update(data)
    print("data_list: %s", data_list)
else: 
    data_list["question_number"][row["question_number"]].append(row["question_template"])
    self.file_data.update({ row["question_number"]: [row["question_type"], row["question_number"], row["question_template"], row["answer"], row["note"]]})

