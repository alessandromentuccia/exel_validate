# exel_validate

A partire dal file di mapping compilato dalla struttura (che di solito contiene migliaia di righe), sarebbe utile avere un automatismo in grado di verificare:

### AMBITO QD
- per ogni Agenda sono inseriti gli stessi QD
- tutti i QD di una determinata Agenda appartengono alla stessa disciplina
- tutti i QD hanno lo stesso separatore (la , )
- se non ci sono le descrizioni corrispondenti vengono inserite
- operatore logico coerente con agenda

### AMBITO METODICHE
- per ogni prestazione sono inserite le Metodiche di pertinenza
- tutte le metodiche hanno lo stesso separatore
- se non ci sono le descrizioni corrispondenti vengono inserite
- operatore logico coerente con prestazione

### AMBITO DISTRETTI
- per ogni prestazione sono inserite i Distretti di pertinenza
- tutti i distretti hanno lo stesso separatore
- se non ci sono le descrizioni corrispondenti vengono inserite
- operatore logico coerente con prestazione

### AMBITO PRIORITA' e TIPOLOGIE DI ACCESSO
- prime visite - controllare se le priorità ci sono
- controlli - controllare se Accesso Programmabile ZP
- esami strumentali

### AMBITO UNIVOCITA' PRESTAZIONE
- controllo univocità dei casi segnalati 1:n con abilitazione esposizione siss "S"


aggiungerei per QD metodiche e distretti anche l'assenza di spazi dopo il separatore e l'aggiunta dei codici qualora l'ente inserisse le descrizioni.

a me è capitato di trovare anche codici o descrizioni duplicati nella stessa cella.
aggiungerei il controlli dei casi 1:n con abilitazione esposizione siss "S"

controllo colonna 1:n e filtrare su una agenda e una prestazione e controllare che ci sia solo una prestazione abilitata al SISS

### domande da fare:

- che ordinamento ha il mapping? su che campo è ordinato? ordinato rispetto codice agenda SISS
- l'ordinamento è sempre lo stesso?
- il campo agenda da considerare è Codice SISS Agenda? si
- i campi codice distretto e descrizione distretto?
- esami strumentali come faccio a riconoscerli? 
- cosa devo controllare che ci sia negli esami strumentali?
- perchè sul catalogo ci sono dei codici disciplina di questo tipo? 
09
11
98 


- COSE DA FARE:
1. check ck_esami_strumentali
2. controllare se posso separare dei check in ck_QD_disciplina_agenda
3. check ck_casi_1n $OK$
4. studiare come mostrare i risultati in modo più pulito
5. verifica dello script con file di mapping nuovo
6. verifica che tutto funzioni anche se i campi sono vuoti
7. verifica che tutto funzioni e gestisca le eccezioni nel caso non si trovi le colonne
8. creare iniziazione dove definire l'indice delle colonne da analizzare
9. studiare come rendere accessibile il prodotto da server. Es. FLASK
10. creare una interfaccia grafica per semplificare la fase iniziale
11. definire un check per verificare che le descrizioni abbiano il codice
12. check per descrizioni duplicati nella stessa cella
13. stampare il risultato in un file creato ex novo $OK$
14. in ck_QD_agenda fare check solo per le agende con abilitazione ed esposizione SISS == S $OK$
15. problemi in ck_QD_discplina_descrizione nel caso QD presenta uno spazio alla fine e nel caso la descrizione presenta uno spazio dopo la virgola. eseguire quindi un controllo sugli spazi in eccesso definendo un errore, di seguito eliminare gli spazi e verificare che le descrizioni riportate sono corrette
16. modificare ck_QD_sintatti e gli altri check sintassi in modo che prima si verifica che non ci siano spazi in cima e in fondo, poi controllare che non ci siano spazi in mezzo o caratteri speciali. Deve fornire 3 tipologie
17. creare una documentazione per il rilascio
18. definire i check per gli operatori logici per QD, metodiche e distretti 
19. definire un container docker

fare piano ferie $$
mandare mail provisiong
call alle 10 $$
chiedere del lettore carta
chiedere come avere la vpn bvtech per 
stima progetto