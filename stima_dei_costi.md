# exel_validate

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
