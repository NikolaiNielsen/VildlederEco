#VildlederEco

Forfatter: Nikolai Nielsen (NikolaiNielsen@outlook.dk)
Et script til at hjælpe med at gøre Vildlederøkonomien lidt nemmere:
- Gennemgår alle arkene set i variablen "SHEETS", finder alle indtastninger, og
  laver en oversigt over, hvad hver person har udlagt, under hvilken kategori,
  og med hvilken betalingsmetode.

Fremtidige forhåbninger:
- Automatisk genkendelse af, hvor de relevante data er, i hvert ark, så det er
  mere robust
- Grafisk brugerflade for dataindtastning
- Grafisk brugerflade til valg af fil

BRUG AF DETTE PROGRAM
- Programmet bruger pakker fra standardbiblioteket, samt "openpyxl" (version
  3.0.0), der skal installeres separat, for eksempel gennem pip. Den er også
  inkluderet med Anaconda distributionen.
- For at benytte dette program skal du have en lokal kopi af regnskabet i
  xlsx-format. Du skal ændre "WORKBOOK_NAME" til den relative sti til
  regnearket. Hvis regnearket er i samme mappe som dette program, kan du bare
  sætte WORKBOOK_NAME til at være filens navn (husk fil-typen!).
- Du skal ændre "SHEETS" til en liste af strings over hvilke ark, der skal
  tjekkes igennem. Som regel behøver denne ikke at blive ændret fra de normale,
  med mindre, der bliver ændret kategorier af øko-gruppen efter dette program
  er oprettet.
- Programmet forventer at data er indtastet fra kolonne A til H, og fra række 4
  og nedad. Det forventes af kolonnerne er som følger:
  - A: Bilagskode - kategorien af indkøbet. eksempelvis "Tema"
  - B: Bilagsnummer - Indkøbets nummer. Starter med 01, så 02 og så videre.
                      Sammen med bilagskoden udgør dette en samlet "ID" for
                      købet, eksempelvis Tema01.
  - C: NBI Rekvireringsnummer - som regel ikke vigtig. Bruges ikke her
  - D: Tekst - beskrivelse af købet
  - E: Status på bilag - Hvor ligger kvitteringen henne? (den skal helst i
                         mappen eller digitalt på dropbox eller lign.)
  - F: Beløb - hvor meget har dette indkøb kostet?
  - G: metode - hvordan indkøbet er foretaget. Der er 3 måder pt, set under
                "PAYMENT_METHODS". I arket forventes det at dette er et heltal,
                svarende til positionen i PAYMENT_METHODS (så 0=kontokort, etc)
  - H: Navn - hvem har foretaget købet? Der skelnes mellem store og små
              bogstaver!
- Det er egentlig kun kolonnerne A, B, F, G og H, der bliver brugt.
- "RANGES" bør kun ændres, hvis layouttet af regnearket ændres. Ligeledes skal
  "NUM_COLS" svare til hvor mange kolonner der skal læses fra (8, i dette
  tilfælde, da der skal læses fra A til H).
- Når disse ting er sørget for, skal programmet bare køres. Så bliver der
  automatisk indlæst og oprettet et nyt ark, "Opsummering", hvor der står en
  opsummering over, hvad hvert "navn" har købt, hvilke kvitteringer samt samlet
  beløb for hver betallingsmulighed. Så kan man nemt se, hvis man har skrevet
  et navn forkert, og hvor mange penge, hver person skal have tilbage gennem
  REJS-ud, samt hvor mange penge, der skal tilbagebetales til vejlederkontoen
  (i form af tilskud)