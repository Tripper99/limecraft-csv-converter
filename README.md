# Limecraft till Word/Inqscribe-konverterare med justering av starttid

Detta är en windows-app som löser problemet att transkriberingssystemet Limecraft inte har någon funktion för att ändra transkriptionens starttid så att den matchar källmaterialets starttid. 
Alla transkriptioner som Limecraft gör startar således på 00:00:00:00 och det gör resultatet klumpigt att använda om fotografen som filmade källmaterialet hade sin starttid satt till 01:54:18:03. Appen låter dig importera Limecraft-transkriberingen i form av en csv-fil, ange valfri starttid och exportera i form i Word- och/eller Inqscribe-format. 

## Så här använder du programmet:

### 1. Välj CSV-fil (som du exporterats från Limecraft)
Klicka på "Bläddra CSV-fil..." och välj den CSV-fil som du har exporterat från Limecraft.
Programmet kommer att läsa in filen och visa dess namn i gränssnittet här nedanför.
(För att exportera din transkribering som csv-fil från Limecraft välj "Export" och klicka på "CSV-fil".

**Viktigt!** När du gör din export i Limecraft **måste rutorna framför** *"Media Start"*, *"Transcript"* och *"Speakers"* **vara ikryssade**. Det är ok om fler rutor är ikryssade men ingen annan metadata dessa tre fält används.

### 2. Justera starttid (valfritt)
Ange tid som ska LÄGGAS TILL alla tidskoder. T ex starttiden på kamerakortet.
Tidkoden du lägger till kommer att omvandlas till formatet HH:MM:SS.FF (timmar:minuter:sekunder.frames).
Det går bra att mata in tidkoden i de här formaten:
- HH:MM:SS:FF (t.ex. 01:30:45:12)
- HH:MM:SS.FF (t.ex. 01:30:45.12)
- HH.MM.SS.FF (t.ex. 01.30.45.12)
- HHMMSSFF (t.ex. 01304512)

### 3. Välj filnamn (valfritt)
Ange önskat filnamn (utan filnamnstillägg).
Samma namn kommer att användas för både .docx- och .inqscr-filer.
Om du inte skriver in något namn används filnamnet från den valda CSV-filen (men med nya filnamnstillägg).

### 4. Välj utdataformat
Markera om du vill ha ut transkriberingen som Word-dokument (.docx) och/eller som en Inqscribe-fil (.inqscr).

### 5. Välj om du vill lägga till filnamnet före varje tidskod
Du kan här välja att lägga till filnamnet inom parentes före varje tidskod:
(Filnamn) [HH:MM:SS.FF]
Exempel: Om ditt filnamn är "2025-06-07 synk Christopher Hitchens" så kommer alla tidkoder i transkriberingsfilen se ut så här:
(2025-06-07 synk Christopher Hitchens) [HH:MM:SS.FF]
Detta kan vara en fördel om du exempelvis i ditt manus blandar bitar ur flera intervjuer gjorda med samma person på olika datum. Du slipper leta efter rätt synk eftersom alla synkbitar i manus även visar källan.  

### 6. Konvertera
Klicka på "Konvertera". Du får då välja en mapp där filerna ska sparas.

### 7. Klart!
Programmet lägger i början av dokumentet automatiskt till den nya starttid som du själv valt, alternativt [00:00:00.00] om du inte ändrat starttiden. 

### Installation
Ladda ned och starta exe-filen. Ingen installation behövs. 

## Upphovsman
Appen är gjord av Dan Josefssson (https://github.com/Tripper99).
Buggar rapporteras till dan@josefsson.net.

## License
This project is licensed under the MIT License - see the LICENSE file for details.
