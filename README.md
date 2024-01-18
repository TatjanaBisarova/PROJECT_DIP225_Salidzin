Projekts uzņēmumu S un L datu salīdzināšana par 2023.gada 5 mēnešiem
Savā ikdienas darbā katrs grāmatvedis sastopas ar lielu datu analīzi, kuru, diemžēl, mūsdienās spiests veikt manuāli. Pateicoties dotajam kursam un iegūtajam zināšanām un iemaņām man izdevās atvieglot savu ikdienas darbu, padarot to efektīvāku.
Lai ievēroto datu aizsardzības principus uzņēmuma nosaukumi tiek aizvietoti ar burtiem. Darbu sadalīju 3.posmos: 
1.	Bankas datu analīze S un L uzņēmumā;
2.	Programmas testēšana uz nelieliem testa failiem;
3.	Visu datu analīze. 
Darba gaitā nolēmu:
1.	bankas datu rezultātu apkopojumu analīzei nemainīt S uzņēmuma negatīvu rezultātu, jo pirmdokumentā bankas dati ir doti ka negatīvas vērtības;
2.	S uzņēmuma datos, sākumā , izmantoju excel funkciju priekšrocības, jo uzņēmums S veic datu korekciju šādi: (ir izrakstīts dokuments ar Nr. SNBL100100 par 100,00 EUR,  dokumentam ir veikta korekcija par -20,00 EUR, un tad vēl viena korekcija par – 5,00 EUR, dokumentu apgrozījumā tās izskatīsies šādi:
Nr. 	Summa
SNBL100100	100,00
SNBL100100	-100,00
SNBL100100_200123	80,00
SNBL100100_200123	-80,00
SNBL100100_250123	75,00
Un ņēmot vērā ka šādu korekciju vienam dokumentam varbūt daudz, izmantojot Excel funkciju MID izgriezu dokumenta skaitlisko vērtību un ar PIVOT Table palīdzību pārveidoju dokumentus, lai to vērtība būtu sastopamā tikai 1 reizi, atstājot tikai 2 kolonnas ar modificēto numuru un gala summu, līdzīgi rīkojos ar L dokumentu datiem.
3.	Paņemu dažas vērtības no S un L dokumenta tabulas datiem, uzrakstīju programmu, kura izvadīs visas nesakritības starp S un L uzņēmuma datiem. 
4.	Pielāgoju izveidoto programmu abu uzņēmumu visam datu apjomam un nesakritības rezultātu izvadīju tabulā.
Darba rezultāts:
1.posmā (bankas datu pārbaudē) programma darbojas korekti, dati ir salīdzināmi un rezultāts tiek sniegt minūtes laikā!
2.posmā (Programmas testēšana uz nelieliem testa failiem) programma darbojas korekti, dati ir salīdzināmi un rezultāts tiek sniegt ļoti  ātri!
3.posmā (Visu datu analīze) programmas darbojas, bet to korektumu pārbaudīt vēl neizdevās, bet pēc manām domām, dati tiek izvadīti nekorekti. Ir jāpiestrādā pie koda!
