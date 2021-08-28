# UV_generator v0.50

Skript pro vytvoreni dokumentu Uvolneni software. 
Na zadane ceste prodje vsechny soubory ve slozkach, nektere vyfiltruje, ulozi do pole.
Otevre dokument UV a podlde tagu a velikosti tabulek vyplni nalezene soubory z pole na prislusna mista.

# PODMINKY POUZITI
- nainstalovat python3 - https://www.python.org/downloads/
- nainstalovat python-docx modul - do cmd dat: pip install python-docx

# POUZITI
- spustit z prikazove radky: G:\$_WORK\Rulík\UV_generator\UV_generator>python UV_generator.py "G:\Vyvoj\AS\AS-MET\SW\Programova sada\210300_AS-MET"
	- !!! cesta musi byt ohranicena uvozovkami "" !!!

-spustit dvojklikem a zadat cestu pri vyzadani na vstupu. Zde pouzit cestu bez uvozovek.

# OMEZENI
- Vyplnovani tabulek se sofware (tabulky u tagu <SOFTWARE>) se identifikuji podle velikosti (pocet radku a sloupcu) -> nemenit UV.docx, s kterym skript pracuje (#TODO - zmenit detekci na tagy)	
														    -> v UV.docx natvrdo pridan radek u tabulky s kontaktnimi osobami


# ZPUSOB PRUCHODU VYPISU
-detekce aplikaci: pokud je na ceste slozka "_pkg_" a "_Project_" a pak jedno z "AIPM","MS","AIPD","AIPQ","APB","PPMB" (#TODO - mozna nektery pridat?)
			- otevre prislusnou slozku a ulozi do pole .tar.gz + udaje z manifestu (pid, major,minor,comment) -> vypise
-detekce slozky "OS" v "AIPM","MS","AIPD","AIPQ","APB","PPMB" -> najde soubor s ".squashfs" (komprimovany OS linux) a vypise

-detekce slozky "_pkg_" -> natvrdo vypisu IscConfig.exe, aniz bych se dival (#TODO - projit tu slozku a realne ISCConfig pridat - mozna zbytecne + je tam vic souboru s nazvem IscConfig.exe)
-detekce slozky "Tools" na ceste s "_pkg_" -> z teto slozky ani podslozek nic nevypisuju
-detekce slozky "Project" na ceste s "_pkg_" -> z teto slozky nic nevypisu (muze byt hafo .xml)
-detekce slozky "AXIS-P3905R" na ceste s "_pkg_" a "Project" -> pridam JEDEN .cmt (ten prvni), pridam vsechny .bin
							     -> zjistim, jestli je na ceste tag "Mk" -> podle toho nastrelim cestu k FW a komentar k FW
							     -> pri vypisu .cmt je komentar natvrdo "Konfigurace pro IP kamery AP3905-R verze <DOPLNIT>" (kamera bez MkII)
-detekce slozky OS v "AIPM","MS","AIPD","AIPQ","APB","PPMB" 

-detekce slozky "ufd-config"  -> natvrdo vypisu "UsbFlashDiskConfigurator.exe"
			      -> natvrdo vypisu verzi 1.0.1 + cestu ke zdroji + okmentar
- zbyle soubory vypisuju tak, jak je najdu

# TODO
-pokud narazim na .zip tak doplnim do vystupniho adresar i nazev zipu (bez .zip) -> OS,BIOSy mimo pkg jsou casto v korenovym adresari prog. sady
		-> obecne smazat cokoli za "."? (asfasdfasdf.tar.gz napr. u konfigutavi routeru)
detekce verze FW podle nazvu souboru -> pridat do vypisu
	



#TODO - filtr na pdf v root složce