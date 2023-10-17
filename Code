import imaplib
import email
from email.header import decode_header
import html2text #Roba per la mail
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By 
from selenium_stealth import stealth
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException  
import random #Roba per internet
import string
import openpyxl
import time
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
from openpyxl.comments import Comment
import calendar
import datetime #Roba per il file excel
from datetime import date
import sys

def mese_string(mese):
	if mese==1:
		monthString = "GENNAIO";
	elif mese==2:
		monthString = "FEBBRAIO";
	elif mese==3: 
		monthString = "MARZO";
	elif mese==4:
		monthString = "APRILE";
	elif mese==5:  
		monthString = "MAGGIO";
	elif mese==6:  
		monthString = "GIUGNO";
	elif mese==7:  
		monthString = "LUGLIO";
	elif mese==8:  
		monthString = "AGOSTO";
	elif mese==9:  
		monthString = "SETTEMBRE";
	elif mese==10: 
		monthString = "OTTOBRE";
	elif mese==11: 
		monthString = "NOVEMBRE";
	elif mese==12: 
		monthString = "DICEMBRE";
	else:
		print("Non ho riconosciuto il mese")
	return monthString

def mese_string_email(mese):
	if mese==1:
		monthString = "Jan";
	elif mese==2:
		monthString = "Feb";
	elif mese==3: 
		monthString = "Mar";
	elif mese==4:
		monthString = "Apr";
	elif mese==5:  
		monthString = "May";
	elif mese==6:  
		monthString = "Jun";
	elif mese==7:  
		monthString = "Jul";
	elif mese==8:  
		monthString = "Aug";
	elif mese==9:  
		monthString = "Sept";
	elif mese==10: 
		monthString = "Oct";
	elif mese==11: 
		monthString = "Nov";
	elif mese==12: 
		monthString = "Dec";
	else:
		print("Non ho riconosciuto il mese")
	return monthString

def mese_string_ita(mese):
	if mese==1:
		monthString = "gen";
	elif mese==2:
		monthString = "feb";
	elif mese==3: 
		monthString = "mar";
	elif mese==4:
		monthString = "apr";
	elif mese==5:  
		monthString = "mar";
	elif mese==6:  
		monthString = "giu";
	elif mese==7:  
		monthString = "lug";
	elif mese==8:  
		monthString = "ago";
	elif mese==9:  
		monthString = "set";
	elif mese==10: 
		monthString = "ott";
	elif mese==11: 
		monthString = "nov";
	elif mese==12: 
		monthString = "dic";
	else:
		print("Non ho riconosciuto il mese")
	return monthString

def ultima_colonna(annoo, mesee):
	fine_mese=calendar.monthrange(int(annoo), mesee)[1]
	wss=wb[mese_string(mesee)+ " "+ annoo]
	#Questo trova la colonna dell'ultimo giorno del mese
	for fine in range(1,100):
		cella=wss.cell(column=fine, row=1)
		valore=cella.value
		if type(valore)==datetime.datetime:
			if valore.day==fine_mese:
				break;
	return fine;

def is_in_list(elemento, lista):
	x=lista.count(elemento)
	if x>0:
		return True
	else:
		return False

def colora(cel, colore):
	aFill = PatternFill(start_color=str(colore), end_color=str(colore), fill_type='solid')
	cel.fill=aFill

def controlla_disp(wb, inizio, z, ws1, mese1, year1, ws2, sost, camera):
	print(inizio, z, mese1, year1)
	end=z
	fine=ultima_colonna(year1, mese1)
	if not ws1.title==ws2.title:
		end=fine
	aa=0
	check=1
	ws=ws1
	anno=int(year1)
	mese=mese1
	y=29
	while y>2:
		if is_in_list(y, camera):
			print(y)
			print("check is ", check)
			if not y==19:
				d=ws1.cell(column=inizio, row=y)
				if not d.value:
					check=1
					for v in range(inizio, end+1):
						ce=ws1.cell(column=v, row=y)
						if ce.value:
							print(ce.value)
							if sost:
								print("è entrato in sostituzione")
								sposta_cliente(wb, ws, ce, mese, anno, camera)
								y=8
							check=0
							break;
					if check:
						ws=ws1
						aa=0
						while not ws.title==ws2.title:
							print("Ciao")
							aa+=1
							if mese+aa==13:
								anno+=1
								mese=1
								aa=0
							ws=wb[mese_string(mese+aa)+ " "+ str(anno)]
							print(ws.title)
							if ws.title==ws2.title:
								end=z
							else:
								end=ultima_colonna(str(anno), mese+aa)
							print(end)
							for i in range(4, end+1):
								cel=ws.cell(column=i, row=y)
								print(y, i)
								print("cell value is ", cel.value)
								if cel.value:
									print("hello")
									check=0
									break;
							if check==0:
								break;
						print("check è ", check)
						if check:
							break;
		y-=1
	return y, check

def fine_prenotazione(cel, ws0, mese, anno, wb):
	ws=ws0
	v=cel.value
	col=cel.column
	ro=cel.row
	por=1
	endd=ultima_colonna(str(anno), mese)
	controllo=1
	aaa=0
	for lll in range(col+2, endd+1):
		if por:
			c=ws.cell(column=lll, row=ro)
			if not c.value==v:
				controllo=0
				break
			por=0
		else:
			por=1
	while controllo:
		messe=mese+aaa
		annno=anno
		wss=ws
		aaa+=1
		if mese+aaa==13:
			anno+=1
			mese=1
			aaa=0
		ws=wb[mese_string(mese+aaa)+ " "+ str(anno)]
		endd=ultima_colonna(str(anno), mese+aaa)
		por=1
		for lll in range(4, endd+1):
			if por:
				c=ws.cell(column=lll, row=ro)
				if not c.value==v:
					controllo=0
					if lll==4:
						ws=wss
						lll=ultima_colonna(str(annno), messe)+2
					break
				por=0
			else:
				por=1
	return (lll-2), ws

def sposta_cliente(wb, ws, cel, mese, anno, camera):
	z, ws2=fine_prenotazione(cel, ws, mese, anno, wb)
	print(z, "Sono in sposta_cliente " + ws2.title)
	col=cel.column
	ro=cel.row
	year=str(anno)
	commento=ws.cell(column=col, row=ro).comment
	cell_prezzo=ws2.cell(column=z+1, row=ro)
	prezzo1=cell_prezzo.value
	colore_prezzo=cell_prezzo.fill.start_color.index
	print("Adesso controlla la sua disponibilità")
	y, check= controlla_disp(wb, col, z, ws, mese, year, ws2, 0, camera)
	if not check:
		y, check= controlla_disp(wb, col, z, ws, mese, year, ws2, 1, camera)
	if check:
		print("Ora sta inserendo i suoi dati da un'altra parte")
		inserisci_cliente(wb, ws2, col, y, z, cel.value, mese, year, cel.fill.start_color.index, prezzo1, colore_prezzo, commento)
		inserisci_cliente(wb, ws2, col, ro, z, None, mese, year, 'FFFF0000', prezzo1, 'FFFFFF00', None)

def inserisci_cliente(wb, ws2, inizio, y, z, nome, mese1, year1, colore, prezzo, colore_prezzo, commento):
	end=z
	fine=ultima_colonna(year1, mese1)
	ws1= wb[mese_string(mese1)+" " +year1]
	por=1
	mese=mese1
	anno=int(year1)
	if not ws1.title==ws2.title:
		end=fine
	for v in range(inizio, end+2):
		if por:
			ce=ws1.cell(column=v, row=y)
			print("nome ", nome)
			print("col=", v, "row=", y)
			ce.value=nome
			colora(ce, colore)
			ce.comment=commento
			por=0
		else:
			ce=ws1.cell(column=v, row=y)
			print("Prezzo:", float(prezzo))
			print("col=", v, "row=", y)
			ce.value=float(prezzo)
			if v==(end+1):
				colora(ce, colore_prezzo)
			else:
				colora(ce, colore)
			por=1
	ws=ws1
	aa=0
	while not ws.title==ws2.title:
		aa+=1
		if mese+aa==13:
			anno+=1
			mese=1
			aa=0
		ws=wb[mese_string(mese+aa)+ " "+ str(anno)]
		if ws.title==ws2.title:
			end=z
		else:
			end=ultima_colonna(str(anno), mese+aa)
		por=1
		for i in range(4, end+2):
			if por:
				ce=ws1.cell(column=v, row=y)
				print("nome ", nome)
				print("col=", v, "row=", y)
				ce.value=nome
				colora(ce, colore)
				ce.comment=commento
				por=0
			else:
				ce=ws1.cell(column=v, row=y)
				print("Prezzo:", int(prezzo))
				print("col=", v, "row=", y)
				ce.value=int(prezzo)
				if v==(end+1):
					colora(ce, colore_prezzo)
				else:
					colora(ce, colore)
				por=1

def check_exists(nome, driver):
    try:
        driver.find_element(By.ID, nome)
    except NoSuchElementException:
        return False
    return True

def controlla_exists(nome, driver):
	try:
		driver.find_element(By.CSS_SELECTOR, nome)
	except NoSuchElementException:
		return False
	return True

def tipo_di_camera(camera):
	camere_singole=[6, 15, 26]
	camere_doppie=[3, 4, 5, 7, 9, 10, 13, 14, 18, 20, 21, 22, 24, 25, 29]
	camere_doppie_separate=[3, 5, 11, 12, 16, 27]
	camere_triple=[5, 23, 28]
	if camera=="Single Room" or camera=="Camera Singola":
		return camere_singole
	elif camera=="Double Room" or camera=="Camera Matrimoniale":
		return camere_doppie
	elif camera=="Camera Doppia con Letti Singoli" or camera=="Twin Room":
		return camere_doppie_separate
	elif camera=="Triple Room" or camera=="Camera Tripla":
		return camere_triple


controllo=0
if controllo==1:
	indirizzo=r"C:\Users\fabio\Desktop\\"
else:
	indirizzo='/Users/Locanda/Desktop/'

link=[]
nome=[]
arrivo=[]
partenza=[]
prezzi_totali=[]
ricevuto=[]
stanze_totali=[]
pagato=[]
controllore=0
username="" #Username erased for privacy
password="" #password erased for privacy
imap = imaplib.IMAP4_SSL('imap.gmail.com')
imap.login(username, password)
imap.select('inbox')
pincopallo, selected_mails = imap.search(None, '(FROM "noreply@booking.com" UNSEEN)')
print("Totale messaggi:" , len(selected_mails[0].split()))
n_mail=len(selected_mails[0].split())
for i in selected_mails[0].split():
	res, msg = imap.fetch(i, "(RFC822)")
	for response in msg:
		if isinstance(response, tuple):
			msg = email.message_from_bytes(response[1])
			subject, encoding = decode_header(msg["Subject"])[0]
			if isinstance(subject, bytes):
				print(type(encoding))
				if type(encoding)== str :
					subject = subject.decode(encoding, 'ignore')
					print("Decodifica")
			subject=str(subject)
			subject = subject.replace("b'", "")
			print("Subject:", subject)
			if subject[:25]!="Booking.com - New booking" and subject[:37]!="Booking.com - New last-minute booking" and subject[:40]!="Booking.com - Hai una nuova prenotazione" and subject[:32]!="Booking.com - Nuova prenotazione":
				n_mail-=1
			else:
				body = msg.get_payload(decode=True).decode()
				html = body.replace("b'", "")
				h = html2text.HTML2Text()
				output = (h.handle(f'''{html}''').replace("\\r\\n", ""))
				output = output.replace("'", "")
				output = output.replace("[", " ")
				output = output.replace("]", " ")
				output = output.replace("(", " ")
				output = output.replace(")", " ")
				output = output.split()
				#print(output)
				print("Sto prendendo il link di", subject)
				for x in range(len(output)):
					if output[x][:32]=="https://admin.booking.com/hotel/":
						if output[x+1][:35]=="https://clicks.booking.com/ls/click":
							link1=output[x+1]
						else:
							link1=output[x]
							if output[x][-6:]!="mail=1":
								link1=output[x]+output[x+1]
						link.append(link1)
						#print(link1)
				break
imap.close()
imap.logout()

if n_mail!=0:
	print("Soooo")
	options = webdriver.ChromeOptions()
	options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36")
	options.add_experimental_option("useAutomationExtension", False)
	options.add_experimental_option("excludeSwitches", ["enable-automation"])
	options.add_argument('--disable-blink-features=AutomationControlled')
	options.add_argument("disable-popup-blocking")
	options.add_argument("disable-notifications")
	options.add_argument("disable-gpu")
	if controllo==1:
		options.add_argument(r"user-data-dir=C:\Users\fabio\AppData\Local\Google\Chrome\User Data\Default")
	else:
		options.add_argument("user-data-dir=/Users/Locanda/Library/Application Support/Google/Chrome/Profile 1")
	driver=webdriver.Chrome(options=options)
	stealth(driver,
	    languages=["en-US", "en"],
	    vendor="Google Inc.",
	    platform="Win64",
	    webgl_vendor="Intel Inc.",
	    renderer="Intel Iris OpenGL Engine",
	    fix_hairline=True,
	    )
	for j in link:
		print("link è ", j)
		driver.get(j)
		time.sleep(0.87921)
		if (check_exists("loginname", driver)):
			elem=driver.find_element(By.ID,"loginname")
			action = ActionChains(driver)
			elem.send_keys(Keys.ENTER)
			action.perform()
			time.sleep(1.5849147)
			elem=driver.find_element(By.ID,"password")
			elem.send_keys(Keys.ENTER)
			action.perform()
		while driver.current_url[:13]!="https://admin":
			time.sleep(0.1)
		time.sleep(2)
		if controlla_exists("span[data-test-id=reservation-overview-name]", driver):
			nome1=driver.find_element(By.CSS_SELECTOR, "span[data-test-id=reservation-overview-name]").text
			stanze=driver.find_elements(By.CLASS_NAME, "res-room-title__name")
			prezzi=driver.find_elements(By.CLASS_NAME, "bui-price-display__value")
			arrivo1=driver.find_element(By.XPATH, "/html/body/div[1]/main/div/div/main/div/div[2]/div[1]/div/div[2]/div/div/div[1]/div/p[2]").text
			partenza1=driver.find_element(By.XPATH, "/html/body/div[1]/main/div/div/main/div/div[2]/div[1]/div/div[2]/div/div/div[1]/div/p[4]").text
			boh=driver.find_element(By.XPATH, "/html/body/div[1]/main/div/div/main/div/div[2]/div[1]/div/div[2]/div/div/div[2]/div/div").text
			prezzo_controllo=driver.find_element(By.XPATH, "/html/body/div[1]/main/div/div/main/div/div[2]/div[1]/div/div[2]/div/div/div[1]/div/p[12]").text
			lun=len(stanze)
			print("num prezzi ", len(prezzi))
			for i in range(lun):
				stanze[i]=stanze[i].text
				if lun!=1:
					stanze[i]=stanze[i].split(' ', 1)[1]
			for i in range(len(prezzi)):
				prezzi[i]=prezzi[i].text.split()[1]
				prezzi[i]=prezzi[i].replace(",", ".")
				prezzi[i]=float(prezzi[i])
				print("BBBBB", prezzi[i])
			pagato.append(0)
			arrivo1=arrivo1.split(' ', 1)[1]
			partenza1=partenza1.split(' ', 1)[1]
			boh=boh.split()
			lunghezza=len(boh)
			for i in range(lunghezza):
				if boh[i]=="Received" or boh[i]=="Ricevuta":
					ricevuto1=boh[i+2]+ " " + boh[i+3]+ " " + boh[i+4]
				elif boh[i]=="guest" and boh[i+1]=="has" and boh[i+2]=="paid" and boh[i+3]=="for":
					pagato.pop()
					pagato.append(1)
				elif boh[i]=="ha" and boh[i+1]=="pagato" and boh[i+3]=="prenotazione" and boh[i+4]=="online.":
					pagato.pop()
					pagato.append(1)
			if prezzo_controllo!="€ 0":
				stanze_totali.append(stanze)
				prezzi_totali.append(prezzi)
				arrivo.append(arrivo1)
				partenza.append(partenza1)
				nome.append(nome1)
				ricevuto.append(ricevuto1)
			else:
				pagato.pop()
		else:
			controllore=1
	driver.quit()

	wb = load_workbook(indirizzo+ 'PLANNING LOCANDA.xlsx')
	if controllo!=1:
		wb.save(indirizzo+ 'Planning Locanda Vecchio/' + str(date.today())+".xlsx")
	for i in range(len(nome)):
		stanza=stanze_totali[i]
		prezzo=prezzi_totali[i]
		print("AHHHH ", prezzo)
		mese1=0
		bohh=arrivo[i].split()
		for j in range(1, 13):
			if (mese_string_email(j)==bohh[1]):
				mese1=j
		if mese1==0:
			for j in range(1, 13):
				if (mese_string_ita(j)==bohh[1]):
					mese1=j
		monthString1=mese_string(mese1)
		year1=str(bohh[2])
		bohh[0]=int(bohh[0])
		if bohh[0]<10:
			giorno1="0"+str(bohh[0])
		else:
			giorno1=str(bohh[0])

		if mese1<10:
			da= year1 + "-" + "0" + str(mese1) + "-" + giorno1
		else:
			da= year1 + "-" + str(mese1) + "-" + giorno1

		data1=da
		bohh=partenza[i].split()
		mese2=0
		for j in range(1, 13):
			if (mese_string_email(j)==bohh[1]):
				mese2=j
		if mese2==0:
			for j in range(1, 13):
				if (mese_string_ita(j)==bohh[1]):
					mese2=j
		monthString2=mese_string(mese2)
		year2=str(bohh[2])
		bohh[0]=int(bohh[0])
		print("Il giorno è", bohh[0])
		print(type(bohh[0]))
		if (bohh[0]!=1):
			giorno2=str(bohh[0]-1)
			if (bohh[0]-1)<10:
				giorno2="0"+giorno2
			print("Il nuovo giorno è", giorno2)
		else:
			mese2-=1
			giorno2=str(calendar.monthrange(int(bohh[2]), mese2)[1])
			print("il nuovo giorno è", giorno2)
		if mese2<10:
			da= year2 + "-" + "0"+ str(mese2) + "-" + giorno2
		else:
			da= year2 + "-" + str(mese2) + "-" + giorno2
		data2=da
		pay=pagato[i]
		if pay==1:
			colore_prezzo='FF7CFC00'
		else:
			colore_prezzo='FFFFFF00'
		ws1= wb[monthString1 +" " +year1]
		ws2= wb[monthString2+ " "+ year2]
		data_checkin=date(int(year1), mese1, int(giorno1))
		data_checkout=date(int(year2), mese2, int(giorno2))
		delta=data_checkout-data_checkin
		n_giorni=delta.days+1
		for cont in range(len(prezzo)):
			prezzo[cont]=prezzo[cont]/n_giorni
			print("Prezzo[cont]=", prezzo[cont])
		#Questo trova la data del check-out: z è il numero della colonna
		for z in range(1, 80):
			c=ws2.cell(column=z, row=1)
			val=c.value
			val=str(val)
			val=val[:10]
			if val==data2:
				break;
		end=z
		fine=ultima_colonna(year1, mese1)
		if not ws1.title==ws2.title:
			end=fine
		#Questo trova la colonna che contiene la data del check-in
		for inizio in range(1, fine+1):
			c=ws1.cell(column=inizio, row=1)
			val=c.value
			val=str(val)
			val=val[:10]
			if val==data1:
				break;
		#comment=ws1.cell(column=inizio, row=y).comment
		comment = Comment("Programma Fabio: Ricevuto da Booking il " + str(ricevuto[i]), "Programma Fabio")
		
		for jiji in range(len(stanza)):
			camera=tipo_di_camera(stanza[jiji])
			y, check= controlla_disp(wb, inizio, z, ws1, mese1, year1, ws2, 0, camera)
			if not check:
				y, check= controlla_disp(wb, inizio, z, ws1, mese1, year1, ws2, 1, camera)
			print("Check èèè", check)
			if check:
				print("Ora mette i miei dati")
				print("jiji ", jiji)
				print(type(jiji))
				print("i è ", i)
				print(type(i))
				inserisci_cliente(wb, ws2, inizio, y, z, nome[i], mese1, year1, 'FF7030A0', prezzo[jiji], colore_prezzo, comment)
		
		#ws1.cell(column=inizio, row=y).comment=comment
		wb.save(indirizzo+ 'PLANNING LOCANDA.xlsx')
	if controllore==1:
		print("UNA O PIù PRENOTAZIONI SONO STATE IGNORATE")
