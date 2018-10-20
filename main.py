# -*- coding: utf-8 -*-
"""
Created on Sat Oct 13 09:11:24 2018

@author: OB-2"""

import os
from time import sleep
import datetime
import csv
from urllib.request import urlopen
#import xxmpp
from openpyxl import Workbook
from openpyxl import load_workbook
import random
import socket
import socks
from bs4 import BeautifulSoup

from telethon import TelegramClient, events, sync

api_id =412522
api_hash = 'ccb44bc32a25e2051116f6a251f7117e'
client = TelegramClient('session_name1', api_id, api_hash)
client.start()
client.send_message('@parsfirmy_bot','start')

#STATIC
ipcheck_url='http://checkip.amazonaws.com/'
pars_IP='https://www.unternehmen24.info/Handelsregister/Deutschland/Neueintragung'
#SOCKS_IP='35.185.187.216'
#SOKCS_PORT=1080
norm_lkn='https://www.unternehmen24.info/Handelsregister/Deutschland/Neueintragung/Firma/'
poll=set()
spli='EUR'



def send_telega(msg):
	try :
		sleep(3)
		client.send_message('@parsfirmy_bot',msg)
	except:
		print ('ne otpravleno')


def connect():
	try:        
		chek=urlopen("https://www.unternehmen24.info/Handelsregister/Deutschland/Neueintragungen/")
		if chek.getcode()==200:
			return True
		else:
			return False
	except Exception as ex:
		print (ex)        
		send_telega('perebor noskov')
		
##connect JABBER 
#def jabber_auth():
#    data = list()
#    with open('jabber_cred.txt') as cred:
#        for str in cred:
#            data.append(str.split(':')[1].rstrip())
#    return data

#def send_to_jabber(msg, username, passwd, to, server, port):
#    client = xxmpp.Client(server)
#    client.connect(server=(server, port))
#    client.auth(username, passwd, 'botty')
#    client.sendInitPresence()
#    message = xxmpp.Message(to, msg)
#    message.setAttr('type', 'chat')
#    client.send(message)
#    print ('send_to_jabber')




def parsing_data_a17 (tds):
	a17 = str(tds[0]).split(':')[0].split('  ')[-1] #a17
	if a17.endswith('Aktenzeichen'):
		a17=a17.split('Aktenzeichen')[0]
		return a17
	else:
		a17='no_pars'
		return a17
	
def parsing_data_a18(tds):
	a18 = str(tds[0]).split(':')[1].split('\n')[0] #a18 Aktenzeichen :  HRB 14316
	return a18


def parsing_data_a19(tds):
	a19 = str(tds[1]).split('\n')[1].split('Bekannt gemacht am:')[-1] #a19 Datum der Eintragung  :  15.09.2018  02:02 Uhr
	a19='Datum der Eintragung  :  '+a19
	return a19

def parsing_data_b9(soup):
	name_firm=soup.findAll('h4')[0].findAll('strong')[0] # nazvanie firmi
	name_firm=str(name_firm).split('<br/>')[0].split('<strong>')[1] # nazvanie firmi
	return name_firm

	
def parsing_data_b10(tds,soup,b_9):
	try:
		b10 = str(tds[6]).split(':')[1].replace('  ','').replace('\n','').split('Name')[0].split(',')[-2]  #b10 ?????°???¤???«? ?­?­?®??
	except IndexError as ex:
		b10=str(tds[6]).split(':')[1].split('.')[0].replace('\n     ', '|') #  ???® ?·???® ???§?­? ?·? ?«???­?®
		b10=(b10.split('|')[0])[1:]
		b10=b10.split('|')[1]
	except IndexError as ex:
		b10=str(soup.select('td')[6].get_text())
		b10=b10.split(b_9)[1].split(',')[2].replace('\n','').replace('  ','')
		print ('==no==pars==')
	except:
		print ('da nu nahuj')
		send_telega('pidory nemcy')
		b10='----'
	return b10

def parsing_data_b11(tds):
	try:
		b11 = str(tds[6]).split(':')[1].replace('  ','').replace('\n','').split('Name')[0].split(',')[-1].split('.')[0] #b11
	except Exception as ex:
		b11=''
	return b11


def parsing_data_EUR(tds):
	try:
		text = str(tds[6]).split('.')
		euro=''
		ind=0
		for tx in text:
			ind = ind + 1
			if tx.find('Stammkapital') != -1:
				euro=tx[1:].replace('      ', 
					   ' ').replace('      ', '') + text[ind]
				euro=str(euro)
				euro=euro.split(spli)
				euro=euro[0].replace('\n', '')
				euro=euro+spli
	except Exception as ex:
		euro=''
	return euro


	
 #SAVE EXCEL
def save_to_xlsx (a_18,a_19,adr,adr_index,a_17, kol_vo_, kol_par, b_9, ext_norm ):
	os.makedirs(kol_vo_, exist_ok=True)
	wb = load_workbook ('Rechnung obrazec.xlsx')
	wb_sh= wb['Sheet1']
	wb_sh.cell(row=18, column=1,value='Aktenzeichen : '+a_18)# Aktenzeichen :  HRB 14316     (bei Zahlungen angeben!)  
	wb_sh.cell(row=19, column=1,value=a_19)  #Datum der Eintragung  :  15.09.2018  02:02 Uhr
	wb_sh.cell(row=9, column=2,value=b_9) #?­? ?§??? ?­???? ?????°?¬?»
	wb_sh.cell(row=10, column=2,value=adr) #adres
	wb_sh.cell(row=11, column=2,value=adr_index) # index gorod
	wb_sh.cell(row=17, column=1,value='Registergericht : '+a_17)
	wb_sh.cell(row=62, column=1,value=a_18)
	save_file='./'+kol_vo_+os.sep+kol_par+ext_norm+'.xlsx'
	print (save_file)
	wb.save(save_file)
	print ('save_to_xlsx')   
	
	
def save_to_cvs ():
	with open('data.csv', 'a') as csvfile:
		fieldnames = ['line1', 'line2', 'line3', 'line4', 'line5', 'line6']
		writer = csv.DictWriter(csvfile,fieldnames=fieldnames )
		writer.writeheader()
		#TODO writer.writerow({'line1': a18, 'line2': a19, 'line3': b9, 'line4': l4, 'line5': l5, 'line6': l6})




def noski():
    data = list()
    with open('noski.txt') as cred:
        for ip in cred:
            ip=ip.split(' [SOCKS5] ')[1].replace('>','').replace('\n','')            
            data.append(ip)
    return random.choice(data)


		  
def connect_retry ():
	n=0
	ip_port=noski()
	print (ip_port)
	send_telega(ip_port)
	SOCKS_IP=ip_port.split(':')[0]
	SOKCS_PORT=int(ip_port.split(':')[1])
	sleep(2)
	socks.set_default_proxy(socks.SOCKS5, SOCKS_IP, SOKCS_PORT )
	socket.socket=socks.socksocket
	try:
		if connect():
			chek=urlopen("https://www.unternehmen24.info/Handelsregister/Deutschland/Neueintragungen/")
			#print (chek)
			print (ip_port,'connect true')
			return True
		else:              
			
			
			print (ip_port,'udalit iz spiska')
			return False
	except Exception as ex:
		print (ex)
		connect_retry()
	
	
	

def main():

	while True:
		sleep(300)
		print ('pause 300 sec')
		# ustanovka siedineni9
		connect_retry()
		

		
		#connect SOCKS & check
		
		
		
		if connect()==True:
			print ('soedinilisi')
			client.send_message('@parsfirmy_bot', 'na4inaem rabotu s noskom' )
			sleep(3)
			
			client.send_message('@parsfirmy_bot','rabotaem')
			try:
				page = urlopen('https://www.unternehmen24.info')
			except:
				main()
			try:
				page = urlopen('https://www.unternehmen24.info/Handelsregister/Deutschland/Handelsregisterauszug')
			except:
				main()
			try:
				page = urlopen('https://www.unternehmen24.info/Handelsregister/Deutschland/Neueintragungen')
			except:
				main()
			#page = urlopen('https://www.unternehmen24.info/Handelsregister/Deutschland/Neueintragung')
			bsObj = BeautifulSoup(page.read(), 'lxml')
			kol_vo_=bsObj.find('h4').get_text()
			kol_vo_=str(kol_vo_).replace(' ','_').replace(':','_').replace('Neueintragungen_','')
			kol_vo_=kol_vo_.replace('.2018_','')
			print (kol_vo_,'====kol_vo_')
			drt=kol_vo_.split('_')[-2]
			drt=drt.replace('2018', '')
			drt=drt.replace('.','_')   #12_10_85
			print (drt,'=====drt')
			nqazv=kol_vo_.split('_')[0] # ?????????? ???? ?? ????? 
			print (nqazv, '====nqazv')
			print (kol_vo_,'====kol_vo_')
			# ????? IP  ?????
			for item_url in bsObj.select('a'):
				if 'href' in item_url.attrs:
					if item_url.attrs['href'] not in item_url:
						#???????? ?????? 
						npage=item_url.attrs['href']
						if npage.startswith(norm_lkn):
							item_url_pars=npage
							if item_url_pars not in poll:
								# ???????? ?????? ???????
								poll.add(item_url_pars)
								kol_par=str(item_url_pars.replace(norm_lkn, '')+'_')
								adr_par=item_url_pars.replace(norm_lkn, '')
								print (kol_par,'=====kol_par')
								print(len(poll),'kol-vo sparsenogo')
								print (item_url_pars)
								now = str(datetime.datetime.now()).split('.')[0].replace(' ',
										 '_').replace(':','_')
								now='KGD_sprs_'+now
								
								#print (npage)
								
								page_pars=urlopen(npage)
								soup=BeautifulSoup(page_pars.read(), 'lxml')
								table = soup.find('table')
								tds=table.find_all('td')
								b_9=parsing_data_b9(soup)
								print(b_9,'====b_9=====')
								a_17=parsing_data_a17(tds)
								print (a_17,'====a17====')
								#b_10=parsing_data_b10(tds,soup,b_9)
								#print (b_10,'=====b_10=====')
								a_18=parsing_data_a18(tds)
								print (a_18,'=====a_18=====')
								a_19=parsing_data_a19(tds)
								print(a_19)
								#b_11=parsing_data_b11(tds)
								#print (b_11,'=======b_11=======')
								p_euro=parsing_data_EUR(tds)
								print (p_euro)
								
								#client.send_message('@parsfirmy_bot', kol_par)
								#client.send_message('@parsfirmy_bot', p_euro)
								#client.send_message('@parsfirmy_bot', '=====================')
	

								try:
									ext_norm=''
									adres_pars=urlopen('https://www.unternehmen24.info/Firmeninformationen/Deutschland/Firma/'+adr_par)
									adresse=BeautifulSoup(adres_pars.read(), 'lxml')
									adresse.findAll('table')
									tr_adresse=adresse.find('td','infotbltd3 lh140')
									adr=str(tr_adresse.get_text()).split('\n')[0]
									adr_index=str(tr_adresse.get_text()).split('\r\n')[1].replace('    ','')
								except:
									adr_index=parsing_data_b11(tds)
									adr=parsing_data_b10(tds,soup,b_9)
									print ('pediki nemcy')
									ext_norm='adres_index_ne norm===='
									
									
								print (adr)
								#print (a_18,a_19,adr,adr_index,a_17, kol_vo_, kol_par, b_9, ext_norm, adr_par )
								sleep(30)
								client.send_message('@parsfirmy_bot', b_9+'\n'+kol_par + '\n' + p_euro+ '\n'+ adr + '\n'+ adr_index + '\n'+ 'https://www.unternehmen24.info/Firmeninformationen/Deutschland/Firma/'+adr_par )
	
	
								#====?????????? ? ??????=====
	
	
	
								save_to_xlsx (a_18,a_19,adr,adr_index,a_17, kol_vo_, kol_par, b_9, ext_norm )
							
			
		else:
			print ('4to to poshlo ne tak na4inaem s na4ala')
			continue
			
			
			# ??????? ?? ????
		   
	
			






main()
