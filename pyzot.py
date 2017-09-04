from pyzotero import zotero
import pypandoc
import os
from subprocess import call
import re
import unicodedata

# PDFs (Exzerpte) im Verzeichnis loeschen, damit nur aktuelle Exzerpte vorhanden
filelist = [ file for file in os.listdir(".") if file.endswith(".pdf") or file.endswith(".mobi") or file.endswith(".docx") or file.endswith(".html") ]
# for file in filelist:
#	os.remove(file)
	
with open('zotero.key', 'r') as f:
    lines = f.readlines()
    consumer_key = lines[0].strip()

zot = zotero.Zotero('4111725', 'user', consumer_key)
items = zot.everything(zot.collection_items('86IBK7ME'))

for item in items:
	if item['data']['itemType'] != 'note' and item['data']['itemType'] != 'attachment':
		zot.add_parameters(content='citation', style='chicago-fullnote-bibliography')
		cites = zot.item(item['data']['key'])
		f=open(item['data']['key']+'.html', 'w+')
		print('<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head>\n<body>\n', file=f)
		print('\n<h1>',cites[0].encode('utf-8'),'</h1>\n', file=f) 
		# Titelzeite mit Referenz
		children=zot.children(item['data']['key'])
		for child in children:
			if child['data']['itemType'] == 'note' and child['data']['note'].startswith('<p><strong>Yellow'):
				c=child['data']['note'][child['data']['note'].find('\n')+1:]
				# removes first line (Yellow Annotations ...)
				print(c.encode('utf-8'), file=f) 
				print('</body></html>', file=f)
				f.close()
				# Dateiname
				def slugify(value):
					value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
					value = re.sub(r'[^\w\s-]', '', value).strip().lower()
					return re.sub(r'[-\s]+', '-', value)
				titel=slugify(item['data']['title'])
				datei=item['meta']['creatorSummary'].replace(' ', '_')+'_'+titel
				# PDF erzeugen
				output = pypandoc.convert_file(item['data']['key']+'.html', 'latex', outputfile=datei+'.pdf',extra_args=['--latex-engine=xelatex'])
				print(datei+'.pdf erzeugt')
				# DOCX erzeugen
				output = pypandoc.convert_file(item['data']['key']+'.html', 'docx', outputfile=datei+'.docx')
				print(datei+'.docx erzeugt')
#				#MOBI erzeugen
#				call(['kindlegen', '-c2', '-o','%s.mobi' %datei, '%s.html' % item['data']['key']])
#				print datei+'.mobi erzeugt'
