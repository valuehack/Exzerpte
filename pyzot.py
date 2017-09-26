#!/usr/bin/env python

'''
Greift auf exzerpierte Texte in Zotero zu unf generiert PDF und DOCX aus den Exzerpten.
'''

import os, pypandoc, progressbar
from pyzotero import zotero
from slugify import slugify


base_dir = os.path.dirname(os.path.abspath(__file__))
paths = {
	'tmp_dir': os.path.join(base_dir,"tmp"),
	'output_dir': os.path.join(base_dir,"Output"), # Statdessen Google Drive implementieren
}

# Pfade erstellen
for path in paths:
	if not os.path.exists(paths[path]):
		os.makedirs(paths[path])

# Import Zotero API Key
with open('zotero.key', 'r') as f:
    lines = f.readlines()
    consumer_key = lines[0].strip()

zot = zotero.Zotero('4111725', 'user', consumer_key)
items = zot.collection_items('86IBK7ME', itemType='-note')
items = [item for item in items if item['data']['itemType'] != 'attachment'] # Pyzotero unterstützt keine doppelten Searchqueries. (-note&-attachment)

# Exzerpierte Dokumente durchlaufen
bar = progressbar.ProgressBar()
for item in bar(items):
	title = "%s-%s" % (slugify(item['meta']['creatorSummary']), slugify(item['data']['title']))
	html_path = os.path.join(paths['tmp_dir'],'%s.html' % title)

	with open(html_path, 'w+') as f:
		# Get Exzerpt Metadata > HTML-Datei
		citation = zot.item(item['data']['key'], content='citation', style='chicago-fullnote-bibliography')
		print('<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head>\n<body>\n\n<h1>',citation[0],'</h1>\n', file=f) # encode('utf-8')?

		# Get Exzerpte > HTML-Datei
		children = zot.children(item['data']['key'], itemType='note')
		excerpts = [child['data']['note'].replace('<strong>Yellow Annotations</strong>','') for child in children if child['data']['note'].startswith('<p><strong>Yellow')]
		for excerpt in excerpts:
			print(excerpt,'</body></html>', file=f) # encode('utf-8')?

	# PDF erzeugen
	output = pypandoc.convert_file(html_path, 'latex', outputfile=os.path.join(paths['output_dir'],('%s.pdf' % title)), extra_args=['--latex-engine=xelatex'])
	# print('%s.pdf erzeugt' % title)

	# DOCX erzeugen
	output = pypandoc.convert_file(html_path, 'docx', outputfile=os.path.join(paths['output_dir'],('%s.docx' % title)))
	# print('%s.docx erzeugt\ln' % title)

	# # MOBI erzeugen
	# call(['kindlegen', '-c2', '-o','%s.mobi' % title, html_path])
	# print('%s.mobi erzeugt' % title)

# tmp aufräumen
for f in os.listdir(paths['tmp_dir']):
	os.remove(os.path.join(paths['tmp_dir'],f))

print('Alles erledigt.')
