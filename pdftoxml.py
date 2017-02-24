import os
import sys
import logging
import codecs

from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdftypes import resolve1
from xmp import xmp_to_dict
from dicttoxml import dicttoxml
from xml.dom.minidom import parseString
from xml.dom import minidom

# argument for path
arg_path = sys.argv[1]
# Start logging
logging.basicConfig(filename='villur.csv',level=logging.ERROR)
logging.error(" | Error Type | Path | File")


count = 0
no_metadata_count = 0
file_count = 0
pdffilecount = 0
corrupted_files = 0

def parsePDFfile(filepath):
	fp = open(filepath, 'rb')
	parser = PDFParser(fp)
	doc = PDFDocument(parser)
	parser.set_document(doc)
	doc.set_parser(parser)
	fp.close()
	return doc

def makeXML(xmlKeyword, xmlDesc, xmlCreator, xmlTitle, xmlFoldername):
	doc = minidom.Document()
	root = doc.createElement('root')
	doc.appendChild(root)

	keyword = doc.createElement('keywords')
	kText = doc.createTextNode(xmlKeyword)
	keyword.appendChild(kText)
	root.appendChild(keyword)

	desc = doc.createElement('desc')
	dText = doc.createTextNode(xmlDesc)
	desc.appendChild(dText)
	root.appendChild(desc)

	creator = doc.createElement('creator')
	cText = doc.createTextNode(xmlCreator)
	creator.appendChild(cText)
	root.appendChild(creator)

	title = doc.createElement('title')
	tText = doc.createTextNode(xmlTitle)
	title.appendChild(tText)
	root.appendChild(title)

	foldername = doc.createElement('folder')
	tFolder = doc.createTextNode(xmlFoldername)
	foldername.appendChild(tFolder)
	root.appendChild(foldername)

	xml_str = doc.toprettyxml(indent="  ")
	return xml_str




def checkMetadata(pdfdoc):
	if 'Metadata' in pdfdoc.catalog:
		return True
	return False

def checkKeywords(pdfdoc):
	if 'Keywords' in pdfdoc:
		return True
	return False

# loop through directories
for subdir, dirs, files in os.walk(arg_path):
	for file in files:
		file_count += 1
		filepath = subdir + os.sep + file
		if filepath.endswith(".pdf"):
			pdffilecount += 1
		try:
				pdfdoc = parsePDFfile(filepath)
				if checkMetadata(pdfdoc):
					metadata = resolve1(pdfdoc.catalog['Metadata']).get_data()
					dirname = subdir.split(os.path.sep)[-1]
					pdfdict = xmp_to_dict(metadata)
					dict1 = pdfdoc.info[0]
					xkeywords = None
					xdesc = None
					xcreator = None
					xtitle = None
					xfolder = None
					try:
						xkeywords = str(pdfdict['pdf']['Keywords']).replace('\r\n', ', ')
					except:
						xkeywords = ''
						pass
					try:
						xdesc = pdfdict['dc']['description']['x-default']
					except:
						xdesc = ''	
						pass
					try:
						xcreator = pdfdict['dc']['creator'][0]
					except:
						xcreator = ''	
						pass
					try:
						xtitle = pdfdict['dc']['title']['x-default']
					except:
						xtitle = ''	
						pass
					try:
						xfolder = dirname
					except:
						xfolder = ''	
						pass
						
					xmlstr = makeXML(xkeywords, xdesc, xcreator, xtitle, xfolder)
					filename = file.strip('.pdf')
					with codecs.open(subdir + os.sep + filename + ".xml", "w", "utf-8") as f:
						f.write(xmlstr)

					count += 1
					print(count, end='\r')
				else:
					no_metadata_count += 1
					logging.error(" | {} | {}\\ | {}".format('No Metadata', subdir, file))
			except Exception as e:
				corrupted_files += 1
				logging.error(" | {} | {}\\ | {}".format(e, subdir, file))
				continue		

#print('-' * 31)
print('    Total files parsed to XML = ', count)
#print('-' * 31)
print(' Total files with no metadata = ', no_metadata_count)
#print('-' * 31)
print('                  Total files = ', file_count)
#print('-' * 31)
print('              Total PDF files = ', pdffilecount)
#print('-' * 31)
print('        Total corrupted files = ', corrupted_files)
#print('-' * 31)