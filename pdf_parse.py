#!/usr/bin/env python

__author__ = 'yun.ding'
__description__ = 'Parse PDF file'

import os
import re
import sys
import argparse
import chardet
from binascii import b2a_hex
from collections import defaultdict
from pdfminer.pdfparser import  PDFParser
from pdfminer.pdfdocument import PDFDocument, PDFNoOutlines
from pdfminer.pdfpage import PDFPage, PDFTextExtractionNotAllowed
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import *

reload(sys)
sys.setdefaultencoding('utf-8')

def with_pdf(pdf_doc, pdf_pwd, fn, *args):
    # Open the pdf document, and apply the function, returning the results
    result = None
    try:
        fp = open(pdf_doc, 'rb')
        # create a parser object associated with the file object
        parser = PDFParser(fp)
        # create a PDFDocument object that stores the document structure
        # supply the password for initialization
        doc = PDFDocument(parser)
        if doc.is_extractable:
            # apply the function and return the result
            result = fn(doc, *args)
        fp.close()
    except IOError:
        # the file doesn't exist or similar problem
        pass
    return result

def _parse_toc(doc):
    # With an open PDFDocument object, get the table of contents(toc) data
    # this is a higher-order function to be passed to with_pdf()
    toc = []
    try:
        outlines = doc.get_outlines()
    except PDFNoOutlines:
        pass
    try:
        for (level,title,dest,a,se) in outlines:
            toc.append( (level, title) )
    except Exception as e:
       #print 'Exception of Outlines'
        pass
    return toc

def get_toc(pdf_doc, pdf_pwd=''):
    # Return the table of contents(toc), if any, for this pdf file
    return with_pdf(pdf_doc, pdf_pwd, _parse_toc)

def _parse_pages(doc, images_folder):
    # With an open PDFDocument object, get the pages and parse each one
    # this is a higher-order function to be passed to with_pdf()
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    text_content = [] # a list of strings, each representing text collected from each page of the doc
    char_content = []
    for i, page in enumerate(PDFPage.create_pages(doc)):
        interpreter.process_page(page)
        # receive the LTPage object for this page
        layout = device.get_result()
        # layout is an LTPage object which may contain child objects like LTTextBox, LTFigure, LTImage, etc.
        text_content.append(parse_lt_objs(layout, (i+1), images_folder)[0])
        char_content.append(parse_lt_objs(layout, (i+1), images_folder)[1])
    return text_content, char_content

def get_pages(pdf_doc, images_folder, pdf_pwd=''):
    # Process each of the pages in this pdf file and print the entire text to stdout
   #print '\n\n'.join(with_pdf(pdf_doc, pdf_pwd, _parse_pages, *tuple([images_folder])))
    return with_pdf(pdf_doc, pdf_pwd, _parse_pages, *tuple([images_folder]))

def parse_lt_objs(lt_objs, page_number, images_folder, text=[], char=[]):
    # Iterate through the list of LT* objects and capture the text or image data contained in each
    text_content = []
    char_content = []
    page_text = {} # k=(x0, x1) of the bbox, v=list of text strings within that bbox width (physical column)
    page_char = {}
    for lt_obj in lt_objs:
        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine) or isinstance(lt_obj, LTTextBoxHorizontal) or isinstance(lt_obj, LTTextBoxVertical) or isinstance(lt_obj, LTTextLineHorizontal) or isinstance(lt_obj, LTTextLineVertical) or isinstance(lt_obj, LTAnno):
            # text, so arrange is logically based on its column width
            page_text = update_page_text_hash(page_text, lt_obj)
        elif isinstance(lt_obj, LTText):
            page_char = update_page_char_hash(page_char, lt_obj)
        elif isinstance(lt_obj, LTImage):
            # an image, so save it to the designated folder, and note it's place in the text
            saved_file = save_image(lt_obj, page_number, images_folder)
            if saved_file:
                # use html style <img /> tag to mark the position of the image within the text
                text_content.append('<img src="'+os.path.join(images_folder, saved_file)+'" />')
            else:
               #print >> sys.stderr, "Error saving image on page", page_number, lt_obj.__repr__
	        pass
        elif isinstance(lt_obj, LTFigure):
            # LTFigure objects are containers for other LT* objects, so recurse through the children
            text_content.append(parse_lt_objs(lt_obj, page_number, images_folder, text_content, char_content)[0])
            char_content.append(parse_lt_objs(lt_obj, page_number, images_folder, text_content, char_content)[1])
    for k, v in sorted([(key,value) for (key,value) in page_text.items()]):
        # sort the page_text hash by the keys (x0,x1 values of the bbox), which produces a top-down, left-to-right sequence of related columns
        vTemp = ''
        for i in v:
	    iTemp = re.sub("\s{2,}"," ",i)
            vTemp += re.sub(r"\s*\n"," ",iTemp)+'\n\n'
        text_content.append(vTemp)
    for k, v in sorted([(key,value) for (key,value) in page_char.items()]):
        char_content.append(''.join(v))
    return '\n'.join(text_content), '\n'.join(char_content)

def to_bytestring(s, enc='utf-8'):
    # Convert the given unicode string to a bytestring, using the standard encoding, unless it's already a bytestring
    LATIN_1_CHARS = {
	'\xef\xac\x80': 'ff',
	'\xef\xac\x81': 'fi',
	'\xef\xac\x82': 'fl',
	'\xef\xac\x83': 'ffi',
	'\xef\xac\x84': 'ffl',
	'\xef\xac\x86': 'st',
        '\xc2\x84': '',
        '\xc3\xa9': 'e',
        '\xe2\x80\x99': "'",
        '\xe2\x80\x90': '-',
        '\xe2\x80\x91': '-',
        '\xe2\x80\x92': '-',
        '\xe2\x80\x93': '-',
        '\xe2\x80\x94': '-',
        '\xe2\x80\x94': '-',
        '\xe2\x80\x98': "'",
        '\xe2\x80\x9b': "'",
        '\xe2\x80\x9c': '"',
        '\xe2\x80\x9c': '"',
        '\xe2\x80\x9d': '"',
        '\xe2\x80\x9e': '"',
        '\xe2\x80\x9f': '"',
        '\xe2\x80\xa6': '...',
        '\xe2\x80\xb2': "'",
        '\xe2\x80\xb3': "'",
        '\xe2\x80\xb4': "'",
        '\xe2\x80\xb5': "'",
        '\xe2\x80\xb6': "'",
        '\xe2\x80\xb7': "'",
        '\xe2\x81\xba': "+",
        '\xe2\x81\xbb': "-",
        '\xe2\x81\xbc': "=",
        '\xe2\x81\xbd': "(",
        '\xe2\x81\xbe': ")" }
    if s:
        if isinstance(s, str):
            for k, v in LATIN_1_CHARS.items():
                if k in s:
                    s = s.replace(k, v)
            return s
        else:
            s = s.encode(enc)
            for k, v in LATIN_1_CHARS.items():
                if k in s:
                    s = s.replace(k, v)
            return s #s.encode(enc)

def update_page_text_hash(h, lt_obj, pct=0.2):
    # Use the bbox x0,x1 values within pct% to produce lists of associated text within the hash
    x0 = lt_obj.bbox[0]
    x1 = lt_obj.bbox[2]
    key_found = False
    for k, v in h.items():
        hash_x0 = k[0]
        if x0 >= (hash_x0 * (1.0-pct)) and (hash_x0 * (1.0+pct)) >= x0:
            hash_x1 = k[1]
            if x1 >= (hash_x1 * (1.0-pct)) and (hash_x1 * (1.0+pct)) >= x1:
                # the text inside this LT* object was positioned at the same width as a prior series of text, so it belongs together
                key_found = True
                v.append(to_bytestring(lt_obj.get_text()))
                #v.append(lt_obj.get_text())
                h[k] = v
    if not key_found:
        # the text, based on width, is a new series, so it gets its own series (entry in the hash)
        h[(x0,x1)] = [to_bytestring(lt_obj.get_text())]
        #h[(x0,x1)] = [lt_obj.get_text()]
    return h

def update_page_char_hash(h, lt_obj, pct=0.2):
    # Use the bbox x0,x1 values within pct% to produce lists of associated char within the hash
   #print lt_obj.bbox
   #print lt_obj.get_text()
    y0 = lt_obj.bbox[1]
    y1 = lt_obj.bbox[3]
    key_found = False
    for k, v in h.items():
        hash_y0 = k[0]
        if y0 >= (hash_y0 * (1.0-pct)) and (hash_y0 * (1.0+pct)) >= y0:
           hash_y1 = k[1]
           if y1 >= (hash_y1 * (1.0-pct)) and (hash_y1 * (1.0+pct)) >= y1:
               key_found = True
               v.append(to_bytestring(lt_obj.get_text()))
               #v.append(lt_obj.get_text())
               h[k] = v
    if not key_found:
        h[(y0,y1)] = [to_bytestring(lt_obj.get_text())]
        #h[(y0,y1)] = [lt_obj.get_text()]
    return h

def save_text(text, folder):
    # Try to save each page text from LTTextBox or LTTextLine object, and return the file name, if successful
    result = None
    if text:
        file_name = pdf_name + '.docx'
        if write_file(folder, file_name, text, flags='wb'):
            result = file_name
    return result

def save_image(lt_image, page_number, images_folder):
    # Try to save the image data from this LTImage object, and return the file name, if successful
    result = None
    if lt_image.stream:
        file_stream = lt_image.stream.get_rawdata()
        file_ext = determine_image_type(file_stream[0:4])
        if file_ext:
            file_name = ''.join(['page', str(page_number), '_', lt_image.name, file_ext])
            if write_file(images_folder, file_name, file_stream, flags='wb'):
                result = file_name
    return result

def determine_image_type(stream_first_4_bytes):
    # Find out the image file type based on the magic number comparison of the first 4 (or 2) bytes
    file_type = None
    bytes_as_hex = b2a_hex(stream_first_4_bytes)
    if bytes_as_hex.startswith('ffd8'):
        file_type = '.jpeg'
    elif bytes_as_hex == '89504e47':
        file_type = '.png'
    elif bytes_as_hex == '47494638':
        file_type = '.gif'
    elif bytes_as_hex.startswith('424d'):
        file_type = '.bmp'
    elif bytes_as_hex.startswith('4949'):
        file_type = '.tiff'
    return file_type

def write_file (folder, filename, filedata, flags='w'):
    # Write the file data to the folder and filename combination(flags: 'w' for write text, 'wb' for write binary, use 'a' instead of 'w' for append)
    result = False
    if os.path.isdir(folder):
        try:
            file_obj = open(os.path.join(folder, filename), flags)
            file_obj.write(filedata)
            file_obj.close()
            result = True
        except IOError:
            pass
    return result

def tree():
    return defaultdict(tree)

Technique = tree()
with open('/mnt/ilustre/users/yun.ding/newmdt/scripts/pdf_parse/info/Sequencing_Tech.xls','r') as f:
    for i in f.readlines():
        Technique[i.strip()] = 0

Sample = tree()
with open('/mnt/ilustre/users/yun.ding/newmdt/scripts/pdf_parse/info/Sample_Type.xls','r') as f:
    for i in f.readlines():
        Sample[i.strip()] = 0

parser = argparse.ArgumentParser(description = 'This script is for PDF parsing.', epilog = 'If any problem occurs, please contact yun.ding@majorbio.com.')
parser.add_argument('-i', '--input', required=True, help='input PDF file')
parser.add_argument('-o', '--output', default='./', help='output path, default: ./')
args = parser.parse_args()

pdf = args.input
pdf_name = os.path.abspath(pdf).split('/')[-1]
outPath = args.output

# Keywords need to find
keywords = tree()
keywords['Title'] = ''
keywords['Journal'] = ''
keywords['Publication date'] = ''
keywords['DOI'] = ''
keywords['PMID'] = ''
keywords['Link'] = ''
keywords['Abstract'] = ''
keywords['First author'] = ''
keywords['Address_FAU'] = ''
keywords['Technique'] = ''
keywords['Sample'] =''
Title = ''
Journal = ''
Date = ''
DOI = ''
Link = ''
author = ''
techDict = tree()
sampleDict = tree()

titleTemp = ''
content = get_toc(pdf)
for (level,title) in content:
    if level == 1:
        #titleTemp = title.encode('utf8')
        titleTemp = to_bytestring(title)
        keywords['Title'] = titleTemp

# Regulation expression
journal1 = re.compile(r'journal homepage.*\/([^\/\n]+)')
journal2 = re.compile(r'j o u r n a l h o m e p a g e.*\/([^\/\n]+)')
online = re.compile(r'online\S*\s*(\d{1,2}\s\w{3,9}\s\d{4})')
accepted1 = re.compile(r'Published:*\s*(\d{1,2}\s\w{3,9}\s\d{4})')
accepted2 = re.compile(r'Accepted\S*\s*(\d{1,2}\s\w{3,9}\s\d{4})')
accepted3 = re.compile(r'Accepted date:\S*\s*(\d{1,2}\s\w{3,9}\s\d{4})')
accepted4 = re.compile(r'Accepted\s(.*\d{1,2}.*\s\d{4})')
doi1 = re.compile(r'(doi.org/|doi:|DOI:)\s*(\S+\w)')
#doi2 = re.compile(r'doi:(\S+\w)')
#doi3 = re.compile(r'DOI:(\S+\w)')
link1 = re.compile(r'(https://\S+\w)')
link2 = re.compile(r'(http://\S+\w)')
abstract1 = re.compile(r'(.*)\n+(A B S T R A C T|a b s t r a c t)')
abstract2 = re.compile(r'(A B S T R A C T|a b s t r a c t).*\n+(.*)')
abstract3 = re.compile(r'Abstract.*\n*.*')

#test = open('ttt.xls','w')
# Match
pages_text, pages_char = get_pages(pdf, images_folder=outPath)
for i, text in enumerate(pages_text):
    for tech in Technique.keys():
        if tech in text:
            techDict[re.search(tech,text).group()] = 0
    for sam in Sample.keys():
        if sam in text:
            sampleDict[re.search(sam,text).group()] = 0
    if i < 4:
        #test.write(text+'\n')
        if titleTemp:
            if titleTemp.endswith('...') and Title != titleTemp:
                titleTemp = titleTemp.replace('...','').strip()
                #test.write(titleTemp+'\n')
                if re.search(titleTemp[:60],text):
                    Title = re.search(titleTemp[:60]+'.*',text).group()
                    keywords['Title'] = Title
                elif re.search(titleTemp[:20],text):
                    Title = re.search(titleTemp[:20]+'.*',text).group()
                    keywords['Title'] = Title
            if re.search(titleTemp[:40],text):
		temp = re.search(titleTemp[:40]+".*\n+(.*)",text).group(1)
		author = temp.split(',')[0]
		keywords['First author'] = author
                #temp = text.split('\n')
                #for i, line in enumerate(temp):
                #    if re.search(titleTemp[:-10],line):
                #        author = temp[i+2].split(',')[0]
                #        keywords['First author'] = author
                #        break
	if journal1.search(text):
	    Journal = journal1.search(text).group(1)
	    keywords['Journal'] = Journal
	elif journal2.search(text):
	    Journal = journal2.search(text).group(1)
	    keywords['Journal'] = Journal
        if online.search(text):
            Date = online.search(text).group(1)
            keywords['Publication date'] = Date
        elif accepted1.search(text):
            Date = accepted1.search(text).group(1)
            keywords['Publication date'] = Date
        elif accepted2.search(text):
            Date = accepted2.search(text).group(1)
            keywords['Publication date'] = Date
        elif accepted3.search(text):
            Date = accepted3.search(text).group(1)
            keywords['Publication date'] = Date
        elif accepted4.search(text):
            Date = accepted4.search(text).group(1)
            keywords['Publication date'] = Date
        if doi1.search(text):
            DOI = doi1.search(text).group(2)
            keywords['DOI'] = DOI
       #elif doi2.search(text):
       #    DOI = doi2.search(text).group(1)
       #    keywords['DOI'] = DOI
       #elif doi3.search(text):
       #    DOI = doi3.search(text).group(1)
       #    keywords['DOI'] = DOI
        if link1.search(text):
            Link = link1.search(text).group(1)
            keywords['Link'] = Link
        if link2.search(text):
            Link = link2.search(text).group(1)
            keywords['Link'] = Link
        elif DOI:
            Link = 'https://doi.org/'+DOI
            keywords['Link'] = Link
	if abstract1.search(text):
	    Abstract = abstract1.search(text).group(1)
	    if len(Abstract) > 300:
		keywords['Abstract'] = Abstract.replace("\n","")
        if not keywords['Abstract'] and abstract2.search(text):
	    Abstract = abstract2.search(text).group(2)
	    if len(Abstract) > 300:
	        keywords['Abstract'] = Abstract.replace("\n","")
        if not keywords['Abstract'] and abstract3.search(text):
	    Abstract = abstract3.search(text).group()
	    if len(Abstract) > 300:
	        keywords['Abstract'] = Abstract.replace("\n","")
keywords['Technique'] = '|'.join(techDict.keys())
keywords['Sample'] = '|'.join(sampleDict.keys())
for i, text in enumerate(pages_char):
    if i < 4:
        # char
        if online.search(text):
            Date = online.search(text).group(1)
            keywords['Publication date'] = Date
        elif accepted1.search(text):
            Date = accepted1.search(text).group(1)
            keywords['Publication date'] = Date
        elif accepted2.search(text):
            Date = accepted2.search(text).group(1)
            keywords['Publication date'] = Date
        elif accepted3.search(text):
            Date = accepted3.search(text).group(1)
            keywords['Publication date'] = Date
        elif accepted4.search(text):
            Date = accepted4.search(text).group(1)
            keywords['Publication date'] = Date
        if doi1.search(text):
            DOI = doi1.search(text).group(2)
            keywords['DOI'] = DOI
       #elif doi2.search(text):
       #    DOI = doi2.search(text).group(1)
       #    keywords['DOI'] = DOI
       #elif doi3.search(text):
       #    DOI = doi3.search(text).group(1)
       #    keywords['DOI'] = DOI
        if link1.search(text):
            Link = link1.search(text).group(1)
            keywords['Link'] = Link
        if link2.search(text):
            Link = link2.search(text).group(1)
            keywords['Link'] = Link
        elif DOI:
            Link = 'https://doi.org/'+DOI
            keywords['Link'] = Link

Header = ['Title','Journal','Publication date','DOI','PMID','Link','Abstract','First author','Address_FAU','Technique','Sample']
temp = list()
for i in Header:
    print i, ' > ', keywords[i]
    temp.append(keywords[i])
#print " > ".join(Header)
#print " > ".join(temp)
print '\n'
out = open('pdf_info.xls','w')
out.write('\t'.join(Header)+'\n')
out.write('\t'.join(temp))
