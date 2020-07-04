# csv2vcf : CSV <-> Contacts(Android) convertor
# usage: csv2vdf.exeを起動し、csvかvcfかを選択した後、sourceとdestinationのfileを指定する。

import tkinter.filedialog as fd
import re
import quopri
import csv


termGmail = [
  'Family Name', #0
  'Given Name',#1
  'Family Name Yomi', #2
  'Given Name Yomi', #3
  'E-mail 1 - Type', #4
  'E-mail 1 - Value', #5
  'E-mail 2 - Type', #6
  'E-mail 2 - Value', #7
  'E-mail 3 - Type', #8
  'E-mail 3 - Value', #9
  'E-mail 4 - Type', #10
  'E-mail 4 - Value', #11
  'E-mail 5 - Type', #12
  'E-mail 5 - Value', #13
  'E-mail 6 - Type', #14
  'E-mail 6 - Value', #15
  'E-mail 7 - Type', #16
  'E-mail 7 - Value', #17
  'E-mail 8 - Type', #18
  'E-mail 8 - Value', #19
  'E-mail 9 - Type', #20
  'E-mail 9 - Value', #21
  'Phone 1 - Type', #22
  'Phone 1 - Value', #23
  'Phone 2 - Type', #24
  'Phone 2 - Value', #25
  'Phone 3 - Type', #26
  'Phone 3 - Value', #27
  'Phone 4 - Type', #28
  'Phone 4 - Value', #29
  'Phone 5 - Type', #30
  'Phone 5 - Value', #31
  'Phone 6 - Type', #32
  'Phone 6 - Value', #33
  'Phone 7 - Type', #34
  'Phone 7 - Value', #35
  'Phone 8 - Type', #36
  'Phone 8 - Value', #37
  'Phone 9 - Type', #38
  'Phone 9 - Value', #39
  'Address 1 - Type', #40
  'Address 1 - Value', #41
  'Address 2 - Type', #42
  'Address 2 - Value', #43
  'Address 3 - Type', #44
  'Address 3 - Value'] #45

inFilename = fd.askopenfilename(filetypes = [('input vcf file','*.vcf'), ('input csv file','*.csv')])
fCsv2Vcf = inFilename.lower().endswith('csv')
#inFile = open(inFilename, mode='r', encoding='shift-jis' if fCsv2Vcf else 'utf-8')
inFile = open(inFilename, mode='r', encoding='utf_8_sig')

outFilename = fd.asksaveasfilename(filetypes = [('vcf file','*.vcf')] if fCsv2Vcf else [('csv file','*.csv')]) 
#file = open(outFilename, mode='w', encoding='utf-8' if fCsv2Vcf else 'shift-jis')
#file = open(outFilename, mode='w', encoding='utf_8_sig')
file = open(outFilename, mode='w', encoding='utf-8' if fCsv2Vcf else 'utf_8_sig')


if fCsv2Vcf == False:
  # Vcf to Csv

  lines = inFile.readlines()

  for i in range(len(lines)):
    if re.search(r';?ENCODING=QUOTED-PRINTABLE', lines[i]): # QUOTED-PRINTABLE to UTF-8
      if re.search(r'=[\r\n]*$', lines[i]): # Continuation line
        lines[i] = re.sub(r'=[\r\n]*$', '', lines[i]) + lines[i+1] # remove '=' at tail, then concatinate
        lines[i+1] = ""
      lines[i] = re.sub(r';?ENCODING=QUOTED-PRINTABLE', '', lines[i])
      lines[i] = re.sub(r'(=[0-9a-fA-F]{2}){1,}', lambda m: quopri.decodestring(m.group()).decode(encoding="utf-8"), lines[i])
    if re.search(r'^(TEL|EMAIL|ADR)', lines[i]):
      lines[i] = re.sub(r';WORK', ';Work', lines[i])
      lines[i] = re.sub(r';HOME', ';Home', lines[i])
      lines[i] = re.sub(r';CELL', ';Mobile', lines[i])
    lines[i] = re.sub(r';CHARSET=UTF-8', '', lines[i])

  file.write(','.join(termGmail) + '\n')

  lines[0] = re.sub(b'\xef\xbb\xbf'.decode(), '', lines[0])
  for line in lines:
    line = line.rstrip()
    if line.startswith('BEGIN:VCARD'):
      iTel = 0
      iEmail = 0
      iAdr = 0
      lst = [''] * len(termGmail)
    elif line.startswith('N'):
      match = re.search(r'N:([^;]*);([^;]*)', line)
      lst[0:2] = match.group(1, 2)
    elif line.startswith('SOUND'):
      match = re.search(r'SOUND(;X-IRMC-N)?:([^;]*);([^;]*)', line)
      lst[2:4] = match.group(2, 3)
    elif line.startswith('X-PHONETIC-LAST-NAME'):
      match = re.search(r':([^;]+)', line)
      lst[2] = match.group(1)
    elif line.startswith('X-PHONETIC-FIRST-NAME'):
      match = re.search(r':([^;]+)', line)
      lst[3] = match.group(1)
    elif line.startswith('TEL'):
      match = re.search(r';(Work|Home|Mobile|X-Mobile).*:([- 0-9()]+)', line)
      lst[22+iTel] = match.group(1)
      lst[23+iTel] = match.group(2)
      iTel += 2
    elif line.startswith('EMAIL'):
      match = re.search(r';(Work|Home|Mobile|X-Mobile).*:(\S+@\S+)', line)
      lst[4+iEmail] = match.group(1)
      lst[5+iEmail] = match.group(2)
      iEmail += 2
    elif line.startswith('ADR'):
      match = re.search(r';(Work|Home):(.+)', line)
      lst[40+iAdr] = match.group(1)
      adrStr = ';'.join(match.group(2).split(';')[::-1])
      if ',' in match.group(2): adrStr = '"' + adrStr + '"' # ;separated text
      lst[41+iAdr] = adrStr # ;separated text
      iAdr += 2
    elif line.startswith('END:VCARD'):
      for txt in lst:
        file.write(txt + ',')
      file.write('\n')

else:
  # Csv to vCard
  charset = 'UTF-8' # 'SHIFT_JIS'
  csvs = csv.reader(inFile)
  next(csvs) # skip header
  for lst in csvs:
    file.write('BEGIN:VCARD\nVERSION:2.1\n')
    if lst[0] or lst[1]:
      file.write('N;CHARSET=' + charset + ':' + lst[0] + ';' + lst[1] + '\n')
    if lst[2] or lst[3]:
      file.write('SOUND;X-IRMC-N;CHARSET=' + charset + ':' + lst[2] + ';' + lst[3] + '\n')
    for i in range(9):
      if lst[5 + i*2]:
        file.write('EMAIL;' + lst[4 + i*2] + ':' + lst[5 + i*2] + '\n')
    for i in range(9):
      if lst[23 + i*2]:
        file.write('TEL;' + lst[22 + i*2] + ':' + lst[23 + i*2] + '\n')
    for i in range(3):
      if lst[41 + i*2]:
        file.write('ADR;' + lst[40 + i*2] + ';CHARSET=' + charset + ':' + lst[41 + i*2] + '\n')
    file.write('END:VCARD\n')
      

file.close()
inFile.close()
exit()
