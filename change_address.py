from docx import Document
import csv
import os
from time import perf_counter
t1_start = perf_counter() 

def readCSV():
	with open ('Addresses.csv') as csv_file:
		read = csv.DictReader(csv_file)
		array = []
		for r in read:
			array.append(r)
	return array
array = readCSV()
file = Document('Template - Talulla IL')
paragraph = file.paragraphs
iteration = len(array)
for i in range(iteration):
	AddressLine1 = array[i]['AddressLine1']
	AddressLine2 = array[i]['AddressLine2']
	print(f"{AddressLine1}\n{AddressLine2}")
	paragraph[2].runs[0].text = AddressLine1
	paragraph[3].runs[0].text = AddressLine2
	fname = f"{AddressLine1}.docx"
	doc = file.save(fname)

os.system(f"mv *.docx DOCX/")
#os.system(f"lowriter --headless --convert-to pdf *.docx")
#os.system(f"mv *.pdf /home/tarek/MY_PROJECTS/Python_Projects/Python-Docx/PDF/")

t1_stop = perf_counter()
  
print("Elapsed time:", t1_stop, t1_start) 
print("Elapsed time during the whole program in seconds: \n", t1_stop-t1_start)

