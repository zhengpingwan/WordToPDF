import sys
import os
import comtypes.client

wdFormatPDF = 17

def convertFile(file):
	in_file = os.path.abspath(file)
	out_file = os.path.abspath(file.replace(".docx", ".pdf").replace(".doc", ".pdf"))
	word = comtypes.client.CreateObject('Word.Application')
	doc = word.Documents.Open(in_file)
	doc.SaveAs(out_file, FileFormat=wdFormatPDF)
	doc.Close()
	word.Quit()


# minArgsLenght is the value that we use to find out if the user introduced file paths or not, 
# if not, we convert all ".doc" and ".docx" in the current directory.

minArgsLenght = 1

if(len(sys.argv)) > minArgsLenght:
	for file in sys.argv:
		if file.endswith(".doc") or file.endswith(".docx"):
			convertFile(file)

else:
	for file in os.listdir("."):
		if file.endswith(".doc") or file.endswith(".docx"):
			convertFile(file)
	

