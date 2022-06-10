'*********************************************************
' Objetive: Find duplicated rows in a CSV file
'    @Date: 09/Jun/2022
'  @Author: edcruces99@gmail.com
'
'*********************************************************

'Create Object File
Set FSO = CreateObject("Scripting.FileSystemObject")

'Create Array to check duplicated values
Set arrData =  CreateObject("System.Collections.ArrayList")

'Read First Argument
Filename = WScript.Arguments.Item(0)

'Read First Argument (without extension)
arrFilename=Split(Filename,".")
for each x in arrFilename
    FilenameJustName=x
	Exit For
next

'Result File
Set ff = FSO.OpenTextFile(FilenameJustName & "_dup.txt" ,2 , True)

'Open File (2 cursors)
Set f1 = fso.OpenTextFile(filename)
Set f2 = fso.OpenTextFile(filename)

'Counters
cCtrl = 1
cPpal = 1
arrData.Add "00"

Do
   lineaCtrl = f1.ReadLine
   Do
     lineaPpal = f2.ReadLine
     'msgbox cCtrl & "|" & cPpal & "|" & lineaCtrl & "|" & lineaPpal
     If StrComp(lineaPpal,lineaCtrl,1) = 0 AND cCtrl <> cPpal Then 'Compare if lines = equal and Is Not the same line
  	      ff.WriteLine cCtrl & "|" & cPpal & "|" & lineaCtrl & "|" & lineaPpal
     End If
     cPpal = cPpal + 1
   Loop Until f2.AtEndOfStream = true
   
   f2.Close
   Set f2 = fso.OpenTextFile(filename)
   cPpal=1
   cCtrl=cCtrl+1

Loop Until f1.AtEndOfStream = true

f1.Close
f2.Close
ff.Close
