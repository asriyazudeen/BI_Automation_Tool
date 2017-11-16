'BI-Automation Tool
'Developer Name: Riyazudeen Abdul Subhan (asriyazudeen@gmail.com)
'Guided By: Shankar Prasad

'Version: BI-Automation-v.3.14.b -- Add a code to remove new lines/Blank lines

'Important instruciton.
'SQL File 
' -------- Script1----------
'/*Test case Tile*/
' SQL 
'Important if you want to comment somethign like --comments, put it in a new line then it will be ignored. It will create issue if it is not in a new line
  
' To print the whole content
 'Set objShell = CreateObject("Wscript.Shell") 
 'Wscript.Echo strSQL

Dim DSN,UserID,Pwd,noofRecords,strOldSTARTdt,strNewSTARTdt,strOldENDdt,strNewENDdt,strOldStg,strNewStg,strOldEdw,strNewEdw,rootFP,inputFP,outputFP,summaryFP,EmailFrom, EmailFromName,EmailTo,EmailSubject,EmailCC,SMTPServer,SMTPLogon,SMTPPassword,SMTPSSL,SMTPPort,cdoSendUsingPickup,cdoSendUsingPort,cdoAnonymous,cdoBasic,cdoNTLM,EmailBody

' Dict.Add is to declar a Variable Globally and it can be used across any functions. 

Set dict = Createobject("Scripting.Dictionary")
'dict.Add "rowCount", ""
dict.Add "fn", ""
dict.Add "xlsx", ""
'dict.Add "errNum", ""
dict.Add "errMsg", ""
dict.Add "inFn", ""
dict.Add "FileTime", ""
dict.Add "YRMO", ""
dict.Add "QRY", ""
'dict.Add "rs", ""
sumCnt=1
cnt = 1


Call Read_Input_Values()

Public Function Read_Input_Values()
           Set objExcel = CreateObject("Excel.Application")
				objExcel.Application.DisplayAlerts = False
				objExcel.Application.Visible = False
				objExcel.UserControl = True
			
		'	oFolder.path="C:\BI-Automation"
		'	oFile.Name="Global.xlsx"
			
           Set objWorkbook = objExcel.Workbooks.Open("C:\BI-Automation\Global.xlsx",0,False,5,"justforyou")
         '  Set objWorkBook = objExcel.Workbooks.open(oFolder.path & "\" &  oFile.Name,"justforyou")

			  objWorkbook.Worksheets(1).select
			  
			  DSN = objExcel.Range("$B$3").Value
			  UserID = objExcel.Range("$B$4").Value
		      Pwd= objExcel.Range("$B$5").Value
			  noofRecords = objExcel.Range("$B$8").Value
			  strOldSTARTdt = objExcel.Range("$B$11").Value
       ' if strOldSTARTd="" THEN 
       '   strOldSTARTdt="}}}}}}}}}}}}"
      '  END IF 
			  strNewSTARTdt = objExcel.Range("$B$12").Value
		     strOldENDdt= objExcel.Range("$B$15").Value
			 strNewENDdt = objExcel.Range("$B$16").Value
			  strOldStg = objExcel.Range("$B$19").Value
			 strNewStg = objExcel.Range("$B$20").Value
		     strOldEdw= objExcel.Range("$B$23").Value
			 strNewEdw = objExcel.Range("$B$24").Value
		   rootFP = objExcel.Range("$B$28").Value
		   inputFP = objExcel.Range("$B$31").Value
		   outputFP= objExcel.Range("$B$34").Value
		  summaryFP = objExcel.Range("$B$37").Value
		  EmailFrom = objExcel.Range("$B$40").Value
EmailFromName = objExcel.Range("$B$41").Value
EmailTo = objExcel.Range("$B$42").Value
EmailCC = objExcel.Range("$B$43").Value
SMTPServer = objExcel.Range("$B$44").Value
SMTPLogon = objExcel.Range("$B$45").Value
SMTPPassword = objExcel.Range("$B$46").Value
SMTPSSL = objExcel.Range("$B$47").Value
SMTPPort = objExcel.Range("$B$48").Value
cdoSendUsingPickup = objExcel.Range("$B$49").Value
cdoSendUsingPort = objExcel.Range("$B$50").Value
cdoAnonymous = objExcel.Range("$B$51").Value
cdoBasic = objExcel.Range("$B$52").Value
cdoNTLM = objExcel.Range("$B$53").Value
EmailBody = objExcel.Range("$B$54").Value
EmailSubject = objExcel.Range("$B$55").Value


'			msgbox dsn
'			msgbox strOldSTARTdt

		   objExcel.ActiveWorkbook.Close
           objExcel.quit
			Set objWorkbook = nothing
			Set  objExcel = nothing
			Set objWorksheet  = nothing
	
'To Get the input Cycle Date Num Value
If trim(strOldSTARTdt)=1 THEN 
CYCLE_DATE = InputBox("Please enter the CYCLE_DATE_NUM value in YYYYMMDD Format:")
  If CYCLE_DATE = "" Then
   msgbox "Execution Cancelled!" 
  Exit Function
  ELSEIF len(CYCLE_DATE) =>1 and len(CYCLE_DATE) <8  Then  
      Do while len(CYCLE_DATE) < 8 
        MsgBox "Incorrect! CYCLE_DATE_NUM value"
        CYCLE_DATE = InputBox("Please enter the YEAR_MONTH value in YYYYMMDD Format:")         
  Loop 
  End if
 dict.item("YRMO") =CYCLE_DATE
END If 
  
  
Call Create_Summary_File()	


End Function



Public Function Create_Summary_File()
   
'msgbox rootFP

			Set fso=createobject("Scripting.FileSystemObject")
		  
			'If the folder doenot exst then create the folder
			   If fso.FolderExists(rootFP) = false Then
					fso.CreateFolder (rootFP)
				end if
				
				If fso.FolderExists(inputFP) = false Then
					fso.CreateFolder (inputFP)
				end if
				
			           
 		'	Create new output folder is not exist & delete the files from output folder
                If fso.FolderExists(outputFP) = True Then
             	      Set objFolder = fso.GetFolder(outputFP)
					      For Each file In FSO.GetFolder(objFolder.Path).Files
					         'FSO.DeleteFile file.Path, True
					      Next
             	   Else
					fso.CreateFolder (outputFP)
					Msgbox "Please provide Input files to "&inputFP&"\"& " Folder then Click OK to Continue" 
                End If               

          
                
     
			Set fso=nothing
   
   			Set xlApp = CreateObject("Excel.Application") 
			Set xlWb = xlApp.Workbooks.Add
			Set objWorksheet = xlWb.Worksheets(1)
       xlApp.DisplayAlerts = False
       xlApp.Application.Visible = False
       xlApp.Cells(1,1).Value = "Execution Summary Report"
       
       xlApp.Range("A1:J1").MergeCells = True
       xlApp.Range("A1:J1").HorizontalAlignment = -4108
			 xlApp.Cells(3,1).Value = "INPUT File Name"
			 xlApp.Cells(3,2).Value = "OUTPUT File Name"
			 xlApp.Cells(3,3).Value = "Total No of Queries"
       xlApp.Cells(3,4).Value = "No of Queries Passed"
			 xlApp.Cells(3,5).Value = "No of Queries Failed - Mismatches"
			 xlApp.Cells(3,6).Value = "No of Queries Failed - SQL Error"
			 xlApp.Cells(3,7).Value = "Passed WorkSheets"
			 xlApp.Cells(3,8).Value = "Failed WorkSheets"
			 xlApp.Cells(3,9).Value = "Error WorkSheets"
			xlApp.Cells(3,10).Value = "File Execution Time"
			 Set objRange = objWorksheet.UsedRange
              objRange.Interior.Color = RGB(33, 89, 103)
              objRange.Font.Color = RGB(255, 255, 255)
			  objRange.Font.bold = True
			    xlWb.saveas  (rootFP&"\"&"Execution-Summary.xlsx" )
         ' xlWb.saveas  (outputFP&"\"&"Execution-Summary.xlsx" )
			 xlApp.quit
			Set xlWb = nothing
			Set xlApp = nothing

		   Call NotePadFunc() 
		   
End Function


Function NotePadFunc()

  Dim fso, openTxtFile, queryAll, oneQuery
  Dim z, Qry, x


' Folder path to read the input files
Const ForReading = 1


dim objFSO, s 
dim vqueries()

'This object to display the msgbox complete query..
Set objShell = CreateObject("Wscript.Shell") 

'Reading an input file from the folder
Set objFSO = CreateObject("Scripting.FileSystemObject")

	
Set objFolder = objFSO.GetFolder(inputFP)

'To check Input folder having atleast one file
If objFolder.Files.Count = 0 Then
	msgbox "Input Files Not Found. Please provide."
	Exit Function
End if	

For Each objFile In objFolder.Files
starttime = timer()
sumCnt=sumCnt+1

	s =replace(replace(objFile.name,".txt",""), ".sql","")
	dict.item("inFn")=objFile.name

filepath=objFile.Path


' Search and replace DB links and Date

Set objFile = objFSO.OpenTextFile(filepath, 1, True)
strText = objFile.ReadAll
objFile.Close

'strNewText=ReplaceTest(strOldStg,strNewStg,strText)
strNewTextA=ReplaceTest(strOldEdw,strNewEdw,strText)
'strNewTextB=ReplaceTest(strOldSTARTdt,strNewSTARTdt,strNewTextA)
'strNewTextC=ReplaceTest(strOldENDdt,strNewENDdt,strNewTextB)
strNewTextD=ReplaceTest("to_char ","to_char",strNewTextA)
strNewTextE=ReplaceTest("to_date ","to_date",strNewTextD)
strNewTextF=ReplaceTest("to_number ","to_number",strNewTextE)
strNewTextG=ReplaceTest("select count ","select count",strNewTextF)


Set objFile = objFSO.OpenTextFile(filepath, 2, True)
objFile.WriteLine strNewTextG 
'objFile.WriteLine strNewText1
'objFile.WriteLine strNewText2
objFile.Close
'end if 


   Set objReadFile = objFSO.OpenTextFile(filepath, ForReading)
			

				
    'Set objFile = objFSO.OpenTextFile (vtextfilepath, 1)
      vNoOfQueries = 0


'			Declarating it as Null, as it is storing its previous executing value. As a result same query is executing multiple times
			strNextLine=""
			 vtempline=""
			 
			Do Until objReadFile.AtEndOfStream
				cnt=1
			 strNextLine = objReadFile.Readline
					 If   Trim(strNextLine) <> empty And Trim(strNextLine) <> "" and left(trim(strNextLine),2) <> "--" and left(trim(strNextLine),2) <> "##" Then
					  vtempline = vtempline & strNextLine & " "
					  vtemplinesaved=false
					 elseif left(trim(strNextLine),4) = "----"  then
							  If trim(vtempline)=""  Then
							  else
							   redim preserve vqueries(vNoOfQueries )
							   vqueries(vNoOfQueries)=vtempline 
							   vtempline =""
							   vNoOfQueries =vNoOfQueries +1
							   vtemplinesaved=True
							  End If
					 End If
			Loop



		If vtemplinesaved=False Then
		redim preserve vqueries(vNoOfQueries )
		 vqueries(vNoOfQueries )=vtempline 
		 vNoOfQueries =vNoOfQueries +1
		 vtemplinesaved=True
		End If


		noQry= ubound(vqueries)+1

'Services.EndTransaction "getTime"

			Set xlApp = CreateObject("Excel.Application") 
			Set xlWb = xlApp.Workbooks.Add

'			Do Until i = 1 
'			  xlApp.Worksheets(i).Delete
'			  i = i - 1
'			Loop
'	msgbox noQry 
			If noQry>3 Then
			  noWs=((noQry)-3)
			  xlWb.Worksheets.Add NULL, xlWb.WorkSheets(3), noWs
			  xlWb.Worksheets(1).select
			  elseif noQry=2 then
			    xlWb.Worksheets(3).Delete
				elseif noQry=1 then
			    xlWb.Worksheets(3).Delete
				xlWb.Worksheets(2).Delete
          End If


				xlApp.DisplayAlerts = False
				Call fn_GetDateTimeText
				dim sGetCurrentDateTime
				sGetCurrentDateTime=""
				sGetCurrentDateTime = fn_GetDateTimeText()
               dict.item("fn")=  s&"_"&sGetCurrentDateTime
				 xlWb.saveas  (outputFP&"\" &dict.item("fn") &".xlsx" )
			
				xlApp.quit
				Set xlWb = nothing
				Set xlApp = nothing

'This is to pickup only YEAR MONTH from the Given input Date
YY_MM=mid(dict.item("YRMO"),1,6)
'YY_MM=replace(YY_MM,"30","")
'YY_MM=replace(YY_MM,"31","")
'YY_MM=replace(YY_MM,"28","")

'MsgBox YY_MM
 
  
		For i=0 to ubound(vqueries)
		 oneQuery=  vqueries(i)    
     Qry1=oneQuery
     Qry1=Replace(Qry1,"#####",dict.item("YRMO"))
     'msgbox oneQuery
    
   'Plug-IN YEAR Month Variables regadless of which month it is...
   'Wscript.echo Qry1 
    If strOldSTARTdt=1 THEN 
  
                    YM1=Replacetest("year_month", "YEAR_MONTH",Qry1)    
                    'YM1=CleanString(YM1) 'Removes Double space . WE NEED  TO SET A EXCEPTION FOR DOUBLE SPACE FOR VALID SCENARIO
                    YM1=Replacetest("year_month in", "YEAR_MONTH IN",YM1)
                    YM1=Replace(YM1,"YEAR_MONTH  IN","YEAR_MONTH IN") 'Removing double space between Month and IN 
                    YM1=Replace(YM1,"YEAR_MONTH   IN","YEAR_MONTH IN") 'Removing Triple space between Month and IN
                    YM1=Replace(YM1,"YEAR_MONTH=20","YEAR_MONTH = 20") 'Handles no space before & after =
                    YM1=Replace(YM1,"YEAR_MONTH =20","YEAR_MONTH = 20") 'Handles no space after =
                    YM1=Replace(YM1,"YEAR_MONTH= 20","YEAR_MONTH = 20") 'Handles no space before =               
                    YM1=Replace(YM1,"YEAR_MONTH  = 20","YEAR_MONTH = 20") 'Handles double space before =
                    YM1=Replace(YM1,"YEAR_MONTH =  20","YEAR_MONTH = 20") 'Handles double space after =
                   ' YM1=Replace(YM1,"YEAR_MONTH =","YEAR_MONTH = ") 'Handles no space after = 
                    'YM1=CleanString(YM1) 'Removes Double space . WE NEED  TO SET A EXCEPTION FOR DOUBLE SPACE FOR VALID SCENARIO
'                      YM1 = Replace(YM1, vbTab, " ")
   					 ' convert all CRLFs to spaces
'   					 YM1 = Replace(YM1, vbCrLf, " ")
'   					     Do While (InStr(YM1, "  "))
        ' if true, the string still contains double spaces,
        ' replace with single space
'        YM1 = Replace(YM1, "  ", "$")
'    Loop
    
                    YM1=Replace(YM1,"YEAR_MONTH = '20","YEAR_MONTH = 20")
                   ' YM1=Trim(YM1)
                  'Wscript.Echo  YM1
              If    instr(YM1,trim("YEAR_MONTH = 20")) > 0 THEN 
              		'	MsgBox "I'm jere"
              		'Finding YEAR MONTH in a Query followed by anydate                      
                      		temp=Mid(YM1,instr(YM1, "YEAR_MONTH = 20"), 19) 
                     'Wscript.Echo temp                    
                      		YM2=replace(YM1, temp, "YEAR_MONTH = Input_YM")    
                      'Wscript.Echo YM2  
                      		YM3=replace(YM2, "YEAR_MONTH = Input_YM'", "YEAR_MONTH = Input_YM")      'removing single quotes incase
                    'Wscript.Echo YM3
                    		 Qry1=replace(YM3, "Input_YM",YY_MM)    
                    ' Wscript.Echo Qry1   
                     
                     	    If  instr(Qry1,trim("YEAR_MONTH IN")) > 0 THEN 
							     ' MsgBox "I'm here"
							      temp=Mid(Qry1,instr(Qry1, "YEAR_MONTH IN")) 
							     ' msgbox temp
							      temp1=Mid(temp,1, instr(temp, ")")) 
							     ' msgbox temp1
							      temp2=replace(Qry1, temp1, "YEAR_MONTH IN (Input_YM)")     
							     ' msgbox temp2       
							      Qry1=replace(temp2, "Input_YM",YY_MM)   
							     'msgbox Qry1
						  ENd if
                     
                End If
                
                              
		    If  instr(YM1,"YEAR_MONTH IN") > 0 THEN 
			      'MsgBox "I'm here"
			      temp=Mid(YM1,instr(YM1, "YEAR_MONTH IN")) 
			     ' msgbox temp
			      temp1=Mid(temp,1, instr(temp, ")")) 
			     ' msgbox temp1
			      temp1=CleanString(temp1) 'Removes Double space
			       ' msgbox temp1
			      temp2=replace(YM1, temp1, "YEAR_MONTH IN (Input_YM)")     
			     'Wscript.Echo temp2    
			      Qry1=replace(temp2, "Input_YM",YY_MM)   
			     'msgbox Qry1
			     
					     
					        If    instr(Qry1,trim("YEAR_MONTH = 20")) > 0 THEN 
		              		'	MsgBox "I'm jere"
		              		'Finding YEAR MONTH in a Query followed by anydate                      
		                      		temp=Mid(Qry1,instr(Qry1, "YEAR_MONTH = 20"), 19) 
		                     ' Wscript.Echo temp                    
		                      		YM2=replace(Qry1, temp, "YEAR_MONTH = Input_YM")    
		                     ' Wscript.Echo YM2  
		                      		YM3=replace(YM2, "YEAR_MONTH = Input_YM'", "YEAR_MONTH = Input_YM")      'removing single quotes incase
		                    ' Wscript.Echo YM3
		                    		 Qry1=replace(YM3, "Input_YM",YY_MM)  
			     			 ENd if   
			     
			  ENd if      
                
            'Wscript.Echo Qry1
            
     END IF   
       
       ' Wscript.echo Qry1
   ' Plug-IN CYCLE DATE MONTH Variables regadless of which month it is...
    If strOldSTARTdt=1 THEN 
 ' MsgBox YM1
                    CDM=Replacetest("cycle_dt_num", "CYCLE_DT_NUM",Qry1)                 
                    CDM1=Replacetest("cycle_dt_num in", "CYCLE_DT_NUM IN",CDM)
                    CDM1=Replace(CDM1,"CYCLE_DT_NUM=20","CYCLE_DT_NUM = 20") 'Handles no space before =
                    CDM1=Replace(CDM1,"CYCLE_DT_NUM =20","CYCLE_DT_NUM = 20") 'Handles no space before =
                    CDM1=Replace(CDM1,"CYCLE_DT_NUM= 20","CYCLE_DT_NUM = 20") 'Handles no space after = 
                    CDM1=Replace(CDM1,"CYCLE_DT_NUM  = 20","CYCLE_DT_NUM = 20") 'Handles double space before =
                    CDM1=Replace(CDM1,"CYCLE_DT_NUM =  20","CYCLE_DT_NUM = 20") 'Handles double space after =
                   ' CDM1=CleanString(CDM1) 'Removes Double space
                   ' msgbox cdm1
                    CDM1=Replace(CDM1,"CYCLE_DT_NUM = '20","CYCLE_DT_NUM = 20") 'Handles with/without single quotes
                      'msgbox cdm1
                    
              If   instr(CDM1,trim("CYCLE_DT_NUM = 20")) > 0 THEN 
              'msgbox "Im here"
                      'Finding YEAR MONTH in a Query followed by anydate                                           
                      temp=Mid(CDM1,instr(CDM1, "CYCLE_DT_NUM = 20"), 23) 
                     ' msgbox temp
                      CDM2=replace(CDM1, temp, "CYCLE_DT_NUM = Input_YM")     
                      CDM3=replace(CDM2, "CYCLE_DT_NUM = Input_YM'", "CYCLE_DT_NUM = Input_YM")     'removing single quotes incase 
                    ' msgbox CDM3
                     Qry1=replace(CDM3, "Input_YM",dict.item("YRMO"))  
                     ' msgbox Qry1
                     
		                  if instr(Qry1,"CYCLE_DT_NUM IN") > 0 THEN 
		                      temp=Mid(CDM1,instr(CDM1, "CYCLE_DT_NUM IN")) 
		                      'msgbox temp
		                      temp1=Mid(temp,1, instr(temp, ")")) 
		                     ' msgbox temp1
		                      temp2=replace(Qry1, temp1, "CYCLE_DT_NUM IN (Input_YM)")     
		                     ' msgbox temp2       
		                      Qry1=replace(temp2, "Input_YM",dict.item("YRMO"))   
		                      'msgbox Qry1
		              ENd if 
                     
                ENd if 
                    
              if instr(CDM1,"CYCLE_DT_NUM IN") > 0 THEN 
                      temp=Mid(CDM1,instr(CDM1, "CYCLE_DT_NUM IN")) 
                      'msgbox temp
                      temp1=Mid(temp,1, instr(temp, ")")) 
                     ' msgbox temp1
                      temp2=replace(CDM1, temp1, "CYCLE_DT_NUM IN (Input_YM)")     
                     ' msgbox temp2       
                      Qry1=replace(temp2, "Input_YM",dict.item("YRMO"))   
                      'msgbox Qry1
                      
                          If  instr(Qry1,trim("CYCLE_DT_NUM = 20")) > 0 THEN 
		              			'msgbox "Im here"
		                      'Finding YEAR MONTH in a Query followed by anydate                                           
		                      temp=Mid(Qry1,instr(Qry1, "CYCLE_DT_NUM = 20"), 23) 
		                     ' msgbox temp
		                      CDM2=replace(Qry1, temp, "CYCLE_DT_NUM = Input_YM")     
		                      CDM3=replace(CDM2, "CYCLE_DT_NUM = Input_YM'", "CYCLE_DT_NUM = Input_YM")     'removing single quotes incase 
		                    ' msgbox CDM3
		                     Qry1=replace(CDM3, "Input_YM",dict.item("YRMO"))  
                     	 ENd if  
                      
              ENd if       
     END IF    
       
    'Removing semicolon ';' from the Query    
      Set objShell = CreateObject("Wscript.Shell") 

	' New Function added to remove the NEW line and the empty line. 
	 Qry1 = SpecialTrim(Qry1) 
	

	
    'Wscript.Echo  Qry1
     'MsgBox Qry1
     Qry1=trim(Qry1) 
	 'Qry1=StrReverse(Qry1)
	 'Qry1=LTrim(Qry1)
	 '   Wscript.Echo  Qry1  
	  'test=  instr(Qry1,";")
	 ' MsgBox test
	    
     If instr(StrReverse(Qry1),";") = 1 THEN    
    ' MsgBox "I'm here"
     Rev1=Mid(StrReverse(trim(Qry1)),2)   
     Rev2=StrReverse(Rev1)
     Qry1=Rev2 
     End If
         
    'After all above compupation assigning Qry1 to Qry   
     Qry=Qry1 
   
   ' msgbox Qry    
		dict.Item("xlsx") =i+1

				Set objExcel = CreateObject("Excel.Application")
				objExcel.Application.DisplayAlerts = False
				objExcel.Application.Visible = False
				objExcel.UserControl = True
				
				Set objWorkbook = objExcel.Workbooks.Open(outputFP&"\"& dict.item("fn")&".xlsx")
				'objExcel.Workbooks.Open(excelPath)
				Set objWorksheet = objWorkbook.Worksheets(dict.Item("xlsx"))
				
	'Setting Default excel proprty to Text for all the columns
			'	Set objRange = objExcel.Range("A","DZ")    
			'	objRange.NumberFormat = "@" 
'				Set xlRng = objExcel.ActiveSheet.Columns("A:DZ") 
'				xlRng.NumberFormat = "@"    
			'	objWorksheet.Range("A:DZ").Select.Selection.NumberFormat = "@" 
				
			 'objWorksheet.Cells(1,2)= Qry
       
      'Splitting Query & Test case title and printing it in seperate cells. 
      
               if InStr(trim(Qry),"/*") = 1 THEN  
                  PQ=mid(Qry,instr(qry,"*/")+3)    ' This prints Query           
                  PTT=mid(Qry,3, instr(qry,"*/")-3)  'This print's Test case title
                  Qry=PQ
                  objWorksheet.Cells(1,1)= PTT                 
                  objWorksheet.Cells(2,4)= PQ                  
                  objWorksheet.Cells(2,4).Font.Size = 8
                  objWorksheet.Cells(2,4).Font.Italic = True
                  objWorksheet.Cells(1,1).Font.Color = RGB(151,71,6)
                  objWorksheet.cells(1,1).Font.bold = True                                  
               ELSE 
                 objWorksheet.Cells(1,1)=" Test Case Title: ?"
                 objWorksheet.Cells(1,1).Font.Color = RGB(151,71,6)
                 objWorksheet.cells(1,1).Font.bold = True   
                 objWorksheet.Cells(2,4)= " "& Qry 
                 objWorksheet.Cells(2,4).Font.Size = 8
                 objWorksheet.Cells(2,4).Font.Italic = True                                
               END If
							 
			   objWorksheet.Columns("A").ColumnWidth = 30      
			   objWorksheet.Name="Query"&dict.Item("xlsx")
			   objExcel.ActiveWorkbook.Save
			   objExcel.ActiveWorkbook.Close
			 
				objExcel.quit
				Set objWorkbook = nothing
				Set  objExcel = nothing
				Set objWorksheet  = nothing

			 dict.item("QRY")=Qry 
			Call DataBaseFunc(DSN, UserID, Pwd, Qry)
		
		Next 
		'Call the function summary after each output file generated
		endtime=timer()

			totalTime=round(endtime-starttime)
			If totaltime>59 Then
			     dict.item("FileTime")= round((totaltime/60)) & " Min"
				else 
			    dict.item("FileTime") = totalTime & " Sec"
			End If
		
		Call Mismatch()
		Call Summary()
		
				
Next
MsgBox("Executed Successfully!") 
Call SendEmail()
End Function


Function  DataBaseFunc(DSN, UserID, Pwd, Qry)

		'Create the objects for retrieving the data
		Dim objDB

		'Declaring the Array
		Dim DBArray()
		
    'Assigning the Query
		strSQL = Qry
    
   'Set objShell = CreateObject("Wscript.Shell") 
 'Wscript.Echo strSQL
	
		'Create an object for the database connection
	    Set objDB = CreateObject("ADODB.Connection")
	  
		objDB.ConnectionTimeout = 1000
		objDB.ConnectionString = "DSN=" & DSN & ";" & "UID=" & UserID & ";" & "PWD=" & Pwd

'		  Set objCommand = CreateObject("ADODB.Command")
'		Set objCommand.ActiveConnection = objDB
'		objDB.CommandTimeout = 10


		'Open the DB Connection
		objDB.open

		'Create a recordset to hold the results
	    Set rs = CreateObject("ADODB.Recordset")
	
		'Options for CursorType are:  0=Forward Only, 1=KeySet, 2=Dynamic, 3=Static (read-only)
	    rs.CursorType = 3
       Set rs.ActiveConnection = objDB


'Open the Output excel file to write the data
 Set objExcel = CreateObject("Excel.Application")
	objExcel.Application.DisplayAlerts = False
	objExcel.Application.Visible = False
	objExcel.UserControl = True

	Set objWorkbook = objExcel.Workbooks.Open(outputFP&"\"& dict.item("fn")&".xlsx" )
	objWorkbook.Worksheets(dict.Item("xlsx")).select


'Note: The below code is used to Fetch the Only First Row. Mainly we used this during our Count validation  but  now our existing code only taking take of count. So we dont't need this  but only for reference we  have this.

'Criteria ONE for Count Validation  Note: Case sensitive is not an issue here, it issue with REPLACE only 
						'strvar="Select Count("
						'chkcma=ReplaceTest("from","FROM",strSQL)
						'chkcmaA = split(chkcma,"FROM")
						'chkcmaB=trim(chkcmaA(0))
						'
						''WScript.TimeOut = 3 
						'
						'If instr(1,strSQL,strvar,1) =1  and instr(1,chkcmaB,",",1) =0 then 
						'
						'						If instr(1,strSQL,"minus",1)=0 Then
						'						  objExcel.Cells(2,2)="One Query" 
						'						  objExcel.cells(2,2).Font.Color = RGB(50,205,50)
						'						End If
						'	
						'On error resume next
						'
						'starttime = timer()
						'rs.Open strSQL
						'endtime=timer()
						'
						'totalTime=round(endtime-starttime)
						'If totaltime>59Then
						'	totaltime= round((totaltime/60)) & " Min"
						'	else 
						'	totalTime = totalTime & " Sec"
						'End If
						'
						'
						'		if 	err.number<> 0 then
						'				errMsg= Err.Source & Err.Description
						'				 objExcel.Cells(2,1)="Error"
						'			     objExcel.cells(2,1).Font.Color = RGB(255, 0, 0)
						'				objExcel.cells(2,1).Font.bold = True
						'				objExcel.Cells(3,1)=errMsg
						'				objExcel.cells(3,1).Font.Color = RGB(255, 0, 0)
						'				objExcel.Cells(3,1).WrapText=FALSE
						'				Err.Clear
						'				else
						'				
						'				intItemCount = rs.Fields.Item(0)
						'							If  	intItemCount<>"" and  intItemCount<>0Then
						'										objExcel.Cells(2,1)="Test Case: Fail"
						'										objExcel.cells(2,1).Font.Color = RGB(255, 0, 0)
						'										objExcel.cells(2,1).Font.bold = True
						'										objExcel.Cells(3,1)="Sample Mismatch Records"
						'										objExcel.cells(3,1).Font.Color = RGB(50,205,50)
						'									   	objExcel.Cells(3,6)= "Execution Time:  " &  totalTime
						'										objExcel.cells(3,6).Font.Color = RGB(0, 0, 255)
						'										objExcel.Cells(4,1)="SOURCE DATA"
						'										objExcel.cells(4,1).Font.bold = True
						'									   objExcel.cells(4,1).Font.Color = RGB(0, 0, 255)
						'									   objExcel.Range("A5:Z5").Interior.Color =RGB(50, 205, 50)
						'                                        objExcel.Cells(5,1)="Count(*)"
						'                                        objExcel.Range("A5:Z5").Font.Color =   RGB(255, 255, 255)
						'										objExcel.Cells(6,1)=intItemCount 
						'									
						'									  elseif   intItemCount="" then
						'									 objExcel.Cells(2,1)="Pass"
						'									 objExcel.cells(2,1).Font.Color = RGB(0, 0, 255)
						'									 objExcel.cells(2,1).Font.bold = True
						'									 objExcel.Cells(3,6)= "Execution Time:  " &  totalTime
						'									objExcel.cells(3,6).Font.Color = RGB(0, 0, 255)
						'									objExcel.Cells(4,1)="DATA MATCHING!!"
						'									 objExcel.cells(4,1).Font.Color = RGB(50,205,50)
						'
						'									 else  intItemCount=0 and  instr(1,strSQL,"minus",1)=0   and instr(1,strSQL,"unoin",1)=0 
						'									 msgbox "No Data was Found  in Source Table!! Execution will STOP!!!"
						'									 ExitTest
						'								   
						'						   End If
						'                   End if     
						'				  objExcel.ActiveWorkbook.Save
						'				  objExcel.ActiveWorkbook.Close
						'				  objExcel.quit
						'				  Set objWorkbook = nothing
						'				  Set  objExcel = nothing
						'				 Set objWorksheet  = nothing
						'		 Exit function
						'end if 	

 'Set objShell = CreateObject("Wscript.Shell") 
 'Wscript.Echo strSQL


On error resume next
starttime = timer()
rs.Open strSQL
endtime=timer()

totalTime=round(endtime-starttime)
If totaltime>59Then
	totaltime=round((totaltime/60)) & " Min"
	else 
	totalTime = totalTime & " Sec"
End If

'Criteria TWO error msg Handeling
if 	err.number<> 0 then
errMsg= Err.Source & Err.Description
objExcel.Cells(2,1)="Error"
objExcel.cells(2,1).Font.Color = RGB(255, 0, 0)
objExcel.cells(2,1).Font.bold = True
objExcel.Cells(3,1)=errMsg
objExcel.cells(3,1).Font.Color = RGB(255, 0, 0)
objExcel.Cells(3,1).WrapText=FALSE
Err.Clear

		   objExcel.ActiveWorkbook.Save
		   objExcel.ActiveWorkbook.Close
           objExcel.quit
			Set objWorkbook = nothing
			Set  objExcel = nothing
			Set objWorksheet  = nothing
Exit Function 
end if


' Criteria THREE when record  NOT found
If rs.EOF = True Then
strNewText=ReplaceTest("minus","MINUS",strSQL)

		If  instr(strNewText,  "MINUS") =0 Then
			objExcel.Cells(2,2)="One Query" 
			objExcel.cells(2,2).Font.Color = RGB(255,255,255)
		End If
		
objExcel.Cells(2,1)="Test Case: Pass"
objExcel.cells(2,1).Font.Color = RGB(0, 0, 255)
objExcel.cells(2,1).Font.bold = True
objExcel.Cells(3,6)= "Execution Time:  " &  totalTime
objExcel.cells(3,6).Font.Color = RGB(0, 0, 255)
objExcel.Cells(4,1)="DATA MATCHING!!"
objExcel.Cells(4,1).Interior.Color =RGB(146, 208, 80)
'objExcel.cells(4,1).Font.Color = RGB(50,205,50)
objExcel.cells(4,1).Font.bold = True



		   objExcel.ActiveWorkbook.Save
		   objExcel.ActiveWorkbook.Close
           objExcel.quit
			Set objWorkbook = nothing
			Set  objExcel = nothing
			Set objWorksheet  = nothing
End If


' Criteria FOUR when record found
If not rs.eof then 
strNewText=ReplaceTest("minus","MINUS",strSQL)
		If  instr(strNewText,  "MINUS") =0 Then
			objExcel.Cells(2,2)="One Query" 
			 objExcel.cells(2,2).Font.Color = RGB(255,255,255)
		End If
	 objExcel.Cells(2,1)="Test Case: Fail" 
	 objExcel.cells(2,1).Font.Color = RGB(255, 0, 0)
	  objExcel.cells(2,1).Font.bold = True
	  objExcel.Cells(3,1)="Sample Mismatch Records"  
    objExcel.cells(3,1).Font.Italic = True
	  objExcel.cells(3,1).Font.Color = RGB(255,0,0)
	  objExcel.Cells(3,6)= "Execution Time:  " &  totalTime
      objExcel.cells(3,6).Font.Color = RGB(0, 0, 255)
	 objExcel.Cells(4,1)="SOURCE DATA"
	objExcel.cells(4,1).Font.bold = True
	objExcel.cells(4,1).Font.Color = RGB(0, 0, 255)
	'objExcel.Range("A5:DZ5").Interior.Color =RGB(166, 166, 166)
	'objExcel.Range("A5:Z5").Font.Color =   RGB(255, 255, 255)
			fldCount =rs.Fields.Count 
     	For iCol = 1 To fldCount 
            objExcel.Cells(5, iCol).Value =rs.Fields(iCol - 1).Name 
            objExcel.Cells(5, iCol).Interior.Color =RGB(166, 166, 166)
			Next 

			recArray =rs.GetRows(noofRecords) 
            recCount = UBound(recArray, 2) + 1
			
'			If recCount >=  100 Then
'			  objExcel.Cells(2,3)="More than 100 Mismath records Found" 
'			   recCount =  noofRecords
'			  else
'			  objExcel.Cells(2,3)=recCount
'			End If
'			noofRecords=fetchrecords
'	    	If recCount >=  noofRecords Then
'				recCount =  noofRecords
'			End If

			' Writing Column Titles in Excel
			objExcel.Cells(6, 1).Resize(recCount, fldCount).Value = TransposeDim(recArray) 
			Set xlRng = objExcel.ActiveSheet.Columns("B:DZ") 
			xlRng.AutoFit    
		 objExcel.Cells(5, 4).ColumnWidth = 30


		'Report the Failure Query 
		
		' First, create the message
		Set objMessage = CreateObject("CDO.Message")
		objMessage.Subject = EmailSubject& " TEST CASE FAILED: " & objExcel.Cells(1,1)
		objMessage.From = """" & EmailFromName & """ <" & EmailFrom & ">"
		objMessage.To = EmailTo
		objMessage.CC= EmailCC
		objMessage.TextBody = strSQL
		
		
	' Second, configure the server		
		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2		
		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer		
		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort		
		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60		
		objMessage.Configuration.Fields.Update
		
		' Now send the message!
		objMessage.Send
				 
		 

end if 


		   objExcel.ActiveWorkbook.Save
		   objExcel.ActiveWorkbook.Close
           objExcel.quit
			Set objWorkbook = nothing
			Set  objExcel = nothing
			Set objWorksheet  = nothing

	'Close the Recordset  
	rs.close
	
	'Close the connection
	objDB.Close
		
	'Destroy the objDB object
	Set objDB = Nothing





End Function


Public Function fn_GetDateTimeText

Dim sDateTime
'sDateTime=Day(Now) & "_"& Month(Now) & "_"&  Year(Now) & "_"&  Hour(Now) & "_"&  Minute(Now) & "_"&  Second(Now)
sDateTime=Month(Now) & "_"& Day(Now) & "_"&  Year(Now) & "_"& replace(dict.item("YRMO"),",","-")
fn_GetDateTimeText =sDateTime
End Function


Function Summary()

		'Opening an Output File
                Set objExcel = CreateObject("Excel.Application")
				objExcel.Application.DisplayAlerts = False
				objExcel.Application.Visible = False
				objExcel.UserControl = True
				
				Set objWorkbook = objExcel.Workbooks.Open(outputFP&"\" & dict.item("fn")&".xlsx" )
				 
				noSql=objWorkbook.Worksheets.count

				Dim pass, fail, eror, VarP, VarF
				CounterP=0
				CounterF=0
				CounterE=0
				varP=","
				varF=","
				varE=","
				For i = 1 to (noSql)
							Set objWorksheet = objWorkbook.Worksheets(i)
							NewRow = objWorksheet.Range("$A$2").Value             
'							msgbox NewRow
								
							   if cstr(trim(NewRow)) =cstr(trim("Test Case: Pass")) then 
									CounterP = CounterP +1
									pass=CounterP
									varP=replace(varP&","& objWorksheet.Name,",,","")
'									varP1=replace(varP,",,")
								  elseIf cstr(trim(NewRow)) =cstr(trim("Test Case: Fail")) then 
									 CounterF=CounterF +1
									 fail=CounterF
									 varF=replace(varF&", "& objWorksheet.Name,",,","")
								 elseIf cstr(trim(NewRow)) =cstr(trim("Error")) then 
									 CounterE=CounterE +1
									 eror=CounterE
									 varE=replace(varE&", "& objWorksheet.Name,",,","")
'									 varF1=replace(varF,",,")
							  end if
					
				Next

    			objExcel.quit
				Set objWorkbook = nothing
				Set  objExcel = nothing
				Set objWorksheet  = nothing

				' Opening an Summary File

			   Set objExcel = CreateObject("Excel.Application")
				objExcel.Application.DisplayAlerts = False
				objExcel.Application.Visible = False
				objExcel.UserControl = True
				
				Set objWorkbook = objExcel.Workbooks.Open(rootFP&"\"&"Execution-Summary.xlsx")

				Set objWorksheet = objWorkbook.Worksheets(1)
'				Set objCell = objWorkSheet.Cells.Find("
		 
        objWorksheet.Range("A1:J1").Interior.Color =RGB(146, 205, 220)
        objWorksheet.Range("A1:J1").Font.Size = 20
        objWorksheet.Range("A1:J1").Font.Color = RGB(255, 255, 255)
        objWorksheet.Range("A2:J2").Interior.Color =RGB(49, 134, 155)
        
        'msgbox sumCnt
				objWorksheet.Cells(sumCnt+2,1)=dict.item("inFn")
			'	objWorksheet.cells(sumCnt,1).Font.Color = RGB(0, 0, 255)
				objWorksheet.Cells(sumCnt+2,2)=dict.item("fn") 
        
        'HYPER Linking Syntax
        'objWorksheet.Cells(sumCnt, 2) = "=HYPERLINK(""http://www.google.com"", ""XYZ"")"
        '.ActiveSheet.Cells(1, 1) = "=HYPERLINK(""" & sLinkAddress & """,""" & sFriendly & """)"
        sLinkAddress=outputFP&"\" & dict.item("fn")&".xlsx"
        objWorksheet.Cells(sumCnt+2, 2) = "=HYPERLINK(""" & sLinkAddress & """,""" & dict.item("fn")  & """)"  
				objWorksheet.cells(sumCnt+2,2).Font.Color = RGB(0, 112, 192)
				objWorksheet.Cells(sumCnt+2,3)=noSql
			  objWorksheet.cells(sumCnt+2,3).Font.bold = True
				objWorksheet.Cells(sumCnt+2,4)=pass
				objWorksheet.cells(sumCnt+2,4).Font.bold = True
				objWorksheet.cells(sumCnt+2,4).Font.Color = RGB(0, 0, 255)
				objWorksheet.Cells(sumCnt+2,5)=fail
				objWorksheet.cells(sumCnt+2,5).Font.bold = True
				objWorksheet.cells(sumCnt+2,5).Font.Color = RGB(255, 0, 0)
				objWorksheet.Cells(sumCnt+2,6)=eror
				objWorksheet.cells(sumCnt+2,6).Font.bold = True
				objWorksheet.cells(sumCnt+2,6).Font.Color = RGB(192, 0, 0)
				objWorksheet.Cells(sumCnt+2,10)=dict.item("FileTime")
				objWorksheet.cells(sumCnt+2,10).Font.bold = True
				objWorksheet.cells(sumCnt+2,10).Font.Color = RGB(0, 0, 255)

			varPA=Split(varP,",")
			m=sumCnt+2
				For each x in varPA
				objWorksheet.Cells(m,7)=x
				objWorksheet.cells(m,7).Font.Color = RGB(0, 0, 255)
        'Syntax to connet Excel & sheet
        '=HYPERLINK("[C:\BI-Automation\Output_Files\Query_12_8_2014_.xlsx]Query2!B1", "H12")
       
        LinkSheet="["&sLinkAddress&"]"&x&"!"&"A1" 
        objWorksheet.Cells(m,7)="=HYPERLINK(""" & LinkSheet & """,""" & x  & """)" 
				m=m+1
				Next

			varFA=Split(varF,",")
				n=sumCnt+2
				For each x in varFA
                 objWorksheet.Cells(n,8)=x
				         LinkSheet="["&sLinkAddress&"]"&x&"!"&"A1" 
                 objWorksheet.Cells(n,8)="=HYPERLINK(""" & LinkSheet & """,""" & x  & """)" 
                 objWorksheet.cells(n,8).Font.Color = RGB(255, 0, 0)   
                 n=n+1
				Next


			varEA=Split(varE,",")
				o=sumCnt+2
				For each x in varEA
         objWorksheet.Cells(o,9)=x				 
         LinkSheet="["&sLinkAddress&"]"&x&"!"&"A1" 
         objWorksheet.Cells(o,9)="=HYPERLINK(""" & LinkSheet & """,""" & x  & """)" 
         objWorksheet.cells(o,9).Font.Color = RGB(192,0,0)
         o=o+1
				Next


				
sumCnt= (objWorksheet.UsedRange.Rows.Count)+1

with  objWorksheet.UsedRange.borders(7)
.LineStyle = 1
.Weight = 2
.ColorIndex = -4105 
End With

with  objWorksheet.UsedRange.borders(8)
.LineStyle = 1
.Weight = 2
.ColorIndex = -4105 
End With


with  objWorksheet.UsedRange.borders(9)
.LineStyle = 1
.Weight = 2
.ColorIndex = -4105 
End With

with  objWorksheet.UsedRange.borders(10)
.LineStyle = 1
.Weight = 2
.ColorIndex = -4105 
End With

with  objWorksheet.UsedRange.borders(11)
.LineStyle = 1
.Weight = 2
.ColorIndex = -4105 
End With


with  objWorksheet.UsedRange.borders(12)
.LineStyle = 1
.Weight = 2
.ColorIndex = -4105 
End With

			  Set objRange = objWorksheet.UsedRange
        objExcel.ActiveWindow.Zoom = 85
			   objRange.EntireColumn.Autofit()         
			   objExcel.ActiveWorkbook.Save
			   objExcel.ActiveWorkbook.Close
    			objExcel.quit
				Set objWorkbook = nothing
				Set  objExcel = nothing
				Set objWorksheet  = nothing
       
End Function


Function TransposeDim(v) 

'Dim X As Long, Y As Long, Xupper As Long, Yupper As Long 
'Dim tempArray As Variant 

Xupper = UBound(v, 2)

'If  Xupper >= noofRecords Then
'	Xupper =  noofRecords
'End If
Yupper = UBound(v, 1) 
ReDim tempArray(Xupper, Yupper) 
For X = 0 To Xupper
For Y = 0 To Yupper 
tempArray(X, Y) = v(Y, X) 
Next 
Next 

TransposeDim = tempArray 

End Function 
'This function is to remove extra spaces
Public Function CleanString(strSource) 
    'On Error GoTo CleanStringErr

    ' convert tabs to spaces first
    strSource = Replace(strSource, vbTab, " ")

    ' convert all CRLFs to spaces
    strSource = Replace(strSource, vbCrLf, " ")

    ' Find and replace any occurences of multiple spaces
    Do While (InStr(strSource, "  "))
        ' if true, the string still contains double spaces,
        ' replace with single space
        strSource = Replace(strSource, "  ", " ")
    Loop

    ' Remove any leading or training spaces and return
    ' result
    CleanString = Trim(strSource)
    Exit Function

'CleanStringErr:
    ' Insert error-handling code here
End Function
'This function is to change the string into upper or lower regardless of what it is orginally 
Public Function ReplaceTest(ptrn,toreplace,strText)
Dim regexp,str
Set regexp=New RegExp
regexp.Pattern=ptrn
regexp.IgnoreCase=True
regexp.Global=True
ReplaceTest=regexp.Replace(strText ,toreplace)
End Function



Function Mismatch()



   		'Opening an Output File
        Set objExcel = CreateObject("Excel.Application")
				objExcel.Application.DisplayAlerts = False
				objExcel.Application.Visible = False
				objExcel.UserControl = True
				
				Set objWorkbook = objExcel.Workbooks.Open(outputFP&"\" & dict.item("fn")&".xlsx" )
				 
				noSql=objWorkbook.Worksheets.count


 For i = 1 to (noSql)
		Set objWorksheet = objWorkbook.Worksheets(i)
		objWorkbook.Worksheets(i).select
		NewRow = objExcel.Range("$A$2").Value
		dupchk = objExcel.Range("$B$2").Value
		Val=" "
		Valu=" "
	   	ValA=" "
		ValuA=" "
		
				'Get the max row occupied in the excel file 
				  Row= objWorksheet.UsedRange.Rows.Count

				  TarSP=(Row)+2
				'Specify which column to be fetched
				  n=1
				  l=2
				  
				'To read the data from the first column Excel file
					For  m= 6 to Row
					Val = Val &"," & "'" & objExcel.cells(m,n).value & "'" 
					Valu= replace(Val, ",","",1,1)
                  Next

				 'To read the data from the second column Excel file
				For  k= 6 to Row
					ValA = ValA &"," & "'" & objExcel.cells(k,l).value & "'" 
					ValuA= replace(ValA, ",","",1,1)
                  Next

	  
						' First Check When SHEET is FAIL and No Null or Duplicate
						If cstr(trim(NewRow)) =cstr(trim("Test Case: Fail"))  and  cstr(trim(dupchk)) <> cstr(trim("One Query")) then 



		Set objDB = CreateObject("ADODB.Connection")
		objDB.ConnectionTimeout = 1000
		objDB.ConnectionString = "DSN=" & DSN & ";" & "UID=" & UserID & ";" & "PWD=" & Pwd
		
		'Open the DB Connection
		objDB.open

		'Create a recordset to hold the results
		Set rs = CreateObject("ADODB.Recordset")
	
		'Options for CursorType are:  0=Forward Only, 1=KeySet, 2=Dynamic, 3=Static (read-only)
		rs.CursorType = 3
		Set rs.ActiveConnection = objDB


							
							  var = objExcel.Range("$D$2").Value
							 'var=  dict.item("QRY")
              '  msgbox var
							 strNewText=ReplaceTest("minus","MINUS",var)
               sql=split(strNewText,"MINUS") 
               On error resume next
               sqlA=trim(sql(1))
          				 
						strvar="SELECT COUNT("
						chkca=ReplaceTest("from","FROM",sqlA)
						chkcaA = split(chkca,"FROM")
						chkcaB=trim(ucase(chkcaA(0)))
							 
														'Cheking IF the QUERY has select Count and  No other columns Key word
														  If instr(1,chkcaB,strvar,1) <>0 and instr(1,chkcaB,",",1) =0  then 
																			On error resume next
																			 rs.Open sqlA
																			if 	err.number<> 0 then
																					errMsg= Err.Source & Err.Description
																					objExcel.Cells(TarSP+4,1)="TARGET DATA"
																					 objExcel.Cells(TarSP+4,1).Font.bold = True
																					objExcel.Cells(TarSP+4,1).Font.Color = RGB(0, 0, 255)
																					objExcel.Cells(TarSP+5,1)=errMsg
																					objExcel.Cells(TarSP+5,1).WrapText=FALSE
																					Err.Clear
																			else
                                                                                    intItemCount = rs.Fields.Item(0)
																										If  	intItemCount<>"" Then
																												objExcel.Cells(TarSP+3,1)=sqlA
																												objExcel.Cells(TarSP+4,1)="TARGET DATA"
																												objExcel.Cells(TarSP+4,1).Font.bold = True
																												objExcel.Cells(TarSP+4,1).Font.Color = RGB(0, 0, 255)
																												objExcel.Cells(TarSP+5,1)="Count(*)"
																												objExcel.Range("A"&TarSP+5,"Z"&TarSP+5).Interior.Color = RGB(166, 166, 166)
																												'objExcel.Range("A"&TarSP+5,"Z"&TarSP+5).Font.Color =   RGB(255, 255, 255)
																												objExcel.Cells(TarSP+6,1)=intItemCount 
																										 else 
																												objExcel.Cells(TarSP+3,1)=sqlA
																												objExcel.Cells(TarSP+4,1)="TARGET DATA"
																												 objExcel.Cells(TarSP+4,1).Font.bold = True
																												objExcel.Cells(TarSP+4,1).Font.Color = RGB(0, 0, 255)
																												 objExcel.Cells(TarSP+5,1)="Target table is Empty!"
																									   End If
																			End if     

																	rs.close
															end if 	



													 'Cheking IF the QUERY has select Count and  HAVING other columns
										  	  If instr(1,chkcaB,strvar,1) <>0 and instr(1,chkcaB,",",1) <> 0  then 

'																	ColA = "'"&objExcel.Range("$B$6").Value&"'"&","
'																	ColB = "'"&objExcel.Range("$B$7").Value&"'"&","
'																	ColC = "'"&objExcel.Range("$B$8").Value&"'"&","
'																	ColD = "'"&objExcel.Range("$B$9").Value&"'"&","
'																	ColE = "'"&objExcel.Range("$B$10").Value&"'"

												 
																 sqlB=split(chkcaB,",")

																	 If  instr(sqlB(1),",")=0 Then
																		 sqlC=Split(trim(sqlB(1))," ")
																		 Col=SqlC(0) & " IN "
																		 else 
																		 col=trim(sqlB(1)) & " IN "
																	 End If
																	 


											
																 'Check for WHERE condition and Add  values to the Columns 
																   qry=ReplaceTest("where","WHERE",sqlA)
																		'msgbox qry
																			If instr(qry,"WHERE") =0 Then
																					tarSQL=qry&"  where " & col  &"("&Valu&")"
																			else 
																					whr= " WHERE "&  col  & "("&Valu&")" & " and "
																					tarSQL= replace (qry, "WHERE", Whr)
																			End If
											
																		'tarSQL = replace(tarSQL, ";", "")
																		
																		On error resume next
																				rs.Open tarSQL
                                        
																		
																		'Criteria TWO error msg Handeling
																		if 	err.number<> 0 then
																				errMsg= Err.Source & Err.Description
																				objExcel.Cells(TarSP+5,1)="Error"
																				objExcel.Cells(TarSP+5,1).Font.bold = True
																				objExcel.Cells(TarSP+5,1).Font.Color = RGB(255, 0, 0)
																				objExcel.Cells(TarSP+6,1)=errMsg
																				objExcel.Cells(TarSP+6,1).Font.Color = RGB(255, 0, 0)
																				objExcel.Cells(TarSP+6,1).WrapText=FALSE
																				Err.Clear
																		
																		End if


																		  ' Criteria THREE when record  NOT found
																		If rs.EOF = True Then
                                    
																			objExcel.Cells(TarSP+3,4)= tarSQL   
                                      objExcel.Cells(TarSP+3,4).Font.Size = 8
                                      objExcel.Cells(TarSP+3,4).Font.Italic = True
																			'objExcel.Range("A"&TarSP+3,"Z"&TarSP+3).Interior.Color = RGB(196, 189, 151)
																			objExcel.Cells(TarSP+4,1)="TARGET DATA"
																			objExcel.Cells(TarSP+5,1)="No Data Found!!"
                                    							  objExcel.Cells(TarSP+5,1).Interior.Color =RGB(255, 0, 0)
																		'	objExcel.cells(TarSP+5,1).Font.Color = RGB(50,205,50)
																			 objExcel.Cells(TarSP+4,1).Font.bold = True
																			 objExcel.Cells(TarSP+4,1).Font.Color = RGB(0, 0, 255)
																		End If
												
												
																	   ' Criteria FOUR when record found
																		 If not rs.eof then 
																					objExcel.Cells(TarSP+4,1)="TARGET DATA"
																					objExcel.Cells(TarSP+4,1).Font.bold = True
																					objExcel.Cells(TarSP+4,1).Font.Color = RGB(0, 0, 255)
																					fldCount =rs.Fields.Count 
																					objExcel.Range("A"&TarSP+5,"Z"&TarSP+5).Interior.Color = RGB(50, 205, 50)
																				  objExcel.Range("A"&TarSP+5,"Z"&TarSP+5).Font.Color =   RGB(255, 255, 255)
																
																				
																					For iCol = 1 To fldCount 
																					objExcel.Cells(TarSP+5, iCol).Value =rs.Fields(iCol - 1).Name 
																					Next 
																					
																					recArray =rs.GetRows(noofRecords) 
																					recCount = UBound(recArray, 2) + 1 
																		End if

												 objExcel.Cells(TarSP+3,4)= tarSQL                           
                         objExcel.Cells(TarSP+3,4).Font.Size = 8
                         objExcel.Cells(TarSP+3,4).Font.Italic = True
												' objExcel.Range("A"&TarSP+3,"Z"&TarSP+3).Interior.Color = RGB(196, 189, 151)
												objExcel.Cells(TarSP+6,1).Resize(recCount, fldCount).Value = TransposeDim(recArray) 
												  Set xlRng = objExcel.ActiveSheet.Columns("B:Z") 
												  xlRng.AutoFit
                          objExcel.Cells(5, 4).ColumnWidth = 30
								 ' Cheking IF the QUERY has select Count and  HAVING other columns
								 End if 	


							'Check If Query having MINUS key word
							 If  instr(1,sqlA,strvar,1) =0    Then
'									valA = "'"&objExcel.Range("$A$6").Value&"'"&","
'									valB = "'"&objExcel.Range("$A$7").Value&"'"&","
'									valC = "'"&objExcel.Range("$A$8").Value&"'"&","
'									valD = "'"&objExcel.Range("$A$9").Value&"'"&","
'									valE = "'"&objExcel.Range("$A$10").Value&"'"
'MsgBox SqlA

									'check for one or more columns. Pick the column name and add IN condition
										   If instr(sqlA, ",") =0 Then
												sqlB=ReplaceTest("from","FROM",sqlA)
												sqlC=split(sqlB,"FROM")
												sqlD=sqlC(0)
												sqlE=ReplaceTest("select","SELECT",sqlD)
												sqlF=split(sqlE,"SELECT")
												sqlE=trim(sqlF(1))
												sqlG=split(sqlE," ")
												Col=sqlG(0) & " IN "
												'msgbox col 
												else 
												sqlA=ReplaceTest("distinct","DISTINCT",sqlA)
												sqlB=split(Replace(sqlA,"DISTINCT",""),",")
												HH=sqlB(0)
											'	MsgBox hh
												'sqlC=split(trim(replace(ucase(sqlB(0)),"SELECT","")), " ")
											sqlC=	replace(ucase(sqlB(0)),"SELECT","")
											sqlC=Trim(sqlC)
											sqlD = Split(sqlC,"")
											sqlE = sqlD(0)
											'MsgBox SQLE
                                                col=sqlE & " IN "
                                                'MsgBox col
											End If
							
							'Check for WHERE condition and Add  values to the Columns 
							qry=ReplaceTest("where","WHERE",sqlA)
							
										If instr(qry,"WHERE") =0 Then
                                                tarSQL=qry&"  where " & col  &"("&Valu&")"
										else 
                                              whr= " WHERE "&  col  & "("&Valu&")" & " and "
                                              'msgbox whr 
											 'tarSQL= replace (qry, "WHERE", Whr)
                       tarSQLA=StrReverse(trim(Qry))
                       whrA=StrReverse(trim(whr))
                       tarSQLB= replace (tarSQLA, "EREHW", WhrA,1,1)
                       tarSQL=StrReverse(trim(tarSQLB))
                       'Set objShell = CreateObject("Wscript.Shell") 
                      ' Wscript.Echo tarSQL
										End If

             ' Sorting query based on column 1,2 
							SortSQL =  "SELECT * FROM  (" & tarSQL & ") Order by 1, 2"              
              
							On error resume next
									'rs.Open tarSQL
                   rs.Open SortSQL
							
							'Criteria TWO error msg Handeling
							if 	err.number<> 0 then
									errMsg= Err.Source & Err.Description
									objExcel.Cells(TarSP+5,1)="Error"
									objExcel.Cells(TarSP+5,1).Font.bold = True
									objExcel.Cells(TarSP+5,1).Font.Color = RGB(255, 0, 0)
									objExcel.Cells(TarSP+6,1)=errMsg
									objExcel.Cells(TarSP+6,1).Font.Color = RGB(255, 0, 0)
									objExcel.Cells(TarSP+6,1).WrapText=FALSE
									Err.Clear
							
							End if


						' Criteria THREE when record  NOT found
						If rs.EOF = True Then
							objExcel.Cells(TarSP+3,4)= tarSQL              
              objExcel.Cells(TarSP+3,4).Font.Size = 8
              objExcel.Cells(TarSP+3,4).Font.Italic = True
							'objExcel.Range("A"&TarSP+3,"Z"&TarSP+3).Interior.Color = RGB(196, 189, 151)
							objExcel.Cells(TarSP+4,1)="TARGET DATA"
							objExcel.Cells(TarSP+5,1)="No Data Found!!"
              objExcel.Cells(TarSP+5,1).Interior.Color =RGB(146, 208, 80)
						'	objExcel.cells(TarSP+5,1).Font.Color = RGB(50,205,50)
							 objExcel.Cells(TarSP+5,1).Font.bold = True
               objExcel.Cells(TarSP+4,1).Font.bold = True
							 objExcel.Cells(TarSP+4,1).Font.Color = RGB(0, 0, 255)
						End If


					   ' Criteria FOUR when record found
					     If not rs.eof then 
                                    objExcel.Cells(TarSP+4,1)="TARGET DATA"
									objExcel.Cells(TarSP+4,1).Font.bold = True
									objExcel.Cells(TarSP+4,1).Font.Color = RGB(0, 0, 255)
									fldCount =rs.Fields.Count 
									
						    	  'objExcel.Range("A"&TarSP+5,"DZ"&TarSP+5).Interior.Color = RGB(166, 166, 166)
							      'objExcel.Range("A"&TarSP+5,"Z"&TarSP+5).Font.Color =   RGB(255, 255, 255)
										'Msgbox fldCount
									For iCol = 1 To fldCount 
									objExcel.Cells(TarSP+5, iCol).Value =rs.Fields(iCol - 1).Name 
                  objExcel.Cells(TarSP+5, iCol).Interior.Color = RGB(166, 166, 166)
									Next 
									  
		
									recArray =rs.GetRows(-1) 
									recCount = UBound(recArray, 2) + 1 
									'msgbox recCount
												If recCount > 	noofRecords Then
															 rs.close
'															ColA = "'"&objExcel.Range("$B$6").Value&"'"&","
'															ColB = "'"&objExcel.Range("$B$7").Value&"'"&","
'															ColC = "'"&objExcel.Range("$B$8").Value&"'"&","
'															ColD = "'"&objExcel.Range("$B$9").Value&"'"&","
'															ColE = "'"&objExcel.Range("$B$10").Value&"'"

																	'Get the max row occupied in the excel file 
'																	Row=mysheet.UsedRange.Rows.Count
'																	J=2
'																	
'																	'To read the data from the entire Excel file
'																	For  i= 6 to Row
'																	Val = Val &"," & "'" & mysheet.cells(i,j).value & "'" 
'																	Valu= replace(Val, ",","",1,1)
'																	Next
											
															' Picking the second column to add in WHERE clause
															sql=	split(tarsql, "FROM")
															sqlA= sql(0)
															
																		If   instr(sqlA,",")= 0 Then
																				 tarSQL=replace (tarSQL,"SELECT", "SELECT DISTINCT ")
																		    else 
																				  tarSQLA=Split(tarSQL, ",")
																				  tarSQLB=split(ltrim(tarSQLA(1))," ")
'																				tarSQL=tarSQL&" and " & tarSQLB(0) & " IN " &"("&ColA&ColB&ColC&ColD&ColE&")"

																				 whr= " WHERE "  & tarSQLB(0) & " IN " &"("&ValuA&")" & " and "
																				  'tarSQL= replace (tarSQL, "WHERE", Whr)
                                             tarA=StrReverse(trim(tarSQL))
                                           whrA=StrReverse(trim(whr))
                                           tarB= replace (tarA, "EREHW", WhrA,1,1)
                                           tarSQL=StrReverse(trim(tarB))
																				
																		End If
																 
                               ' Sorting query based on column 1,2 
						                  	SortSQL =  "SELECT * FROM  (" & tarSQL & ") Order by 1, 2"     
                                 
														  On error resume next
															'rs.Open tarSQL
                              rs.Open SortSQL
															recArray =rs.GetRows(noofRecords) 
															recCount = UBound(recArray, 2) + 1 
															
												'No of record count  check   		
											   	End If
                        '  msgbox "I m here"
												objExcel.Cells(TarSP+3,4)= tarSQL                        
                        objExcel.Cells(TarSP+3,4).Font.Size = 8
                        objExcel.Cells(TarSP+3,4).Font.Italic = True
											        'objExcel.Range("A"&TarSP+3,"Z"&TarSP+3).Interior.Color = RGB(196, 189, 151)
												objExcel.Cells(TarSP+6,1).Resize(recCount, fldCount).Value = TransposeDim(recArray) 
                        
                         'HIGHLIGHTING MISSMATCHE's
                                   For Check = 1 to noofRecords                                   
                                            If objExcel.Cells(Check+5,1) = objExcel.Cells((TarSP+5)+Check,1) THEN                                       
                                                 For Col = 1 to (fldCount)-1  'Still last position of the Column                                           
                                                       ' If trim(objExcel.Cells(Check+5,col+1)) <> trim(objExcel.Cells((TarSP+5)+Check,Col+1))  THEN   
                                                       ' If Replace (CleanString(objExcel.Cells(Check+5,col+1))," ","**") <> Replace(CleanString(objExcel.Cells((TarSP+5)+Check,Col+1))," ","**")  THEN 
                                                     
                                                        If Replace (objExcel.Cells(Check+5,col+1)," ","**") <> Replace(objExcel.Cells((TarSP+5)+Check,Col+1)," ","**")  THEN   
                                                                                                           
                                                         
                                                            objExcel.Cells(Check+5,col+1).Interior.Color = RGB(255,255,140)
                                                            objExcel.Cells((TarSP+5)+Check,Col+1).Interior.Color = RGB(255,255,140)                                                 
                                                        End If 
                                                        
                                                  Next 
                                                  Col=1
                                            '  objExcel.Cells(Check+22,1).Interior.Color = RGB(255,255,0)
                                            END IF  
                                        NEXT   
                        
												  Set xlRng = objExcel.ActiveSheet.Columns("B:Z") 
												  xlRng.AutoFit
                          objExcel.Cells(5, 4).ColumnWidth = 30
							' when record found check												
							  End if 
							  
				' when Select Count not found check							  
				  End If		

			   'Close the Recordset  
				rs.close
				objDB.Close
				'Destroy the objDB object
				Set objDB = Nothing
			 
						
'										
				' when sheet FAIL check  															
			   End if
'				  'Close the Recordset  
             
Next

  
                     

		   objWorkbook.Worksheets(1).select
		   objExcel.ActiveWorkbook.Save
		   objExcel.ActiveWorkbook.Close
           objExcel.quit
		   Set mysheet =nothing
			Set objWorkbook = nothing
			Set  objExcel = nothing
			Set objWorksheet  = nothing
			

'		    rs.close
'			objDB.Close
'			 'Destroy the objDB object
'			Set objDB = Nothing
	
End Function




Function SendEmail()


'Email the files
'strFolder = rootFP
'strExt = "xlsx"

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFolder = objFSO.GetFolder(strFolder)


' First, create the message

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = EmailSubject & " - Execution Summary"
objMessage.From = """" & EmailFromName & """ <" & EmailFrom & ">"
objMessage.To = EmailTo
objMessage.CC= EmailCC
objMessage.TextBody = EmailBody

'To Send Single File, hardcoded

objMessage.AddAttachment  (rootFP&"\"&"Execution-Summary.xlsx" )

'This is Used to send Multiple Files in the Folder

		'For Each objFile In objFolder.Files
		'strFileExt = objFSO.GetExtensionName(objFile.Path)
		'objMessage.AddAttachment objFile.Path
		'msgbox "Got file!"
		'Next




' Second, configure the server

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

'objMessage.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'objMessage.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon
'objMessage.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

'objMessage.Configuration.Fields.Item _
'("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update

' Now send the message!

objMessage.Send

End Function

Function SpecialTrim(ByVal sString)
 On Error Resume Next
 If VarType(sString) <> vbString Then Exit Function
 Dim objRegExp:Set objRegExp = New RegExp
 objRegExp.Pattern = "^\s+|\s+$|\r\n"
 objRegExp.Ignorecase = True
 objRegExp.Global = True
 SpecialTrim = objRegExp.Replace(sString, "")
End Function
