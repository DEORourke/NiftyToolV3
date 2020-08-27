' VBA Macros below were pulled from the spreadsheet responsilbe for opening various other company spreadsheets and aggregating relivant information. The spreadsheet itself was not included in this repository as it contained company inforamtion.

Sub BigTestMac()
	
	MsgBox hi

End Sub


Sub Update_Parts_List()
	
	Dim csvOpen As Boolean
	Dim partsListOpen As Boolean
	
	Dim offset As Integer
	
	Dim outputCol As Range
	Dim partsHPData As Range
	Dim PartsNCRData As Range
	Dim PartsMiscData As Range
	
	Dim Equipment As String
	Dim exportPath As String
	Dim fileName As String
	Dim itemType As String
	Dim partsListPath As String
	Dim partsListFile As String
	
	Dim wkb As Workbook
	
	Dim partsHP As Worksheet
	Dim partsNCR As Worksheet
	Dim partsMisc As Worksheet
	Dim Update As Worksheet
	Dim wks As Worksheet
	
	Set wkb = Workbooks("NiftyInfoUpdate.xlsm")
	Set wks = wkb.Worksheets("Parts List")
	exportPath = "http://teams.company.com/sites/slrs/Files/"
	fileName = "NiftyPartsInfo.csv"
	Set outputCol = wks.Range("A2")
	Set Update = wkb.Worksheets("Nifty Info Update Tool")
	partsListPath = Trim(Update.Range("D6").Value)
	partsListFile = Right(partsListPath, Len(partsListPath) - InStrRev(partsListPath, "\"))
	
	For Each File In Workbooks
	If File.Name = partsListFile Then
		partsListOpen = True
	End If
	If File.Name = fileName Then
		csvOpen = True
    End If
	Next File

	If csvOpen = True Then
	MsgBox "Macro can not continue while " & fileName & " is open. Please close the workbook and run the macro again."
	End
	End If

	If partsListOpen = True Then
	Set partsHP = Workbooks(partsListFile).Worksheets("HP ")
	Set partsNCR = Workbooks(partsListFile).Worksheets("NCR")
	Set partsMisc = Workbooks(partsListFile).Worksheets("Miscellaneous Equipment")
	Else
	Set partsHP = Workbooks.Open(partsListPath, , True).Worksheets("HP ")
	Set partsNCR = Workbooks(partsListFile).Worksheets("NCR")
	Set partsMisc = Workbooks(partsListFile).Worksheets("Miscellaneous Equipment")
	End If
	
	Set partsHPData = partsHP.Range("C2:C60")
	Set PartsNCRData = partsNCR.Range("C2:C60")
	Set PartsMiscData = partsMisc.Range("C2:C60")
	
	offset = 0
	Equipment = ""
	itemType = ""
	wks.Range("2:2000").ClearContents
	
	For Each Line In partsHPData
	
	If Line.offset(0, -1).Font.Bold And Line.offset(0, -1).Value <> "" Then
		Equipment = Line.offset(0, -1).Value
	End If
	If Line.Font.Bold Then
		itemType = Line.Value
	End If
	
	If IsNumeric(Line) And Not IsEmpty(Line) Then
		outputCol.offset(offset, 0).Value = "HP"
		outputCol.offset(offset, 1).Value = Equipment
		outputCol.offset(offset, 2).Value = itemType
		outputCol.offset(offset, 3).Value = Line.Value
		outputCol.offset(offset, 4).Value = Line.offset(0, 1).Value
		offset = offset + 1
	End If
	
	Next Line
	
	For Each Line In PartsNCRData
	
	If Line.offset(0, -1).Font.Bold And Line.offset(0, -1).Value <> "" Then
		Equipment = Line.offset(0, -1).Value
	End If
	If Line.Font.Bold Then
		itemType = Line.Value
	End If
    
	If IsNumeric(Line) And Not IsEmpty(Line) Then
		outputCol.offset(offset, 0).Value = "NCR"
		outputCol.offset(offset, 1).Value = Equipment
		outputCol.offset(offset, 2).Value = itemType
		outputCol.offset(offset, 3).Value = Line.Value
		outputCol.offset(offset, 4).Value = Line.offset(0, 1).Value
		offset = offset + 1
	End If
    
	Next Line

	For Each Line In PartsMiscData
	
	If Line.offset(0, -1).Font.Bold And Line.offset(0, -1).Value <> "" Then
		Equipment = Line.offset(0, -1).Value
	End If
	If Line.Font.Bold Then
		itemType = Line.Value
    End If
    
    If IsNumeric(Line) And Not IsEmpty(Line) Then
        outputCol.offset(offset, 0).Value = "Misc"
        outputCol.offset(offset, 1).Value = Equipment
        outputCol.offset(offset, 2).Value = itemType
        outputCol.offset(offset, 3).Value = Line.Value
        outputCol.offset(offset, 4).Value = Line.offset(0, 1).Value
        offset = offset + 1
    End If
    
	Next Line

	'Export_As_Csv exportPath, wks, fileName

	'Update.Range("F10").Value = "Last run " & Format(Now(), "MM/dd HH:mm")

	If partsListOpen = False Then  ' closes workbook after updating if it was not open before.
		Workbooks(partsListFile).Close savechanges:=False
	End If

End Sub


Sub Update_Phone_List()
	
	Dim phoneDirOpen As Boolean
	Dim csvOpen As Boolean
	
	Dim offset As Integer
	
	Dim outputCol As Range
	Dim phoneDirData As Range
	
	Dim exportPath As String
	Dim fileName As String
	Dim Header As String
	Dim listName As String
	Dim phoneDirPath As String
	Dim phoneDirFile As String

	Dim wkb As Workbook

	Dim phoneDir As Worksheet
	Dim Update As Worksheet
	Dim wks As Worksheet

	Set wkb = Workbooks("NiftyInfoUpdate.xlsm")
	Set wks = wkb.Worksheets("Phone List")
	exportPath = "http://teams.company.com/sites/slrs/Files/"
	fileName = "NiftyPhoneInfo2.csv"
	Set outputCol = wks.Range("A2")
	Set Update = wkb.Worksheets("Nifty Info Update Tool")
	phoneDirPath = Trim(Update.Range("D5").Value)
	phoneDirFile = Right(phoneDirPath, Len(phoneDirPath) - InStrRev(phoneDirPath, "\"))

	For Each File In Workbooks
	If File.Name = phoneDirFile Then
		phoneDirOpen = True
	End If
	If File.Name = fileName Then
		csvOpen = True
	End If
	Next File
	
	If csvOpen = True Then
	MsgBox "Macro can not continue while " & fileName & " is open. Please close the workbook and 	run the macro again."
	End
	End If

	If phoneDirOpen = True Then
	Set phoneDir = Workbooks(phoneDirFile).Worksheets("Retail Systems Contact List")
	Else
	Set phoneDir = Workbooks.Open(phoneDirPath, , True).Worksheets("Retail Systems Contact List")
	End If

	Set phoneDirData = phoneDir.Range("B2", phoneDir.Range("B2").End(xlDown))   'Ends at row 40 for some reason. Probably because there's a space on row 39?
	'Set phoneDirData = phoneDir.Range("A3:A150")    'Go back and make this work like it's supposed to.

	offset = 0
	wks.Range("2:2000").ClearContents

	For Each Name In phoneDirData
    
	If Name.Font.Bold Then
		Header = Trim(Name.Value)
		listName = Trim(Name.offset(0, -1).Value)
	End If
    
	If Name <> "" Then
		If Name.offset(0, 1).Value <> "" Or Name.offset(0, 2).Value <> "" Or Name.offset(0, 4).Value <> "" Or Name.offset(0, 6).Value <> "" Then
		outputCol.offset(offset, 0).Value = listName 'Category
		outputCol.offset(offset, 1).Value = Header  ' Department
		outputCol.offset(offset, 2).Value = Trim(Name.Value)  ' Name
		outputCol.offset(offset, 3).Value = Trim(Name.offset(0, 1).Value) ' Office
		outputCol.offset(offset, 4).Value = Trim(Name.offset(0, 2).Value) ' Mobile
			outputCol.offset(offset, 5).Value = Trim(Name.offset(0, 4).Value) ' UID
			outputCol.offset(offset, 6).Value = Trim(Name.offset(0, 5).Value) ' Email
			outputCol.offset(offset, 7).Value = Trim(Name.offset(0, 6).Value) ' IM
			outputCol.offset(offset, 8).Value = Trim(Name.offset(0, 7).Value) ' Notes
            
			offset = offset + 1
			End If
		End If
	Next Name
    
	Export_As_Csv exportPath, wks, fileName

	Update.Range("H10").Value = "Last run " & Format(Now(), "MM/dd HH:mm")

	If phoneDirOpen = False Then  ' closes workbook after updating if it was not open before.
		Workbooks(phoneDirFile).Close savechanges:=False
	End If

End Sub


Sub Update_CP_Only()
	
	Dim cpInfoOpen As Boolean
	Dim csvOpen As Boolean
	
	Dim offset As Integer
	
	Dim cpInfoData As Range
	Dim outputCol As Range
	Dim storeList As Range
	
	Dim cpinfoFile As String
	Dim cpInfoPath As String
	Dim exportPath As String
	Dim fileName As String
	
	Dim wkb As Workbook
	
	Dim cpInfo As Worksheet
	Dim Update As Worksheet
	Dim wks As Worksheet
	
	exportPath = "http://teams.company.com/sites/slrs/Files/"
	fileName = "NiftyStoreInfo.csv"
	Set wkb = Workbooks("NiftyInfoUpdate.xlsm")
	Set wks = wkb.Worksheets("StoreData")
	Set Update = wkb.Worksheets("Nifty Info Update Tool")
	cpInfoPath = Trim(Update.Range("D3").Value)
	cpinfoFile = Right(cpInfoPath, Len(cpInfoPath) - InStrRev(cpInfoPath, "\"))
	Set outputCol = wks.Range("A2")
	Set storeList = wks.Range("A2", wks.Range("A2").End(xlDown))
	
	For Each File In Workbooks
	If File.Name = cpinfoFile Then
		cpInfoOpen = True
	End If
	If File.Name = fileName Then
        csvOpen = True
	End If
	Next File
	
	If csvOpen = True Then
	MsgBox "Macro can not continue while " & fileName & " is open. Please close the workbook and run the macro again."
	End
	End If
	
	If cpInfoOpen = True Then
	Set cpInfo = Workbooks(cpinfoFile).Worksheets("CP Info")
	Else
	Set cpInfo = Workbooks.Open(cpInfoPath, , True).Worksheets("CP Info")
	End If
	
	Set cpInfoData = cpInfo.Range("A5", cpInfo.Range("A5").End(xlDown))
	
	For Each Store In storeList
	For Each Row In cpInfoData
		If Store = Row Then
		If Store > 9000 And outputCol.offset(offset, 12).Value <> "SBO" Then
			outputCol.offset(offset, 28).Value = Trim(Row.offset(0, 12).Value)  'CP username
			outputCol.offset(offset, 29).Value = Trim(Row.offset(0, 13).Value)  'CP password
		End If
		End If
	Next
	offset = offset + 1
	Next Store
	
	Export_As_Csv exportPath, wks, fileName
	
	Update.Range("D10").Value = "Last run " & Format(Now(), "MM/dd HH:mm")

	If cpInfoOpen = False Then  ' closes workbook after updating if it was not open before.
	Workbooks(cpinfoFile).Close savechanges:=False
	End If
	
End Sub


Sub Update_Store_Info()
	
	Dim Response
	
	Dim cpInfoOpen As Boolean
	Dim csvOpen As Boolean
	Dim masterOpen As Boolean
	Dim rolloutOpen As Boolean
	
	Dim HouAPM As Integer
	Dim offset As Integer
	Dim Total As Integer
	
	Dim HouAPL As String
	Dim HouAPR As String
	Dim Vlan1Oct As String
	Dim Vlan3Oct As String
	
	Dim outputCol As Range
	Dim rolloutData As Range
	Dim cpInfoData As Range
	Dim masterData As Range
	Dim listingData As Range
	Dim adtlInfoData As Range
	
	Dim exportPath As String
	Dim fileName As String
	Dim rolloutPath As String
	Dim rolloutFile As String
	Dim cpInfoPath As String
	Dim cpinfoFile As String
	Dim masterPath As String
	Dim masterFile As String
	
	Dim wkb As Workbook
	
	Dim adtlInfo As Worksheet
	Dim cpInfo As Worksheet
	Dim Master As Worksheet
	Dim Rollout As Worksheet
	Dim Update As Worksheet
	Dim storeListing As Worksheet
	Dim wks As Worksheet
	
	Response = MsgBox("This function may take upwards of 10 minutes, during which time Excel will not respond to user input. Would you like to continue?", vbYesNo + vbAlert, "Hold on a moment")
	
	If Response = vbNo Then
	End
	End If
	
	Set wkb = Workbooks("NiftyInfoUpdate.xlsm")
	Set wks = Workbooks("NiftyInfoUpdate.xlsm").Worksheets("StoreData")
	Set adtlInfo = wkb.Worksheets("adtlInfo")
	Set storeListing = wkb.Worksheets("Store Listing")
	Set Update = wkb.Worksheets("Nifty Info Update Tool")
	
	Set outputCol = wks.Range("A2")
	
	exportPath = "http://teams.company.com/sites/slrs/Files/"
	fileName = "NiftyStoreInfoWithClosed.csv"
	rolloutPath = Trim(Update.Range("D2").Value)
	rolloutFile = Right(rolloutPath, Len(rolloutPath) - InStrRev(rolloutPath, "\"))
	cpInfoPath = Trim(Update.Range("D3").Value)
	cpinfoFile = Right(cpInfoPath, Len(cpInfoPath) - InStrRev(cpInfoPath, "\"))
	masterPath = Trim(Update.Range("D4").Value)
	masterFile = Right(masterPath, Len(masterPath) - InStrRev(masterPath, "\"))

	
	' The next several lines check the names of open workbooks to determine if Rollout sheet is not already open.
	
	For Each File In Workbooks
	If File.Name = rolloutFile Then
		rolloutOpen = True
	End If
	If File.Name = cpinfoFile Then
		cpInfoOpen = True
	End If
	If File.Name = masterFile Then
		masterOpen = True
	End If
	If File.Name = fileName Then
		csvOpen = True
	End If
	Next File
	
	If csvOpen = True Then
	MsgBox "Macro can not continue while " & fileName & " is open. Please close the workbook and run the macro again."
	End
	End If
	
	If rolloutOpen = True Then
	Set Rollout = Workbooks(rolloutFile).Worksheets("Stores")
	Else
	Set Rollout = Workbooks.Open(rolloutPath, , True).Worksheets("Stores")
	End If
	
	If cpInfoOpen = True Then
	Set cpInfo = Workbooks(cpinfoFile).Worksheets("CP Info")
	Else
	Set cpInfo = Workbooks.Open(cpInfoPath, , True).Worksheets("CP Info")
	End If
	
	If masterOpen = True Then
	Set Master = Workbooks(masterFile).Worksheets("Complete Store List")
	Else
	Set Master = Workbooks.Open(masterPath, , True).Worksheets("Complete Store List")
	End If
	
	
	Set rolloutData = Rollout.Range("A3", Rollout.Range("A3").End(xlDown))
	Set cpInfoData = cpInfo.Range("A5", cpInfo.Range("A5").End(xlDown))
	Set masterData = Master.Range("A3", Master.Range("A3").End(xlDown))
	Set listingData = storeListing.Range("A3", storeListing.Range("A3").End(xlDown))
	Set adtlInfoData = adtlInfo.Range("A2", adtlInfo.Range("A2").End(xlDown))
	
	offset = 0
	wks.Range("2:2000").ClearContents
	
	' Retrieve store number and address from Store Rollout workbook.
	For Each Store In rolloutData
		'If Store.offset(0, 2) <> "Closed" And Store.offset(0, 2) <> "Cancelled" And IsNumeric(Store) Then
        'If Store.offset(0, 2) <> "Cancelled" And IsNumeric(Store) Then
        If IsNumeric(Store) Then
            outputCol.offset(offset, 0).Value = Trim(Store.Value) 'store number
            outputCol.offset(offset, 2).Value = Trim(Store.offset(0, 4).Value)   'Street address
            If Len(Trim(Store.offset(0, 5))) > 1 Then
                outputCol.offset(offset, 2).Value = outputCol.offset(offset, 2).Value & ", " & Trim(Store.offset(0, 5).Value)   'Adding Street address field 2, if applicable
            End If
            outputCol.offset(offset, 3).Value = Trim(Store.offset(0, 6).Value)    'city
            outputCol.offset(offset, 5).Value = Left(Store.offset(0, 8).Value, 5)    'zip
            outputCol.offset(offset, 15).Value = Store.offset(0, 26).Value  'ap model
            If Store < 9000 Then
                outputCol.offset(offset, 13).Value = "Corporate"    'Determine and list corprate or licencee store
            Else
                outputCol.offset(offset, 13).Value = Trim(Store.offset(0, 3).Value)
            End If
            
            If Store.offset(0, 3) = "Specific Franchise" Then

                outputCol.offset(offset, 19).Value = Store.offset(0, 27).Value  'Obtain only stand-beside PC and AP address if Specific Franchise store. Retrieve store information from rollout, cpinfo, and master store list workbooks otherwise.

                If Trim(Store.offset(0, 27).Value) <> "" Then   'If the stand-beside PC address is known, calculates address of the AP as the 3rd octet of the PC address +2
                    HouAPL = Left(Store.offset(0, 27), InStr(4, Store.offset(0, 27).Value, "."))
                    HouAPR = Right(Store.offset(0, 27), (Len(Store.offset(0, 27))) - (InStrRev(Store.offset(0, 27).Value, ".")) + 1)
                    HouAPM = Mid(Store.offset(0, 27), InStr(4, Store.offset(0, 27), ".") + 1, (InStrRev(Store.offset(0, 27), ".")) - (InStr(4, Store.offset(0, 27), ".") + 1)) + 2
                    outputCol.offset(offset, 17).Value = HouAPL & HouAPM & HouAPR
                    HouAPM = 0
                End If
                
            Else
                If IsNumeric(Store.offset(0, 28).Value) Then
                    outputCol.offset(offset, 1).Value = Store.offset(0, 28).Value   'orical number
                End If
                outputCol.offset(offset, 11).Value = Trim(Store.offset(0, 18).Value) & " - " & Trim(Store.offset(0, 19).Value)  'Broadband provider
                outputCol.offset(offset, 14).Value = Trim(Store.offset(0, 24).Value)    'router model
                outputCol.offset(offset, 12).Value = Trim(Store.offset(0, 17).Value)    'T1 curcuit ID
                
                Vlan1Oct = Store.offset(0, 40).Value
                Vlan3Oct = Store.offset(0, 34).Value
  ' so here we need a method to test if these variables are text, an if that's the case we can't do maths to them and will instead have to append the below values or skip them entirely?
  
                outputCol.offset(offset, 16).Value = Vlan1Oct & Store.offset(0, 41).Value + 1    'router address
                outputCol.offset(offset, 17).Value = Vlan1Oct & Store.offset(0, 41).Value + 4    'AP
                outputCol.offset(offset, 18).Value = Vlan1Oct & Store.offset(0, 41).Value + 10   'Produce scale
                outputCol.offset(offset, 19).Value = Vlan1Oct & Store.offset(0, 41).Value + 3    'Server
                outputCol.offset(offset, 20).Value = Vlan1Oct & Store.offset(0, 41).Value + 8    'Safe
                outputCol.offset(offset, 21).Value = Vlan1Oct & Store.offset(0, 41).Value + 2    'meat scale
                outputCol.offset(offset, 22).Value = Vlan3Oct & Store.offset(0, 35).Value + 5    'Sw2
                outputCol.offset(offset, 23).Value = Vlan1Oct & Store.offset(0, 41).Value + 12   'Mio PC
                outputCol.offset(offset, 24).Value = Vlan1Oct & Store.offset(0, 41).Value + 6    'Time clock

                Vlan1Oct = 0
                Vlan3Oct = 0
                
                For Each Line In masterData
                    If Store = Line Then
                        If InStr(Line.offset(0, 16).Value, "V8") > 0 Or InStr(Line.offset(0, 16).Value, "v8") > 0 Then
                            outputCol.offset(offset, 25).Value = "V8"
                        Else
                            outputCol.offset(offset, 25).Value = "V7"
                        End If
                        outputCol.offset(offset, 26).Value = Trim(Left(Line.offset(0, 16).Value, 3))
                        outputCol.offset(offset, 27).Value = Line.offset(0, 13).Value   'number of lanes
                    End If
                Next
                
                For Each Row In cpInfoData
                    If Store = Row Then
                        If Store > 9000 Then
                            outputCol.offset(offset, 28).Value = Trim(Row.offset(0, 12).Value)    'CP username
                            outputCol.offset(offset, 29).Value = Trim(Row.offset(0, 13).Value)    'CP password
                        End If
                        outputCol.offset(offset, 30).Value = Trim(Row.offset(0, 11).Value)    'CP company number
                        outputCol.offset(offset, 31).Value = Trim(Row.offset(0, 8).Value) 'old CP company number
                    End If
                Next
            End If
            
            For Each Item In listingData
                If Store = Item Then
                    outputCol.offset(offset, 9).Value = Trim(Item.offset(0, 10).Value)   'email address
                    outputCol.offset(offset, 32).Value = Trim(Item.offset(0, 13).Value)   'Sunday hours
                    outputCol.offset(offset, 33).Value = Trim(Item.offset(0, 14).Value)   'M-Sa hours
                    outputCol.offset(offset, 6).Value = Right(Item.offset(0, 17).Value, Len(Item.offset(0, 17).Value) - 7) 'Region
                    outputCol.offset(offset, 7).Value = Item.offset(0, 15).Value & " " & Item.offset(0, 16).Value  'DC
                    outputCol.offset(offset, 4).Value = Trim(Item.offset(0, 7).Value) 'state
                    outputCol.offset(offset, 8).Value = Trim(Item.offset(0, 9).Value) 'phone number
                End If
            Next
                
            For Each thing In adtlInfoData
                If Store = thing Then
                    outputCol.offset(offset, 34).Value = thing.offset(0, 1).Value   'Time zone
                    outputCol.offset(offset, 10).Value = thing.offset(0, 2).Value   'URL path
                    outputCol.offset(offset, 35).Value = thing.offset(0, 3).Value   'Has InTouch timeclock?
                    outputCol.offset(offset, 36).Value = thing.offset(0, 4).Value   'New helpdesk pilot launch date
                End If
            Next
            
            offset = offset + 1
            
        End If
	Next Store
	
	Export_As_Csv exportPath, wks, fileName
	
	Update.Range("B10").Value = "Last run " & Format(Now(), "MM/dd HH:mm")
	Update.Range("D10").Value = "Last run " & Format(Now(), "MM/dd HH:mm")
	
	If cpInfoOpen = False Then  ' closes workbook after updating if it was not open before.
	Workbooks(cpinfoFile).Close savechanges:=False
	End If
	If masterOpen = False Then
	Workbooks(masterFile).Close savechanges:=False
	End If
	If rolloutOpen = False Then
	Workbooks(rolloutFile).Close savechanges:=False
	End If

End Sub


Sub Export_As_Csv(exportPath As String, wks As Worksheet, fileName As String)
	
	Application.DisplayAlerts = False
	wks.Copy
	ActiveWorkbook.SaveAs fileName:=exportPath & fileName, FileFormat:=xlCSV   'exporting a copy of the results in CSV format
	ActiveWorkbook.Close
	Application.DisplayAlerts = True
	Workbooks("NiftyInfoUpdate.xlsm").Save
	
End Sub
