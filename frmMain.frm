VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Approach Flimsy Maker v2.6"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   OleObjectBlob   =   "frmMain.frx":0000
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim userClicked As Boolean
Dim xmlDoc As MSXML2.DOMDocument
Dim lastCycle As Integer
Dim lastYear As Integer

Private Sub cmdNextCycle_Click()
    Dim nextVol As String
    Dim remotePath As String
    Dim metaFile As String
    
    metaFile = Application.ActiveWorkbook.Path & "\TPPMetafile.xml"
    nextVol = getVolume(DateAdd("d", 28, Now))
    remotePath = "http://aeronav.faa.gov/d-tpp/" & nextVol & "/xml_data/d-TPP_Metafile.xml"
    'Call DownloadFile(remotePath, metaFile)
    frmProgressTPP.DownloadSource = remotePath
    frmProgressTPP.StorageLocation = metaFile
    frmProgressTPP.Show
    
    Call getAirfieldsFromXML(metaFile)
    Me.FLIPCycleComboBox.Text = ("[" & nextVol & "] " & getValidDates(nextVol))
    'Call MsgBox("Next Cycle Loaded!", vbOKOnly + vbInformation, "Download Complete")
End Sub

Private Sub GenPDFButton_Click()
    If (chkCloseAdobe.Value = True) Then
        Call killAdobeProcess
    End If
    'If (checkIfAdobeExists = True) Then
    '    Exit Sub
    'Else
        frmProgress.Show
    'End If
End Sub

Public Sub GenerateTheFlimsy()
Dim currPath As String
Dim tempPath As String
Dim downloadWeb As String
Dim cycleToSend As String
Dim endLoc As Integer
Dim dateHelper As Date
Dim i As Integer
Dim curJob As Integer
Dim totJobs As Integer
Dim oShell As Object

    totJobs = YourFlimsyListBox.ListCount
    If (chkPageNums.Value = True) Then
        printPageNums = True
        'one to download, one to add to the flimsy, one to print the expiration
        totJobs = totJobs * 3
    Else
        printPageNums = False
        'one to download, one to add to the flimsy
        totJobs = totJobs * 2
    End If
    If (chkTOC.Value = True) Then
        printTOC = True
        'one to make, one to add to the flimsy
        totJobs = totJobs + 2
    Else
        printTOC = False
    End If
    If (chkExpire.Value = True) Then
        totJobs = totJobs + 1
        printExpiration = True
    Else
        printExpiration = False
    End If
    'setup download, setup PDF, finish PDF, clean up files
    totJobs = totJobs + 4
    curJob = 1
    
    
    Call IterateProgress(curJob, totJobs, "Setting up file download")

    Set oShell = CreateObject("WScript.Shell")
    'currPath = Application.ActiveWorkbook.Path + "\"
    currPath = oShell.specialfolders("Desktop") + "\"
    tempPath = currPath + "tempFlipDIR" + Format(Now, "hh-nn-ss") + "\"
    MkDir tempPath
        
    downloadWeb = "http://aeronav.faa.gov/d-tpp/"
    
    cycleToSend = Me.FLIPCycleComboBox.Text
    endLoc = InStr(1, cycleToSend, "]")
    'store the new string in the global variable
    cycleData = Mid(cycleToSend, 2, endLoc - 2)
    endLoc = InStr(1, cycleToSend, " to ")
    cycleExpire = Mid(cycleToSend, endLoc + Len(" to "))
    dateHelper = CDate(cycleExpire)
    dateHelper = DateAdd("d", 1, dateHelper)
    cycleExpire = Format(dateHelper, "dd-MMM-yy")
    downloadWeb = downloadWeb + cycleData + "/"
    
    curJob = curJob + 1
    For i = 0 To YourFlimsyListBox.ListCount - 1
        Call IterateProgress(curJob, totJobs, "Downloading " & YourFlimsyListBox.List(i, 1) & " for " & YourFlimsyListBox.List(i, 0))
        Call DownloadFile(downloadWeb + YourFlimsyListBox.List(i, 2), tempPath + leastSigTwoDigits(i) + "_" + YourFlimsyListBox.List(i, 2))
        curJob = curJob + 1
    Next i
    
    
    Dim finalDoc As String
    finalDoc = combineFiles(curJob, totJobs, tempPath, currPath)
    'frmProgress.Hide
    Call frmProgress.CloseProgressBar
    rmDir tempPath
    If (finalDoc <> "") Then
        VBA.Shell "Explorer.exe " & Chr(34) & finalDoc & Chr(34), vbNormalFocus ' Use Windows Explorer to launch  the file.
    End If
End Sub

Private Sub MoveDownButton_Click()
    If (checkIfRowSelected = True) Then
        Dim i As Integer
        Dim j As Integer
        Dim holder As Integer
        Dim a, b, c As String
       
        For i = 0 To YourFlimsyListBox.ListCount - 2 Step 1
            If YourFlimsyListBox.Selected(i) = True Then
                a = YourFlimsyListBox.List(i + 1, 0)
                b = YourFlimsyListBox.List(i + 1, 1)
                c = YourFlimsyListBox.List(i + 1, 2)
                YourFlimsyListBox.List(i + 1, 0) = YourFlimsyListBox.List(i, 0)
                YourFlimsyListBox.List(i + 1, 1) = YourFlimsyListBox.List(i, 1)
                YourFlimsyListBox.List(i + 1, 2) = YourFlimsyListBox.List(i, 2)
                YourFlimsyListBox.List(i, 0) = a
                YourFlimsyListBox.List(i, 1) = b
                YourFlimsyListBox.List(i, 2) = c
                YourFlimsyListBox.Selected(i) = False
                YourFlimsyListBox.Selected(i + 1) = True
                Exit Sub
            End If
        Next i
    Else
        Call MsgBox("You must select only one item to move", vbExclamation + vbOKOnly, "Error")
    End If
End Sub

Private Sub MoveUpButton_Click()
    If (checkIfRowSelected = True) Then
        Dim i As Integer
        Dim j As Integer
        Dim holder As Integer
        Dim a, b, c As String
       
        For i = 1 To YourFlimsyListBox.ListCount - 1 Step 1
            If YourFlimsyListBox.Selected(i) = True Then
                a = YourFlimsyListBox.List(i - 1, 0)
                b = YourFlimsyListBox.List(i - 1, 1)
                c = YourFlimsyListBox.List(i - 1, 2)
                YourFlimsyListBox.List(i - 1, 0) = YourFlimsyListBox.List(i, 0)
                YourFlimsyListBox.List(i - 1, 1) = YourFlimsyListBox.List(i, 1)
                YourFlimsyListBox.List(i - 1, 2) = YourFlimsyListBox.List(i, 2)
                YourFlimsyListBox.List(i, 0) = a
                YourFlimsyListBox.List(i, 1) = b
                YourFlimsyListBox.List(i, 2) = c
                YourFlimsyListBox.Selected(i) = False
                YourFlimsyListBox.Selected(i - 1) = True
                Exit Sub
            End If
        Next i
    Else
        Call MsgBox("You must select only one item to move", vbExclamation + vbOKOnly, "Error")
    End If
End Sub

Private Function checkIfRowSelected() As Boolean
    Dim i As Integer
    Dim foundOne As Boolean
    foundOne = False
    For i = 0 To YourFlimsyListBox.ListCount - 1
        If (YourFlimsyListBox.Selected(i) = True) Then
            If (foundOne = True) Then
                checkIfRowSelected = False
                Exit Function
            End If
            foundOne = True
        End If
    Next i
    If (foundOne = True) Then
        checkIfRowSelected = True
    Else
        checkIfRowSelected = False
    End If
End Function

Private Sub RemoveAllButton_Click()
    Dim i As Integer
    i = YourFlimsyListBox.ListCount
    
    Do While (i >= 1)
        YourFlimsyListBox.RemoveItem (0)
        i = YourFlimsyListBox.ListCount
    Loop
End Sub

Private Sub UserForm_Initialize()
    'populate combo box w/ current and next FLIP volumes
    Application.Visible = False
    Dim currVol As String
    Dim nextVol As String
    Dim currCycle As String
    Dim validDates As String
    
    userClicked = False
    currVol = getVolume(Now)
    'Debug.Print "curr volume: " & currVol
    
    'next volume 28 days from now
    nextVol = getVolume(DateAdd("d", 28, Now))
    ''Dim tempVol As String 'used for checking date function
    ''tempVol = getVolume(DateAdd("d", 56, Now)) 'used for checking date function
    'Debug.Print "Next volume: " & nextVol
    currCycle = "[" & currVol & "] " & getValidDates(currVol)
    Me.FLIPCycleComboBox.Text = currCycle
    Me.FLIPCycleComboBox.AddItem (currCycle)
    Me.FLIPCycleComboBox.AddItem ("[" & nextVol & "] " & getValidDates(nextVol))
    ''Me.FLIPCycleComboBox.AddItem ("[" & tempVol & "] " & getValidDates(tempVol)) 'checks to see if my date function is working
    Me.Left = 100
    Me.Top = 100
    Me.chkExpire.Value = True
    Me.chkPageNums.Value = True
    Me.chkTOC.Value = True
    Me.chkCloseAdobe.Value = True
    'hides all the worksheets
    Call ChangeAllWSVis(True)
    '
    'Add tooltips
    '
    Me.FLIPCycleComboBox.ControlTipText = "Current FLIP cycle: [" & currVol & "] " & getValidDates(currVol)
    Me.ICAOComboBox.ControlTipText = "Step 1: Choose airfield from this list"
    Me.ProceduresListBox.ControlTipText = "Step 2: Select procedures you want in your flimsy"
    Me.AddProceduresButton.ControlTipText = "Step 3: Click to add these procedures to flimsy"
    Me.YourFlimsyListBox.ControlTipText = "These approaches will be added to your flimsy"
    'Me.FileFindButton.ControlTipText = "Step 4: Tell me where you want to save your flimsy"
    Me.GenPDFButton.ControlTipText = "Step 4: Click to generate your flimsy in .pdf format"
    
    Me.ShowWSButton.ControlTipText = "Show Excel worksheet (For debugging only)"
    'Me.ICAOTextBox.ControlTipText = "Comma-separated list of ICAO identifiers"
    
    '
    'Check for presence of FAA TPP Metafile
    '
    Dim metaFile As String
    metaFile = Application.ActiveWorkbook.Path & "\TPPMetafile.xml"
    Debug.Print "TPP Metafile Path: " & metaFile
    
    Dim downloadMetafileResponse As Integer
    Dim remotePath As String
    'Download the TPP Metafile if not existing already
    If Dir(metaFile) = "" Then
        downloadMetafileResponse = MsgBox("FAA Digital Airfield Database not present on system, would you like to download it now?" & vbCrLf & "This is required to use the Flimsy Maker", vbYesNo + vbQuestion, "Download Approach Database")
        
        If downloadMetafileResponse = vbYes Then
            remotePath = "http://aeronav.faa.gov/d-tpp/" & currVol & "/xml_data/d-TPP_Metafile.xml"
            'Call DownloadFile(remotePath, metaFile)
            frmProgressTPP.DownloadSource = remotePath
            frmProgressTPP.StorageLocation = metaFile
            frmProgressTPP.Show
            Call getAirfieldsFromXML(metaFile)
        End If
    Else
        If (isXMLCurrent(metaFile, currVol) = True) Then
            Call getAirfieldsFromXML(metaFile)
        Else
            downloadMetafileResponse = MsgBox("FAA Digital Airfield Database is not current; would you like to update it now?", vbYesNo + vbQuestion, "Out of Date Approach Database")
            If downloadMetafileResponse = vbYes Then
                remotePath = "http://aeronav.faa.gov/d-tpp/" & currVol & "/xml_data/d-TPP_Metafile.xml"
                'Call DownloadFile(remotePath, metaFile)
                frmProgressTPP.DownloadSource = remotePath
                frmProgressTPP.StorageLocation = metaFile
                frmProgressTPP.Show
                Call getAirfieldsFromXML(metaFile)
            Else
                Call getAirfieldsFromXML(metaFile)
            End If
        End If
    End If
End Sub


Private Sub AboutButton_Click()
    MsgBox "Flimsy Maker Version 2.6" + vbCrLf + vbCrLf + "Written by John ""ATIS"" Ayers, Matt ""Elmo"" Elmore and Jon ""Doc"" Andrew for the 559 FTS Billygoats and pilots everywhere." & vbCrLf & vbCrLf & "If you like our work, let us know!", vbOKOnly, "About Flimsy Maker"
End Sub

Private Sub RemoveButton_Click()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To YourFlimsyListBox.ListCount - 1
        If YourFlimsyListBox.Selected(i - j) = True Then
            YourFlimsyListBox.RemoveItem (i - j)
            j = j + 1
        End If
    Next i
End Sub

' Sub FileFindButton_Click()
'  Opens Save As dialog for the output file
'  TODO: Filter for .pdf only
'
Private Sub FileFindButton_Click()
    Dim fd As FileDialog
    Dim objFile As Variant
    Dim strFileName As String
    
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .ButtonName = "Select"
        .AllowMultiSelect = False
        .Title = "Save as"
        .InitialView = msoFileDialogViewList
        With .Filters
            .Clear
            .Add "All Files", "*.*"
            .Add "Adobe PDF", "*.pdf"
        End With
        
        If .Show = -1 Then
        
            For Each objFile In .SelectedItems
                strFileName = objFile
            Next objFile
        Else
            'user pressed cancel
        End If
    End With
    
    OutputFileText.Text = strFileName
    
    Set fd = Nothing
End Sub

Private Sub ICAOComboBox_Click()
    Dim currICAO As String
    
    ProceduresListBox.Clear
    
    currICAO = Left(ICAOComboBox.Text, 4)
    
    Debug.Print "selected: " & currICAO
    
    If Trim(currICAO) = "" Then
        ProceduresListBox.Clear
        Exit Sub
    End If
    
    'Now need to find the terminal procedures available for this field.
    If (Not xmlDoc Is Nothing) Then
            Dim recordNodes As IXMLDOMNodeList
            Dim xpath As String
            xpath = "//airport_name[@icao_ident='" & UCase(currICAO) & "']/record"
            Debug.Print xpath
            Set recordNodes = xmlDoc.SelectNodes(xpath)
            Debug.Print " Found " & recordNodes.Length & " procedures for " & currICAO
            
            'iterate through records and find procedures
            Dim recordNode As IXMLDOMNode
            Dim procedureStr As String
            Dim pdfStr As String
            Dim i As Integer
            i = 0
            For Each recordNode In recordNodes
                procedureStr = recordNode.SelectSingleNode("chart_name").Text
                pdfStr = recordNode.SelectSingleNode("pdf_name").Text
                
                'ProceduresListBox.AddItem (procedureStr & ";" & pdfStr)
                ProceduresListBox.AddItem
                ProceduresListBox.List(i, 0) = procedureStr
                ProceduresListBox.List(i, 1) = pdfStr
                i = i + 1
            Next
    Else
        MsgBox "No FAA Procedure Database Loaded. Restart the app and select 'Yes' when prompted to download it.", vbOKOnly
    End If
    
End Sub

Private Sub AddProceduresButton_Click()
    'Determine which procedures to add
    'TODO: Clear selection after adding (or remove added items)
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim currICAO As String
    
    currICAO = Left(ICAOComboBox.Text, 4)
    j = YourFlimsyListBox.ListCount
    
    For i = 0 To ProceduresListBox.ListCount - 1
        If ProceduresListBox.Selected(i) = True Then
            
            'Make sure this isn't already part of Your Flimsy by checking the pdf name
            Dim alreadyHere As Boolean
            alreadyHere = False
            
            'Search through Your Flimsy for procedure about to be added
            If YourFlimsyListBox.ListCount > 0 Then
                For k = 0 To YourFlimsyListBox.ListCount - 1
                    If YourFlimsyListBox.List(k, 2) = ProceduresListBox.List(i, 1) Then
                        alreadyHere = True
                        Exit For
                    End If
                Next k
            End If
            
            If alreadyHere = False Then
                YourFlimsyListBox.AddItem
                YourFlimsyListBox.List(j, 0) = currICAO
                YourFlimsyListBox.List(j, 1) = ProceduresListBox.List(i, 0)
                YourFlimsyListBox.List(j, 2) = ProceduresListBox.List(i, 1)
                j = j + 1
            End If
        End If
    Next i
    
    'TODO: Unselect procedures
End Sub

Private Sub SaveFavoritesButton_Click()
    Dim i As Integer
    
    Sheet1.Activate
    Sheet1.Range("A1", Columns("A").SpecialCells(xlCellTypeLastCell)).Clear
    Sheet1.Range("B1", Columns("B").SpecialCells(xlCellTypeLastCell)).Clear
    Sheet1.Range("C1", Columns("C").SpecialCells(xlCellTypeLastCell)).Clear
    
    For i = 0 To YourFlimsyListBox.ListCount - 1
        Sheet1.Cells(i + 1, 1) = YourFlimsyListBox.List(i, 0)
        Sheet1.Cells(i + 1, 2) = YourFlimsyListBox.List(i, 1)
        Sheet1.Cells(i + 1, 3) = YourFlimsyListBox.List(i, 2)
    Next i
    ThisWorkbook.Save
End Sub

Private Sub LoadFavoritesButton_Click()
    Dim i As Integer
    Dim endOfFavorites As Boolean
    
    YourFlimsyListBox.Clear
    
    ' Max of 200 favorites (even if more are loaded)
    For i = 0 To 200
        If (Trim(Sheet1.Cells(i + 1, 1)) <> "" And Trim(Sheet1.Cells(i + 1, 2)) <> "" And Trim(Sheet1.Cells(i + 1, 3)) <> "") Then
            YourFlimsyListBox.AddItem
            YourFlimsyListBox.List(i, 0) = Sheet1.Cells(i + 1, 1)
            YourFlimsyListBox.List(i, 1) = Sheet1.Cells(i + 1, 2)
            YourFlimsyListBox.List(i, 2) = Sheet1.Cells(i + 1, 3)
        End If
    Next i
End Sub

Private Sub ShowWSButton_Click()
    Application.Visible = True
    userClicked = True
    Unload Me
    'this tricks the application into thinking it's saved--it does NOT save the workbook though
    ThisWorkbook.Saved = True
    Application.Quit
End Sub

Private Function isXMLCurrent(ByRef metaFile As String, ByRef currCycle As String) As Boolean
    Set xmlDoc = New MSXML2.DOMDocument
    Call xmlDoc.Load(metaFile)
    Dim cycleNode As IXMLDOMNodeList
    
    Set cycleNode = xmlDoc.SelectNodes("//digital_tpp")
    Dim node As IXMLDOMNode
    Dim attr As IXMLDOMAttribute
    
    For Each node In cycleNode
        Set attr = node.Attributes.getNamedItem("cycle")
        If (Not attr Is Nothing) Then
            If (attr.Text = currCycle) Then
                isXMLCurrent = True
            Else
                isXMLCurrent = False
            End If
            Exit Function
        End If
    Next node
End Function

Private Sub getAirfieldsFromXML(ByRef metaFile As String)
    'Parse the TPP Metafile
    Set xmlDoc = New MSXML2.DOMDocument
    Call xmlDoc.Load(metaFile)
    
    Dim airfieldNodes As IXMLDOMNodeList
    Set airfieldNodes = xmlDoc.SelectNodes("//airport_name")
    Debug.Print " Found " & airfieldNodes.Length & " airfields"
          
    Dim icaoAttr As IXMLDOMAttribute
    Dim nameAttr As IXMLDOMAttribute
    
    Dim node As IXMLDOMNode
    Dim childNode As IXMLDOMNode
           
    Dim icaoList() As String
    Dim i As Integer
    i = 0

    For Each node In airfieldNodes
        Set icaoAttr = node.Attributes.getNamedItem("icao_ident")
        Set nameAttr = node.Attributes.getNamedItem("ID")
        
        If (Not icaoAttr Is Nothing) Then
            If (Trim(icaoAttr.Text) <> "") Then
                i = i + 1
                ReDim Preserve icaoList(i)         'bad for performance, but only running this loop once.
                
                icaoList(i) = icaoAttr.Text & " " & nameAttr.Text
            End If
        End If
    Next
    
    Call Sort(icaoList)
    
    For i = 0 To UBound(icaoList)
        ICAOComboBox.AddItem (icaoList(i))
    Next
End Sub

'Function getValidDates
Private Function getValidDates(ByRef volume As String)
    Dim year As Integer
    Dim cycle As Integer
    
    year = CInt(Left(volume, 2))
    cycle = CInt(Right(volume, 2))
    
    Dim withYear As Date
    Dim withCycleDays As Date
    
    Dim startDate As Date
    Dim endDate As Date
    Dim yearsSince2016 As Integer
    Dim yearDays As Integer
    
    yearsSince2016 = year - 16
    yearDays = yearsSince2016 * (13 * 28) '13 cycles per year, 28 days per cycle
    startDate = DateAdd("d", 28 * (cycle - 1) + yearDays, #1/7/2016#)
    'startDate = DateAdd("d", 28 * (cycle - 1) * (yearsSince2016 + 1), #1/7/2016#)
    endDate = DateAdd("d", 27, startDate)
    
    getValidDates = Format(startDate, "dd-mmm-yy") & " to " & Format(endDate, "dd-mmm-yy")
End Function

'Function getVolume
' Arguments: arbitrary Date in date format (after 7 Jan 2016)
' Returns: String in format (YYCC) where CC is the 28-day FLIP cycle
' Example: getVolume(DateValue("January 30 2016")) returns "1602"
Private Function getVolume(ByVal myDate As Date)
    Dim currYear As String
    Dim julianDate As Integer
    Dim cycle As Integer
    Dim currCycle As String
    
    currYear = leastSigTwoDigits(DatePart("yyyy", myDate))
    If lastYear = 0 Then
        lastYear = currYear
    End If
    'Cycle 1601 started on 7 Jan 2016, so reference future dates from that
    Dim daysSince As Integer
    Dim myCycle As Integer
    daysSince = DateDiff("d", #1/7/2016#, myDate)
    'Debug.Print "days since: " & daysSince
    
    'use \ instead of / to force integer division
    myCycle = ((daysSince \ 28) Mod 13) + 1
    'Debug.Print "my Cycle: " & myCycle
    myCycle = leastSigTwoDigits(myCycle)
    If lastCycle = 0 Then
        lastCycle = myCycle
    Else
        If (myCycle > lastCycle) And (currYear <> lastYear) Then
            currYear = lastYear
        End If
    End If
    getVolume = currYear & leastSigTwoDigits(myCycle)
End Function

Private Function leastSigTwoDigits(ByVal myInt As Integer)
    Dim retStr As String
    
    myInt = myInt Mod 100
    
    If myInt < 10 Then
        leastSigTwoDigits = "0" & CStr(myInt)
    Else
        leastSigTwoDigits = CStr(myInt)
    End If
End Function

Private Sub UserForm_Terminate()
    If (userClicked = False) Then
        Unload Me
        'ThisWorkbook.Close False
        Application.DisplayAlerts = False
        Application.Quit
    End If
End Sub
