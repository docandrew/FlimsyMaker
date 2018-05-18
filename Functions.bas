Attribute VB_Name = "Functions"
Option Explicit

Const MaxWidth As Integer = 420
Global cycleData As String
Global cycleExpire As String
Global printPageNums As Boolean
Global printExpiration As Boolean
Global printTOC As Boolean

Sub HideWB()
    Application.Visible = False
    frmMain.Show
End Sub

Public Sub ChangeAllWSVis(ByVal hideWS As Boolean)
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If (ws.name <> "Launch Flimsy Maker") Then
            If (hideWS = True) Then
                ws.Visible = xlSheetHidden
            Else
                ws.Visible = xlSheetVisible
            End If
        End If
    Next ws
End Sub

Public Sub StartCopyOperation()
    frmProgress.Show
End Sub

Private Function FileExists(ByVal file As String) As Boolean
    FileExists = (Dir(file) <> "")
End Function

Private Sub DeleteFile(ByVal file As String)
    If (FileExists(file)) Then
        SetAttr file, vbNormal
        Kill file
    End If
End Sub

Public Function combineFiles(ByVal curJob As Integer, ByVal totalJobs As Integer, ByVal fileLoc As String, ByVal saveLoc As String) As String
    Dim app As Acrobat.CAcroApp
    Dim acroSource As Acrobat.CAcroPDDoc
    Dim acroDest As Acrobat.CAcroPDDoc
    Dim allFiles As Collection
    Dim jso As Object
    Dim i As Integer
    Dim j As Integer
    Dim saveAs As String
    Dim numPages As Integer
    Dim tempName As String
    Dim expireText As String
    
    Set allFiles = New Collection
    Application.ScreenUpdating = False
    expireText = "Expired as of " & cycleExpire
    saveAs = saveLoc & cycleExpire & ".pdf"
    
    If (printTOC = True) Then
        Call IterateProgress(curJob, totalJobs, "Creating Table of Contents")
        Call GenerateTOC(fileLoc + "000_toc.pdf") ', app)
        curJob = curJob + 1
        Call killAdobeProcess
    End If
    
    Call getAllFiles(fileLoc, allFiles)
    If (allFiles.count > 0) Then
        Set app = CreateObject("AcroExch.App")
        Set acroSource = CreateObject("AcroExch.PDDoc")
        Set acroDest = CreateObject("AcroExch.PDDoc")
        
        'curJob = curJob + 1
        Call IterateProgress(curJob, totalJobs, "Setting up PDF")
        acroDest.Open (fileLoc & allFiles(1))
        For i = 2 To allFiles.count Step 1
            curJob = curJob + 1
            Call IterateProgress(curJob, totalJobs, "Adding " & allFiles(i) & " to PDF doc")
            numPages = acroDest.GetNumPages()
            acroSource.Open (fileLoc & allFiles(i))
            If (acroDest.InsertPages(numPages - 1, acroSource, 0, acroSource.GetNumPages(), True)) Then
                'works
            Else
                Call MsgBox("Unable to add " & allFiles(i), vbOKOnly + vbInformation, "File Add Error")
            End If
            acroSource.Close
        Next i
        curJob = curJob + 1
        Call IterateProgress(curJob, totalJobs, "Finishing up PDF merge")
        
        If (printExpiration = True Or printPageNums = True) Then
            Set jso = acroDest.GetJSObject
            If Not jso Is Nothing Then
                If (printExpiration = True) Then
                    Call IterateProgress(curJob, totalJobs, "Printing expiration dates")
                    Call jso.addWatermarkFromText(expireText, jso.app.Constants.Align.center, "Arial", 24, jso.Color.red, 0, jso.numPages, True, True, True, jso.app.Constants.Align.center, jso.app.Constants.Align.bottom, 0, 0, False, 1#, False, 0, 1#)
                    Call jso.addWatermarkFromText(expireText, jso.app.Constants.Align.center, "Arial", 24, jso.Color.red, 0, jso.numPages, True, True, True, jso.app.Constants.Align.center, jso.app.Constants.Align.Top, 0, 0, False, 1#, False, 0, 1#)
                    curJob = curJob + 1
                End If
                If (printPageNums = True) Then
                    'If (printTOC = True) Then
                    '    j = 2
                    'Else
                    '    j = 1
                    'End If
                    If (printTOC = True) Then
                        For i = 1 To jso.numPages - 1 Step 1
                            Call IterateProgress(curJob, totalJobs, "Printing page number on page " & CStr(i))
                            Call jso.addWatermarkFromText(CStr(i) & "  ", jso.app.Constants.Align.center, "Arial", 24, jso.Color.red, i, i, True, True, True, jso.app.Constants.Align.Right, jso.app.Constants.Align.bottom, 0, 0, False, 1#, False, 0, 1#)
                            curJob = curJob + 1
                        Next i
                    Else
                        For i = 1 To jso.numPages Step 1
                            Call IterateProgress(curJob, totalJobs, "Printing page number on page " & CStr(i))
                            Call jso.addWatermarkFromText(CStr(i) & "  ", jso.app.Constants.Align.center, "Arial", 24, jso.Color.red, i - 1, i - 1, True, True, True, jso.app.Constants.Align.Right, jso.app.Constants.Align.bottom, 0, 0, False, 1#, False, 0, 1#)
                            curJob = curJob + 1
                        Next i
                    End If
                End If
            End If
        End If
        Call IterateProgress(curJob, totalJobs, "Saving PDF document")
        acroDest.Save PDSaveFull, saveAs
        app.CloseAllDocs
        acroDest.Close
        app.Exit
        Set acroSource = Nothing
        Set acroDest = Nothing
        Set app = Nothing
        Call killAdobeProcess
        curJob = curJob + 1
        Call IterateProgress(curJob, totalJobs, "Cleaning up files")
        Call clearDirectory(fileLoc, allFiles)
        curJob = curJob + 1
        Call IterateProgress(curJob, totalJobs, "Finishing up...")
        combineFiles = saveAs
    Else
        Call MsgBox("No files found in that directory", vbOKOnly + vbCritical, "Error")
        combineFiles = ""
    End If
End Function

Public Sub killAdobeProcess()
    Call doSomethingToAdobe(True)
End Sub

Public Function checkIfAdobeExists()
    Dim toDie As Boolean
    toDie = doSomethingToAdobe(False)
    checkIfAdobeExists = toDie
End Function

Private Function doSomethingToAdobe(ByVal closeIt As Boolean) As Boolean
Dim oServ As Object
Dim cProc As Variant
Dim oProc As Object

Set oServ = GetObject("winmgmts:")
Set cProc = oServ.ExecQuery("Select * from Win32_Process")
For Each oProc In cProc

    If LCase(oProc.name) = "acrobat.exe" Then
        If (closeIt = True) Then
            Call oProc.Terminate
            doSomethingToAdobe = False
            Exit Function
        Else
            Call MsgBox("You must close all instances of Adobe Acrobat before running this program.", vbOKOnly + vbInformation, "Close Adobe")
            doSomethingToAdobe = True
            Exit Function
        End If
    End If
Next

doSomethingToAdobe = False
End Function

Private Sub clearDirectory(ByVal rmDir As String, ByVal fileCol As Collection)
    Dim i As Integer
    For i = 1 To fileCol.count Step 1
        Call DeleteFile(rmDir & fileCol(i))
    Next i
End Sub

Private Function getAllFiles(ByVal loc As String, ByVal fileCol As Collection) As Variant
    Dim count As Integer
    Dim name As String
    
    On Error GoTo NoFilesFound
    
    count = 0
    name = Dir(loc)
    If name = "" Then GoTo NoFilesFound
    Do While (name <> "")
        fileCol.Add (name)
        name = Dir()
    Loop
    Exit Function
    
NoFilesFound:

End Function

Public Sub IterateProgress(ByVal currentJob As Integer, ByVal totalJobs As Integer, ByVal jobInfo As String)
    Dim percentDone As Double
    Dim newWidth As Integer
    Dim displayText As String
    Const minWidth As Integer = 96
    
    percentDone = (CDbl(currentJob) / CDbl(totalJobs))
    'frmProgress.lblBar.Caption = Format(percentDone, "Percent") 'CStr(percent * 100) & "%"
    frmProgress.lblPercent.Caption = Format(percentDone, "Percent") 'CStr(percent * 100) & "%"
    newWidth = CInt(percentDone * MaxWidth)
    If (newWidth > minWidth) Then
        frmProgress.lblPercent.Width = newWidth
    Else
        frmProgress.lblPercent.Width = minWidth
    End If
    
    frmProgress.lblBar.Width = newWidth
    displayText = jobInfo & ", completing job " & CStr(currentJob) & " of " & CStr(totalJobs) & " jobs"
    frmProgress.lblInfo.Caption = displayText
    DoEvents
End Sub

Public Sub IterateProgressTPP(ByVal downloaded As Long, ByVal totalDownload As Long)
    Dim percentDone As Double
    Dim newWidth As Integer
    Dim display As String
    Const minWidth As Integer = 96
    percentDone = (CDbl(downloaded) / CDbl(totalDownload))
    frmProgressTPP.lblPercent.Caption = Format(percentDone, "Percent")
    newWidth = CInt(percentDone * MaxWidth)
    If (newWidth > minWidth) Then
        frmProgressTPP.lblPercent.Width = newWidth
    Else
        frmProgressTPP.lblPercent.Width = minWidth
    End If
    frmProgressTPP.lblBar.Width = newWidth
    Dim downloadAmt As String
    Dim totDownload As String
    downloadAmt = Format((CDbl(downloaded) / 1024), "Standard") & " KB"
    totDownload = Format((CDbl(totalDownload) / 1024), "Standard") & " KB"
    frmProgressTPP.lblInfo.Caption = "Downloaded " & downloadAmt & " of " & totDownload
    DoEvents
End Sub

Public Sub DownloadFile(ByVal source As String, ByVal dest As String)
    Call URLDownloadToFile(0, source, dest, 0, 0)
End Sub

Public Function GetFileSize(ByVal URL As String) As Long
Dim oXHTTP As Object

    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    oXHTTP.Open "HEAD", URL, False
    oXHTTP.send
    If (oXHTTP.Status = 200) Then 'http ok
        GetFileSize = oXHTTP.getResponseHeader("Content-Length")
    Else
        GetFileSize = -1
    End If
End Function

Sub Sort(ByRef arr() As String)
  Dim strTemp As String
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  
  lngMin = LBound(arr)
  lngMax = UBound(arr)
  
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If arr(i) > arr(j) Then
        strTemp = arr(i)
        arr(i) = arr(j)
        arr(j) = strTemp
      End If
    Next j
  Next i
End Sub

Public Sub ShowExcel()
    Application.Visible = True
End Sub
