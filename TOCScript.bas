Attribute VB_Name = "TOCScript"
Public Sub GenerateTOC(ByVal saveLoc As String) ', ByVal app As Acrobat.AcroApp)
    Dim myApp As Acrobat.AcroApp
    Dim myDoc As Acrobat.AcroPDDoc
    Dim myPage As Acrobat.AcroPDPage
    Dim docObject As Object
    Dim jso As Object
    Dim reportGen As String
    Dim reportText As String
    Dim currICAO As String
    Dim lastICAO As String
    Dim temp As String
    Dim i As Integer
    Dim pageNum As Integer
    
    Set myApp = CreateObject("AcroExch.App")
    Set myDoc = CreateObject("AcroExch.PDDoc")
    myDoc.Create
    Set docObject = myDoc.GetJSObject
    
    reportGen = "var rep = new Report(); rep.size=1.5; rep.color=color.black; "
    reportText = reportGen + "rep.indent(250); rep.writeText(""Flimsy TOC""); rep.outdent(250); "
    
    pageNum = 1
    currICAO = ""
    Dim icao As String
    Dim appr As String
    For i = 0 To frmMain.YourFlimsyListBox.ListCount - 1 Step 1
        icao = frmMain.YourFlimsyListBox.List(i, 0)
        appr = frmMain.YourFlimsyListBox.List(i, 1)
        If (icao <> currICAO) Then
            currICAO = icao
            reportText = reportText + "rep.writeText(""" + currICAO + """); "
        End If
        If (appr <> "") Then
            reportText = reportText + "rep.indent(20); rep.writeText(""" + CStr(pageNum) + ".  " + vbTab + appr + """); rep.outdent(20); "
        End If
        pageNum = pageNum + 1
    Next i
    'Dim saveLoc As String
    'saveLoc = Application.ActiveWorkbook.Path + "\TOC.pdf"
    reportText = reportText + "rep.open(""My Report"");"
    Debug.Print reportText
    Call docObject.addScript("myScript", reportText)
    Dim toSaveDoc As Acrobat.AcroAVDoc
    Set toSaveDoc = myApp.GetActiveDoc
    Set myDoc = toSaveDoc.GetPDDoc
    myDoc.Save PDSaveFull, saveLoc
    myDoc.Close
    myApp.CloseAllDocs
    myApp.Exit
End Sub
