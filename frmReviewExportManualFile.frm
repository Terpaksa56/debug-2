VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReviewExportManualFile 
   Caption         =   "EXPORT QA REVIEW - MANUAL FILE"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565.001
   OleObjectBlob   =   "frmReviewExportManualFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReviewExportManualFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadHeaderLogo()
    On Error Resume Next
    modUtility.ApplyHeaderLogo Me.imgHeaderLogo
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    LoadHeaderLogo
    
    ' Populate Years
    cmbYear.Clear
    Dim y As Long
    For y = 2024 To 2028
        cmbYear.AddItem CStr(y)
    Next y
    
    Dim curYr As Long
    curYr = Year(Date)
    If curYr < 2024 Or curYr > 2028 Then curYr = 2026
    cmbYear.value = CStr(curYr)
    
    ' Populate Months
    cmbMonth.Clear
    cmbMonth.AddItem "1 - January"
    cmbMonth.AddItem "2 - February"
    cmbMonth.AddItem "3 - March"
    cmbMonth.AddItem "4 - April"
    cmbMonth.AddItem "5 - May"
    cmbMonth.AddItem "6 - June"
    cmbMonth.AddItem "7 - July"
    cmbMonth.AddItem "8 - August"
    cmbMonth.AddItem "9 - September"
    cmbMonth.AddItem "10 - October"
    cmbMonth.AddItem "11 - November"
    cmbMonth.AddItem "12 - December"
    
    Dim curMo As Long
    curMo = Month(Date)
    If curMo < 1 Or curMo > 12 Then curMo = 8
    cmbMonth.ListIndex = curMo - 1
    
    lblStatus.Caption = ""
    lblFolderPath.Caption = modDB.GetQAReviewFolder()
    
    RefreshPreview
End Sub

Private Sub cmbYear_Change()
    RefreshPreview
End Sub

Private Sub cmbMonth_Change()
    RefreshPreview
End Sub

Public Sub RefreshPreview()
    On Error GoTo ErrHandler
    Dim yr As Long, mo As Long
    yr = val(cmbYear.value)
    mo = cmbMonth.ListIndex + 1
    
    If yr <= 2020 Or mo < 1 Or mo > 12 Then
        lblRecordStatus.Caption = "Please select a valid year and month."
        lblRecordStatus.ForeColor = RGB(120, 120, 120)
        lblTargetFile.Caption = "-"
        btnExport.Enabled = False
        Exit Sub
    End If
    
    ' Detect Tool Name (NS or NNS)
    Dim isNNS As Boolean
    isNNS = (InStr(1, ThisWorkbook.Name, "NNS", vbTextCompare) > 0)
    
    Dim monthAbbr As String, yrShort As String
    monthAbbr = Format(DateSerial(yr, mo, 1), "mmm")
    yrShort = Right$(CStr(yr), 2)
    
    Dim fName As String
    If isNNS Then
        fName = "Indonesia_NNS_Manual_File_" & monthAbbr & "'" & yrShort & ".xlsx"
    Else
        fName = "Indonesia_NS_Manual_File_" & monthAbbr & "'" & yrShort & ".xlsx"
    End If
    lblTargetFile.Caption = fName
    
    ' Count available records from database per month
    Dim sql As String, cntVal As Variant
    sql = "SELECT COUNT(*) FROM qa_reviews WHERE year_no = " & yr & " AND month_no = " & mo
    cntVal = modDB.ExecuteScalar(sql)
    
    Dim recCount As Long
    recCount = CLng(val(cntVal))
    
    If recCount > 0 Then
        lblRecordStatus.Caption = "Found " & recCount & " QA Review record(s) ready to export."
        lblRecordStatus.ForeColor = RGB(0, 128, 0)
        btnExport.Enabled = True
    Else
        lblRecordStatus.Caption = "No QA Review records found for " & monthAbbr & " " & yr & "."
        lblRecordStatus.ForeColor = RGB(200, 30, 30)
        btnExport.Enabled = False
    End If
    
    lblFolderPath.Caption = modDB.GetQAReviewFolder()
    Exit Sub
ErrHandler:
    lblRecordStatus.Caption = "Status: " & Err.description
    lblRecordStatus.ForeColor = RGB(200, 30, 30)
End Sub

Private Sub btnExport_Click()
    On Error GoTo ErrHandler
    Dim yr As Long, mo As Long
    yr = val(cmbYear.value)
    mo = cmbMonth.ListIndex + 1
    
    If yr <= 2020 Or mo < 1 Or mo > 12 Then
        MsgBox "Please select a valid Year and Month.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    lblStatus.Caption = "Exporting data, please wait..."
    lblStatus.ForeColor = RGB(0, 51, 102)
    DoEvents
    
    Dim ok As Boolean
    ' Export full monthly QA Review data (weekNum = 0)
    ok = modExport.ExportQAReviewToManualFile(CInt(yr), CInt(mo), 0)
    
    If ok Then
        lblStatus.Caption = "Monthly export completed successfully!"
        lblStatus.ForeColor = RGB(0, 128, 0)
        RefreshPreview
    Else
        lblStatus.Caption = "Export canceled or no data."
        lblStatus.ForeColor = RGB(200, 30, 30)
    End If
    Exit Sub
ErrHandler:
    lblStatus.Caption = "Error: " & Err.description
    lblStatus.ForeColor = RGB(200, 30, 30)
    MsgBox "Failed to export: " & Err.description, vbCritical, "Export Error"
End Sub

Private Sub btnOpenFolder_Click()
    On Error Resume Next
    Dim folderPath As String
    folderPath = modDB.GetQAReviewFolder()
    If Trim$(folderPath) <> "" Then
        Shell "explorer.exe """ & folderPath & """", vbNormalFocus
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
