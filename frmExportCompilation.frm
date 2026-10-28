VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportCompilation 
   Caption         =   "QARS - Export Compilation Wizard"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "frmExportCompilation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportCompilation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With Me.cboUnit
        .Clear
        .AddItem "Manual Payment"
        .AddItem "OBPS"
        .AddItem "E-Channel"
        .AddItem "Tax Settlement"
        .AddItem "E-Tax"
        .AddItem "PBI"
        .AddItem "UPILS"
        .AddItem "RLP"
        .ListIndex = 0
    End With

    On Error Resume Next
    With Me.cboQAStatus
        .Clear
        .AddItem "Done (Selesai Audit QA)"
        .AddItem "Pending (Menunggu Review QA)"
        .AddItem "Done & Pending (Semua QA)"
        .AddItem "Belum Masuk QA (Unassigned)"
        .AddItem "All Data (Keseluruhan Transaksi)"
        .ListIndex = 0
    End With
    On Error GoTo 0

    SetDefaultDateThisMonth

    On Error Resume Next
    Me.lblStatus.Caption = "Ready to export data."
    On Error GoTo 0
End Sub

Private Sub btnPickStart_Click()
    Dim selectedDate As Date
    On Error Resume Next
    selectedDate = CalendarForm.GetDate
    On Error GoTo 0

    If selectedDate <> 0 Then
        Me.txtStartDate.Text = Format$(selectedDate, "yyyy-mm-dd")
    End If
End Sub

Private Sub btnPickEnd_Click()
    Dim selectedDate As Date
    On Error Resume Next
    selectedDate = CalendarForm.GetDate
    On Error GoTo 0

    If selectedDate <> 0 Then
        Me.txtEndDate.Text = Format$(selectedDate, "yyyy-mm-dd")
    End If
End Sub

Private Sub btnThisMonth_Click()
    SetDefaultDateThisMonth
End Sub

Private Sub btnToday_Click()
    Me.txtStartDate.Text = Format$(Date, "yyyy-mm-dd")
    Me.txtEndDate.Text = Format$(Date, "yyyy-mm-dd")
End Sub

Private Sub SetDefaultDateThisMonth()
    Dim firstDay As Date
    firstDay = DateSerial(Year(Date), Month(Date), 1)
    Me.txtStartDate.Text = Format$(firstDay, "yyyy-mm-dd")
    Me.txtEndDate.Text = Format$(Date, "yyyy-mm-dd")
End Sub

Private Sub btnExportFile_Click()
    ExecuteExportAction True
End Sub

Private Sub btnPullSheet_Click()
    ExecuteExportAction False
End Sub

Private Sub ExecuteExportAction(ByVal exportToNewFile As Boolean)
    Dim selectedUnit As String
    Dim selectedStatus As String
    Dim strStart As String
    Dim strEnd As String
    Dim dStart As Date
    Dim dEnd As Date
    Dim targetSampling As Double
    Dim recCount As Long

    selectedUnit = Trim$(Me.cboUnit.Text)
    If Len(selectedUnit) = 0 Then
        MsgBox "Please select a Unit first!", vbExclamation, "Select Unit"
        Me.cboUnit.SetFocus
        Exit Sub
    End If

    selectedStatus = "DONE"
    On Error Resume Next
    If Me.cboQAStatus.ListIndex >= 0 Then
        Select Case Me.cboQAStatus.ListIndex
            Case 0: selectedStatus = "DONE"
            Case 1: selectedStatus = "PENDING"
            Case 2: selectedStatus = "QA_ALL"
            Case 3: selectedStatus = "UNASSIGNED"
            Case 4: selectedStatus = "ALL"
            Case Else: selectedStatus = "DONE"
        End Select
    End If
    On Error GoTo 0

    strStart = Trim$(Me.txtStartDate.Text)
    If Not IsDate(strStart) Then
        MsgBox "Invalid Start Date format! Use YYYY-MM-DD or the calendar picker.", vbExclamation, "Invalid Date"
        Me.txtStartDate.SetFocus
        Exit Sub
    End If
    dStart = CDate(strStart)

    strEnd = Trim$(Me.txtEndDate.Text)
    If Not IsDate(strEnd) Then
        MsgBox "Invalid End Date format! Use YYYY-MM-DD or the calendar picker.", vbExclamation, "Invalid Date"
        Me.txtEndDate.SetFocus
        Exit Sub
    End If
    dEnd = CDate(strEnd)

    If dEnd < dStart Then
        MsgBox "End Date cannot be earlier than Start Date!", vbExclamation, "Invalid Date Range"
        Me.txtEndDate.SetFocus
        Exit Sub
    End If

    targetSampling = 0
    On Error Resume Next
    If Len(Trim$(Me.txtTargetSampling.Text)) > 0 Then
        targetSampling = CDbl(Me.txtTargetSampling.Text)
    End If
    On Error GoTo 0

    Me.lblStatus.Caption = "Processing " & selectedStatus & " export for " & selectedUnit & "..."
    DoEvents

    Me.btnExportFile.Enabled = False
    Me.btnPullSheet.Enabled = False
    Me.btnClose.Enabled = False

    On Error GoTo ErrHandler

    recCount = modExportCompilation.ExecuteExportEngine(selectedUnit, dStart, dEnd, targetSampling, exportToNewFile, selectedStatus)

    Me.btnExportFile.Enabled = True
    Me.btnPullSheet.Enabled = True
    Me.btnClose.Enabled = True

    If recCount >= 0 Then
        Me.lblStatus.Caption = "Complete! Successfully exported " & recCount & " rows of data (" & selectedStatus & ")."
    Else
        Me.lblStatus.Caption = "Process cancelled or an error occurred."
    End If

    Exit Sub

ErrHandler:
    Me.btnExportFile.Enabled = True
    Me.btnPullSheet.Enabled = True
    Me.btnClose.Enabled = True
    Me.lblStatus.Caption = "An error occurred during export process."
    MsgBox "An error occurred while exporting: " & Err.Description, vbCritical, "Export Error"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
