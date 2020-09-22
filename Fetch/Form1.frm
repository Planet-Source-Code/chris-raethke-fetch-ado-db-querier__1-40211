VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Fetch"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9045
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4890
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Results 
      Height          =   2775
      Left            =   60
      TabIndex        =   8
      Top             =   1020
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.CommandButton BuildCnn 
      Caption         =   "..."
      Height          =   315
      Left            =   8670
      TabIndex        =   1
      Top             =   360
      Width           =   285
   End
   Begin VB.TextBox CnnStr 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
   Begin VB.CommandButton FetchCmd 
      Caption         =   "FetchCmd"
      Height          =   375
      Left            =   7830
      TabIndex        =   3
      Top             =   4410
      Width           =   1095
   End
   Begin VB.TextBox SQLText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   60
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   990
      Width           =   8925
   End
   Begin VB.Label CallTime 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1830
      TabIndex        =   7
      Top             =   750
      Width           =   45
   End
   Begin VB.Label ShowResults 
      AutoSize        =   -1  'True
      Caption         =   "Results:"
      Height          =   195
      Left            =   930
      TabIndex        =   6
      Top             =   750
      Width           =   570
   End
   Begin VB.Label ShowSQL 
      AutoSize        =   -1  'True
      Caption         =   "SQL Text:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   750
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connection String:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   1305
   End
   Begin VB.Menu mnuREsults 
      Caption         =   "Results"
      Visible         =   0   'False
      Begin VB.Menu mnuREsultsCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuResultsSelectAll 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BuildCnn_Click()
On Error GoTo Errorhandler
    Dim str As String
    Dim x As MSDASC.DataLinks
    Set x = New MSDASC.DataLinks
    str = x.PromptNew
    If str <> "" Then
        CnnStr.Text = str
    End If
    
    Set x = Nothing
ExitSub:
    Exit Sub
    
Errorhandler:
    MsgBox FormatErr()
End Sub

Private Sub FetchCmd_Click()
On Error GoTo Errorhandler

    ' Set up Command and Connection objects
    Dim rs As ADODB.Recordset, cmd As ADODB.Command, field As ADODB.field
    Set rs = New ADODB.Recordset
    Set cmd = New ADODB.Command

    Dim tm1 As Single, tm2 As Single
    Dim ResultLine As String
    Dim first As Boolean
    Dim arr As Variant
    Dim colnum As Integer

    'Run the procedure
    If CnnStr.Text <> "" Then
        cmd.ActiveConnection = CnnStr.Text
    Else
        MsgBox "You must enter a valid ado connection string"
        CnnStr.SetFocus
    End If

    If SQLText.Text <> "" Then
        ' remember our last queries
        SaveSetting "Fetch", "Defaults", "SQLTEXT", SQLText.Text
        SaveSetting "Fetch", "Defaults", "CNNSTR", CnnStr.Text

        cmd.CommandText = SQLText.Text
        cmd.CommandType = adCmdText
        rs.CursorLocation = adUseClient
        tm1 = Timer
        rs.Open cmd, , adOpenStatic, adLockReadOnly
        tm2 = Timer
        CallTime.Caption = "Call Time: " & FormatNumber(tm2 - tm1, 8)

        ' show records in result listbox
        If Not rs.EOF Then

            ' change size of flexgrid
            Results.Cols = rs.Fields.Count
            Results.Rows = 1
'            colnum = 0
'            Results.ColWidth(0) = 10000

            ' add field names
            first = True
            For Each field In rs.Fields
                If first Then
                    ResultLine = "<" & field.Name  ' padstr(field.Name, field.DefinedSize)
                    first = False
                Else
                    ResultLine = ResultLine & "|<" & field.Name ' padstr(field.Name, field.DefinedSize)
                End If
            Next field
            Results.FormatString = ResultLine

            arr = rs.GetRows
            For x = LBound(arr, 2) To UBound(arr, 2)
                For y = LBound(arr) To UBound(arr)
                    If y = 0 Then
                        If IsNull(arr(y, x)) Then
                            ResultLine = "NULL"
                        Else
                            ResultLine = arr(y, x)
                        End If
                    Else
                        If IsNull(arr(y, x)) Then
                            ResultLine = ResultLine & Chr(9) & "NULL"
                        Else
                            ResultLine = ResultLine & Chr(9) & arr(y, x)
                        End If
                    End If
                    'DoEvents
                Next y

                ' add the line
                Results.AddItem ResultLine
                Results.Refresh
                ResultLine = ""
            Next x
            
            Results.ColWidth(0) = 1000
            For x = 1 To rs.Fields.Count - 1
                Results.ColWidth(x) = 1500
            Next x
            Call ShowResults_Click
        Else
            Results.Clear
            Results.AddItem "Recordset is Empty"
        End If
    Else
        MsgBox "You must enter a valid SQL query"
        SQLText.SetFocus
    End If

    Set rs = Nothing
    Set cmd = Nothing

ExitSub:
    Exit Sub

Errorhandler:
    MsgBox FormatErr()
End Sub

Private Sub Form_Load()
    SQLText.Text = GetSetting("Fetch", "Defaults", "SQLTEXT", "")
    CnnStr.Text = GetSetting("Fetch", "Defaults", "CNNSTR", "")
    Results.Visible = False
    SizeControls
End Sub

Sub SizeControls()
    On Error Resume Next
    FetchCmd.Top = Form1.Height - (FetchCmd.Height * 2) - 120
    FetchCmd.Left = Form1.Width - FetchCmd.Width - 240
    Results.Width = Form1.Width - 240
    Results.Height = Form1.Height - 2000
    SQLText.Width = Form1.Width - 240
    SQLText.Height = Form1.Height - 2000
End Sub

Private Function FormatErr()
    FormatErr = _
        "ERROR" & vbCrLf & _
        "Number: " & Err.Number & vbCrLf & _
        "Source: " & Err.Source & vbCrLf & _
        "Description: " & Err.Description
End Function

Function padstr(ByVal val As String, ByVal length As Integer) As String
    Do While Len(val) < length
        val = val & " "
    Loop
    padstr = val
End Function

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ShowSQL_Click
    If Data.GetFormat(vbCFFiles) Then
        SQLText.Text = ReadTextFile(Data.Files(1))
    ElseIf Data.GetFormat(vbCFText) Then
        SQLText.Text = Data.GetData(vbCFText)
    End If
End Sub

Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Effect = vbDropEffectCopy
End Sub

Private Sub Form_Resize()
    SizeControls
End Sub

Private Sub mnuResultsCopy_Click()
    EditCopy
End Sub

Private Sub mnuResultsSelectAll_Click()
    EditSelectAll
End Sub

Private Sub Results_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuREsults
    End If
End Sub

Private Sub Results_OLEDragDrop(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ShowSQL_Click
    If Data.GetFormat(vbCFFiles) Then
        SQLText.Text = ReadTextFile(Data.Files(1))
    ElseIf Data.GetFormat(vbCFText) Then
        SQLText.Text = Data.GetData(vbCFText)
    End If
End Sub

Private Sub Results_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Effect = vbDropEffectCopy
End Sub

Private Sub ShowResults_Click()
    SQLText.Visible = False
    Results.Visible = True
End Sub

Private Sub ShowSQL_Click()
    SQLText.Visible = True
    Results.Visible = False
End Sub

Function ReadTextFile(filepath)
    Dim fs As Scripting.FileSystemObject
    Set fs = New Scripting.FileSystemObject

    Dim f As Scripting.TextStream

    Set f = fs.OpenTextFile(filepath, ForReading)
    ReadTextFile = f.ReadAll
    f.Close

    Set f = Nothing
    Set fs = Nothing
End Function

Private Sub SQLText_KeyDown(KeyCode As Integer, Shift As Integer)
'    If (Shift = vbCtrlMask) And (KeyCode = Asc("a") Or KeyCode = Asc("A")) Then
'        SQLText.SelStart = 0
 '       SQLText.SelLength = Len(SQLText.Text)
 '   End If
End Sub

Private Sub SQLText_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Data.GetFormat(vbCFFiles) Then
        SQLText.Text = ReadTextFile(Data.Files(1))
    ElseIf Data.GetFormat(vbCFText) Then
        SQLText.Text = Data.GetData(vbCFText)
    End If
End Sub

Private Sub SQLText_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Effect = vbDropEffectCopy
End Sub


Private Sub EditCut()
    'Cut the selection and put it on the Clipboard
    EditCopy
    EditDelete
End Sub

Private Sub EditCopy()
    'Copy the selection and put it on the Clipboard
    Clipboard.Clear
    Clipboard.SetText Results.Clip
End Sub

Private Sub EditPaste()
    'Insert Clipboard contents
    If Len(Clipboard.GetText) Then Results.Clip = _
         Clipboard.GetText
End Sub

Private Sub EditDelete()
    'Deletes the selection
    Dim i As Integer
    Dim j As Integer
    Dim strClip As String
    With Results
        For i = 1 To .RowSel
            For j = 1 To .ColSel
                strClip = strClip & "" & vbTab
            Next
            strClip = strClip & vbCr
        Next
        .Clip = strClip
    End With
End Sub

Private Sub EditSelectAll()
    'Selects the whole Grid
    With Results
        .Visible = False
        If .Rows > 1 And .Cols > 1 Then
            .Row = 1
            .Col = 0
            .RowSel = .Rows - 1
            .ColSel = .Cols - 1
            .TopRow = 1
        End If
        .Visible = True
    End With
End Sub

