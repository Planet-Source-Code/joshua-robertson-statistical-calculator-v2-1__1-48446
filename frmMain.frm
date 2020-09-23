VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMain"
   ClientHeight    =   6510
   ClientLeft      =   75
   ClientTop       =   2595
   ClientWidth     =   4455
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   1515
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   2775
      TabIndex        =   25
      Top             =   900
      Width           =   2775
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Double-click the rows in this grid to modify data."
         Height          =   435
         Left            =   120
         TabIndex        =   26
         Top             =   60
         Width           =   2595
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   3855
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   2580
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   16
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   0
      SelectionMode   =   1
      BorderStyle     =   0
   End
   Begin VB.CheckBox chkPlotGraph 
      Caption         =   "Plot Graph"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3000
      TabIndex        =   23
      Top             =   4260
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtSolve 
      Height          =   285
      Left            =   2940
      TabIndex        =   16
      Top             =   5220
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   315
      Left            =   2940
      TabIndex        =   15
      Top             =   2100
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As..."
      Height          =   315
      Left            =   2940
      TabIndex        =   6
      Top             =   1740
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open..."
      Height          =   315
      Left            =   2940
      TabIndex        =   5
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Regression"
      Height          =   1335
      Left            =   2940
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   60
         ScaleHeight     =   1035
         ScaleWidth      =   1335
         TabIndex        =   21
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton optRegressionModel 
            Caption         =   "Option1"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton cmdSolveForY 
      Caption         =   "Solve for y"
      Height          =   315
      Left            =   2940
      TabIndex        =   12
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSolveForX 
      Caption         =   "Solve for x"
      Height          =   315
      Left            =   2940
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2940
      TabIndex        =   10
      Top             =   3900
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Data"
      Default         =   -1  'True
      Height          =   315
      Left            =   2940
      TabIndex        =   3
      Top             =   660
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Row(s)"
      Height          =   315
      Left            =   2940
      TabIndex        =   4
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2940
      TabIndex        =   2
      Top             =   300
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   1
      Top             =   300
      Width           =   1395
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2940
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   1755
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   660
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3096
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
      BorderStyle     =   0
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   195
      Left            =   2940
      TabIndex        =   20
      Top             =   4740
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   195
      Left            =   2940
      TabIndex        =   19
      Top             =   4500
      Width           =   1455
   End
   Begin VB.Label lblSolveOut 
      Height          =   195
      Left            =   2940
      TabIndex        =   18
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label lblSolveFor 
      Caption         =   "Enter value here:"
      Height          =   195
      Left            =   2940
      TabIndex        =   17
      Top             =   4980
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "frequency"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "y"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "x"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   1395
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFileDash 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Double
Dim b As Double
Dim RegressionModeChanged As Boolean
Dim DataSaved As Boolean
Dim ModifyData As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub chkPlotGraph_Click()
    'Code not implemented yet, sorry.
End Sub

Private Sub Form_Initialize()
    Call InitCommonControls
End Sub

Private Sub CmdNew_Click()
    Dim Msg As VbMsgBoxResult
    Dim q As Integer
    If Not DataSaved Then
        Msg = MsgBox("The existing data has not been saved. Save Data first?", vbYesNoCancel + vbQuestion)
        Select Case Msg
            Case vbCancel
                Exit Sub
            Case vbYes
                cmdSave_Click
                If Not DataSaved Then Exit Sub
        End Select
    End If
    MSFlexGrid(0).Rows = 1
    RegressionModeChanged = True
    MSFlexGrid(0).ZOrder 1
    ModifyDataMode (False)
    For q = 0 To 2
        Text1(q) = ""
    Next
    DataSaved = True
    txtSolve = ""
    lblSolveOut = ""
    For q = 1 To 15
        MSFlexGrid(1).TextArray(Fgi(q, 1, MSFlexGrid(1))) = ""
    Next
End Sub

Private Sub cmdRemove_Click()
    Dim q As Long
    Dim w As Long
    If MSFlexGrid(0).Rows = 2 Then
        MSFlexGrid(0).FixedRows = 1
        MSFlexGrid(0).Rows = 1
        MSFlexGrid(0).ZOrder 1
        ModifyDataMode (False)
        For q = 0 To 2
            Text1(q) = ""
        Next
        DataSaved = True
    ElseIf MSFlexGrid(0).Rows > 2 Then
        If MSFlexGrid(0).Row > MSFlexGrid(0).RowSel Then
            For q = MSFlexGrid(0).RowSel To MSFlexGrid(0).Row
                MSFlexGrid(0).RemoveItem (q - w)
                w = w + 1
            Next
        ElseIf MSFlexGrid(0).Row < MSFlexGrid(0).RowSel Then
            For q = MSFlexGrid(0).Row To MSFlexGrid(0).RowSel
                MSFlexGrid(0).RemoveItem (q - w)
                w = w + 1
            Next
        Else
            MSFlexGrid(0).RemoveItem (MSFlexGrid(0).RowSel)
        End If
        ModifyDataMode (False)
        For q = 0 To 2
            Text1(q) = ""
        Next
        DataSaved = False
    End If
End Sub

Private Sub cmdSolveForX_Click()
    Dim index As Integer
    If RegressionModeChanged Then
        MsgBox "Press ""Calculate"" first before solving for X.", vbExclamation
    Else
        If IsNumeric(txtSolve) Then
            For index = 0 To optRegressionModel.Count
                If optRegressionModel(index).Value Then Exit For
            Next
            On Error GoTo ErrorInXCalc
            Select Case index
                Case 0
                    lblSolveOut = "x = " & CSng((txtSolve - a) / b)
                Case 1
                    lblSolveOut = "x = " & CSng(Exp((txtSolve - a) / b))
                Case 2
                    lblSolveOut = "x = " & CSng(Ln(txtSolve / a) / b)
                Case 3
                    lblSolveOut = "x = " & CSng((txtSolve / a) ^ (1 / b))
            End Select
            txtSolve.SetFocus
        Else
            MsgBox "A value must be entered in the ""Enter value here:"" box.", vbExclamation
        End If
    End If
    Exit Sub
ErrorInXCalc:
    lblSolveOut = "<" & Err.Description & ">"
    'Debug.Print "Run-time error '" & Err.Number & "': " & vbCrLf & vbCrLf & _
            Err.Description
    Resume Next
End Sub

Private Sub cmdSolveForY_Click()
    Dim index As Integer
    If RegressionModeChanged Then
        MsgBox "Press ""Calculate"" first before solving for Y.", vbExclamation
    Else
        If IsNumeric(txtSolve) Then
            For index = 0 To optRegressionModel.Count
                If optRegressionModel(index).Value Then Exit For
            Next
            On Error GoTo ErrorInYCalc
            Select Case index
                Case 0
                    lblSolveOut = "y = " & CSng(a + b * txtSolve)
                Case 1
                    lblSolveOut = "y = " & CSng(a + b * Ln(txtSolve))
                Case 2
                    lblSolveOut = "y = " & CSng(a * Exp(b * txtSolve))
                Case 3
                    lblSolveOut = "y = " & CSng(a * txtSolve ^ b)
            End Select
            txtSolve.SetFocus
        Else
            MsgBox "A value must be entered in the ""Enter value here:"" box.", vbExclamation
        End If
    End If
    Exit Sub
ErrorInYCalc:
    lblSolveOut = "<" & Err.Description & ">"
    'Debug.Print "Run-time error '" & Err.Number & "': " & vbCrLf & vbCrLf & _
            Err.Description
    Resume Next
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Response As Integer
    If Not DataSaved Then
        Response = MsgBox("Save Data before closing?", vbQuestion + vbYesNoCancel, "Save Dialog")
        Select Case Response
            Case vbCancel   ' Don't allow close.
                Cancel = -1
            Case vbYes
                cmdSave_Click
                If Not DataSaved Then Cancel = -1
        End Select
    End If
End Sub

Private Sub mnuFileNew_Click()
    CmdNew_Click
End Sub

Private Sub mnuFileOpen_Click()
    cmdOpen_Click
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    cmdSave_Click
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub MSFlexGrid_DblClick(index As Integer)
    Dim q As Integer
    Select Case index
        Case 0
            If MSFlexGrid(index).Rows > 1 Then
                ModifyDataMode (Not ModifyData)
                If ModifyData Then
                    For q = 0 To 2
                        Text1(q) = MSFlexGrid(index).TextArray(Fgi(MSFlexGrid(index).RowSel, q, MSFlexGrid(index)))
                    Next
                Else
                    For q = 0 To 2
                        Text1(q) = ""
                    Next
                End If
            End If
    End Select
End Sub

Private Sub optRegressionModel_Click(index As Integer)
    Select Case index
        Case 0
            Label5 = "x = (y - a) / b"
            Label6 = "y = a + b * x"
        Case 1
            Label5 = "x = Exp((y - a) / b)"
            Label6 = "y = a + b * Ln(x)"
        Case 2
            Label5 = "x = Ln(y / a) / b"
            Label6 = "y = a * Exp(b * x)"
        Case 3
            Label5 = "x = (y / a) ^ (1 / b)"
            Label6 = "y = a * x ^ b"
    End Select
    RegressionModeChanged = True
End Sub

Private Sub Text1_GotFocus(index As Integer)
    Text1(index).SelStart = 0
    Text1(index).SelLength = Len(Text1(index))
End Sub

Private Sub cmdAdd_Click()
    Dim q As Integer
    If ModifyData Then
        If IsNumeric(Text1(0)) And IsNumeric(Text1(1)) And IsNumeric(Text1(2)) Then
            MSFlexGrid(0).ZOrder
            ModifyDataMode (False)
            Text1(0).SetFocus
            DataSaved = False
            For q = 0 To 2
                MSFlexGrid(0).TextArray(Fgi(MSFlexGrid(0).RowSel, q, MSFlexGrid(0))) = Text1(q)
                Text1(q) = ""
            Next
        Else
            MsgBox "Please enter numeric values in the data entry boxes.", vbExclamation
        End If
        Exit Sub
    Else
        If IsNumeric(Text1(0)) And IsNumeric(Text1(1)) And IsNumeric(Text1(2)) Then
            MSFlexGrid(0).AddItem Text1(0) & Chr(9) & Text1(1) & Chr(9) & Text1(2)
            MSFlexGrid(0).ZOrder
            ModifyDataMode (False)
            Text1(0).SetFocus
            DataSaved = False
            For q = 0 To 2
                Text1(q) = ""
            Next
        Else
            MsgBox "Please enter numeric values in the data entry boxes.", vbExclamation
        End If
    End If
End Sub

Private Sub cmdCalculate_Click()
    Dim n(15) As Double
    Dim x() As Double
    Dim y() As Double
    Dim f() As Double
    Dim q As Integer
    Dim w As Integer
    Dim m As Integer
    Dim h As Double
    Dim p As Double
    Dim o As Double
    
    If MSFlexGrid(0).Rows = 1 Then
        MsgBox "There's nothing to compute. Either enter data or load data from a CSV file.", vbExclamation
        Exit Sub
    End If
    
    RegressionModeChanged = False 'because calculate has been pressed.
    m = MSFlexGrid(0).Rows - 1
    If m <= 1 Then
        MsgBox "You must have more than one set of data. Calculculation cannot continue.", vbExclamation
        Exit Sub
    End If
    
    ReDim x(m), y(m), f(m)
    
    For q = MSFlexGrid(0).FixedRows To MSFlexGrid(0).Rows - 1
        x(q) = MSFlexGrid(0).TextArray(Fgi(q, 0, MSFlexGrid(0)))
        y(q) = MSFlexGrid(0).TextArray(Fgi(q, 1, MSFlexGrid(0)))
        f(q) = MSFlexGrid(0).TextArray(Fgi(q, 2, MSFlexGrid(0)))
    Next

    If optRegressionModel(0).Value Then        'linear regression
        For q = 1 To m
            If Not f(q) <= 0 Then
                n(1) = n(1) + f(q)
                For w = 1 To f(q)
                    n(2) = n(2) + x(q)
                    n(3) = n(3) + y(q)
                    n(4) = n(4) + x(q) ^ 2
                    n(5) = n(5) + y(q) ^ 2
                    n(6) = n(6) + (x(q) * y(q))
                Next
            End If
         Next
    ElseIf optRegressionModel(1).Value Then    'logarithmic regression
        For q = 1 To m
            If Not f(q) <= 0 Then
                n(1) = n(1) + f(q)
                For w = 1 To f(q)
                    n(2) = n(2) + Ln(x(q))
                    n(3) = n(3) + y(q)
                    n(4) = n(4) + Ln(x(q)) ^ 2
                    n(5) = n(5) + y(q) ^ 2
                    n(6) = n(6) + Ln(x(q)) * y(q)
                Next
            End If
        Next
    ElseIf optRegressionModel(2).Value Then    'exponential regression
        For q = 1 To m
            If Not f(q) <= 0 Then
                n(1) = n(1) + f(q)
                For w = 1 To f(q)
                    n(2) = n(2) + x(q)
                    n(3) = n(3) + Ln(y(q))
                    n(4) = n(4) + x(q) ^ 2
                    n(5) = n(5) + Ln(y(q)) ^ 2
                    n(6) = n(6) + (x(q) * Ln(y(q)))
                Next
            End If
        Next
    ElseIf optRegressionModel(3).Value Then    'power regression
        For q = 1 To m
            If Not f(q) <= 0 Then
                n(1) = n(1) + f(q)
                For w = 1 To f(q)
                    n(2) = n(2) + Ln(x(q))
                    n(3) = n(3) + Ln(y(q))
                    n(4) = n(4) + Ln(x(q)) ^ 2
                    n(5) = n(5) + Ln(y(q)) ^ 2
                    n(6) = n(6) + Ln(x(q)) * Ln(y(q))
                Next
            End If
        Next
    End If
    On Error Resume Next
    n(7) = n(2) / n(1)
    n(8) = n(3) / n(1)
    h = n(1) * n(6) - n(2) * n(3)
    p = n(1) * n(4) - n(2) ^ 2
    o = n(1) * n(5) - n(3) ^ 2
    n(9) = Sqr(p / n(1) ^ 2)
    n(10) = Sqr(o / n(1) ^ 2)
    n(11) = Sqr(p / (n(1) * (n(1) - 1)))
    n(12) = Sqr(o / (n(1) * (n(1) - 1)))
    n(14) = h / p
    If optRegressionModel(0).Value Or optRegressionModel(1).Value Then
        n(13) = (n(3) - n(14) * n(2)) / n(1)
    Else
        n(13) = Exp((n(3) - n(14) * n(2)) / n(1))
    End If
    
    n(15) = h / Sqr(p * o)
    
    a = n(13)
    b = n(14)
    
    For q = 1 To 15
        MSFlexGrid(1).TextArray(Fgi(q, 1, MSFlexGrid(1))) = n(q)
    Next
End Sub

Private Sub cmdOpen_Click()
    Dim Msg As VbMsgBoxResult
    Dim x As Integer
    Dim j As Long
    Dim inputString(3) As String
    If Not DataSaved Then
        Msg = MsgBox("The existing data has not been saved. Save Data first?", vbYesNoCancel + vbQuestion)
        Select Case Msg
            Case vbCancel
                Exit Sub
            Case vbYes
                cmdSave_Click
                If Not DataSaved Then Exit Sub
        End Select
    End If
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "Comma Delimited Text (*.txt;*.csv)|*.txt;*.csv"
    On Error GoTo CancelLoad
    CommonDialog1.ShowOpen
    On Error GoTo ErrorLoadingFile
    MSFlexGrid(0).Rows = 1
    
    Open CommonDialog1.FileName For Input As #1
        Line Input #1, inputString(1)
        For x = 1 To Len(inputString(1))
            If Mid(inputString(1), x, 1) = "," Then j = j + 1
        Next
    Close #1
    If Not j = 2 Then Err.Raise (13)
    
    Open CommonDialog1.FileName For Input As #1
        Do
            
            For x = 1 To 3
                Input #1, inputString(x)
                If Not IsNumeric(inputString(x)) Then
                    Err.Raise (13)
                End If
            Next
            MSFlexGrid(0).AddItem inputString(1) & Chr(9) & inputString(2) & Chr(9) & inputString(3)
        Loop Until EOF(1)
    Close #1
    DataSaved = True
    MSFlexGrid(0).ZOrder
    ModifyDataMode (False)
    RegressionModeChanged = True
    For x = 0 To 2
        Text1(x) = ""
    Next
    txtSolve = ""
    lblSolveOut = ""
    For x = 1 To 15
        MSFlexGrid(1).TextArray(Fgi(x, 1, MSFlexGrid(1))) = ""
    Next
    
    Exit Sub
ErrorLoadingFile:
    Close #1
    MSFlexGrid(0).Rows = 1
    MsgBox "It appears that the file selected is not a valid Comma Delimited Text file.", vbExclamation
CancelLoad:
End Sub

Private Sub cmdSave_Click()
    Dim q As Integer
    If Not MSFlexGrid(0).Rows = 1 Then
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNOverwritePrompt
        CommonDialog1.Filter = "Comma Delimited Text (*.csv;*.txt)|*.csv;*.txt"
        On Error GoTo CancelSave
        CommonDialog1.ShowSave
        'On Error GoTo 0
        On Error GoTo ErrorSavingFile
        Open CommonDialog1.FileName For Output As #1
            For q = MSFlexGrid(0).FixedRows To MSFlexGrid(0).Rows - 1
                Write #1, Val(MSFlexGrid(0).TextArray(Fgi(q, 0, MSFlexGrid(0)))), _
                          Val(MSFlexGrid(0).TextArray(Fgi(q, 1, MSFlexGrid(0)))), _
                          Val(MSFlexGrid(0).TextArray(Fgi(q, 2, MSFlexGrid(0))))
            Next
        Close #1
        DataSaved = True
    Else
        MsgBox "Nothing to save. Either enter data or load data from a CSV file.", vbExclamation
    End If
    Exit Sub
ErrorSavingFile:
    Close #1
    MsgBox "Run-time error '" & Err.Number & "': " & vbCrLf & vbCrLf & _
            Err.Description, vbExclamation
CancelSave:
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Names
    
    Caption = App.Title
    
    Names = Array("n", _
                  "Sum of x", _
                  "Sum of y", _
                  "Sum of x^2", _
                  "Sum of y^2", _
                  "Sum of xy", _
                  "Sum of x/n", _
                  "Sum of y/n", _
                  "x o n", _
                  "y o n", _
                  "x o n-1", _
                  "y o n-1", _
                  "a", _
                  "b", _
                  "r")
                  
    DataSaved = True
    For i = 1 To 15
        MSFlexGrid(1).TextArray(Fgi(i, 0, MSFlexGrid(1))) = Names(i - 1)
    Next
                  
    For i = 1 To 3
        Load optRegressionModel(i)
        optRegressionModel(i).Top = optRegressionModel(i - 1).Top + 250
        optRegressionModel(i).Visible = True
    Next i
    MSFlexGrid(0).ToolTipText = "Double-click the rows in this grid to modify data."
    MSFlexGrid(0).TextArray(Fgi(0, 0, MSFlexGrid(0))) = "x data"
    MSFlexGrid(0).TextArray(Fgi(0, 1, MSFlexGrid(0))) = "y data"
    MSFlexGrid(0).TextArray(Fgi(0, 2, MSFlexGrid(0))) = "frequency"
    MSFlexGrid(0).ColWidth(0) = 840
    MSFlexGrid(0).ColWidth(1) = 840
    MSFlexGrid(0).ColWidth(2) = 825
    MSFlexGrid(1).ColWidth(0) = 1020
    MSFlexGrid(1).ColWidth(1) = 1755
    MSFlexGrid(0).SelectionMode = flexSelectionByRow
    
    
    optRegressionModel(0).Caption = "Linear"   ' Put caption on each option button.
    optRegressionModel(1).Caption = "Logarithmic"
    optRegressionModel(2).Caption = "Exponential"
    optRegressionModel(3).Caption = "Power"
    
    optRegressionModel(0).Value = True

End Sub

Private Sub txtSolve_GotFocus()
    txtSolve.SelStart = 0
    txtSolve.SelLength = Len(txtSolve)
End Sub

Private Function Fgi(r As Integer, c As Integer, FlexGrid As Control) As Integer
    Fgi = c + FlexGrid.Cols * r
End Function

Private Sub ModifyDataMode(DataMode As Boolean)
    If DataMode Then
        cmdAdd.Caption = "Modify Data"
        ModifyData = True
        MSFlexGrid(0).ToolTipText = "Double-click again to cancel modifying data."
    Else
        cmdAdd.Caption = "Add Data"
        ModifyData = False
        MSFlexGrid(0).ToolTipText = "Double-click the rows in this grid to modify data."
    End If
End Sub
