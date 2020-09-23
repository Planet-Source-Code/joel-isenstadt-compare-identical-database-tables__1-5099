VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Compare Two Tables In Access"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTypeOutput 
      Caption         =   "Type of Output"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   4815
      Begin VB.OptionButton optTypeOutput 
         Caption         =   "Printer"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optTypeOutput 
         Caption         =   "Print Preview"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optTypeOutput 
         Caption         =   "File"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   4815
      Begin VB.TextBox txtOutputFile 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton cmdDB 
         Caption         =   "..."
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   5100
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6345
      Width           =   5160
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGetTables1 
      Caption         =   "Get Tables -->"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1245
      Width           =   1695
   End
   Begin VB.ComboBox cboTable1 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton cmdGetTables2 
      Caption         =   "Get Tables -->"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2925
      Width           =   1695
   End
   Begin VB.ComboBox cboTable2 
      Height          =   315
      Left            =   1920
      TabIndex        =   11
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Compare and Report"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   4815
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "..."
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtDB2 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "..."
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtDB1 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblTable1 
      Caption         =   "Second Database Table"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblTable2 
      Caption         =   "Second Database Table"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblDB2 
      Caption         =   "Second Database Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblDB1 
      Caption         =   "First Database Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:   Demonstrate the clsCompareDB Class
'Author:    Joel Isenstadt  12/25/99
'SEE THE NOTES IN THE CLASS FILE
'*********************************************************************************
Option Explicit

Private WithEvents cdb As clsCompareDB
Attribute cdb.VB_VarHelpID = -1
Private optChoice As Integer 'the index of the File type option buttons
Private Sub WriteStatus(sMsg As String)
'Purpose:   Displays a Line of text in a Picture Box (being used as a status bar)
    picStatus.Cls
    picStatus.Print sMsg
End Sub

Private Sub cdb_Error(sMsg As String)
    WriteStatus sMsg
End Sub

Private Sub cdb_OutputLine(sMsg As String)
    On Error GoTo errOutputline
    
    frmPP.txtText.Text = frmPP.txtText.Text & sMsg & vbCrLf
    Exit Sub
    
errOutputline:
    MsgBox "Error " & Err.Description
End Sub


Private Sub cdb_StatusMessage(sMsg As String)
    WriteStatus sMsg
End Sub

Private Sub cmdDB_Click(Index As Integer)
'Purpose:   To use the commondialog Open dialog
'           to select a filename
    With CommonDialog1
        .Filter = "Access databases|*.mdb|All Files|*.*"
        .FilterIndex = 1
        Select Case Index
            Case 0
                .FileName = txtDB1.Text
            Case 1
                .FileName = txtDB2.Text
            Case 2
                .FileName = txtOutputFile.Text
        End Select
        .ShowOpen
        If .FileName = "" Then Exit Sub
        Select Case Index
            Case 0
                txtDB1.Text = .FileName
            Case 1
                txtDB2.Text = .FileName
            Case 2
                txtOutputFile.Text = .FileName
        End Select
    End With
End Sub

Private Sub cmdGetTables1_Click()
'Purpose: Retrieve Table names and place into the drop down
    Dim i As Integer

    cdb.DB1Path = txtDB1.Text
    If cdb.GetDBTables(1, False) Then
        For i = 1 To cdb.DB1TableCount
            cboTable1.AddItem cdb.GetDB1TableName(i)
        Next i
    End If
End Sub

Private Sub cmdGetTables2_Click()
'Purpose: Retrieve Table names and place into the drop down
    Dim i As Integer

    cdb.DB2Path = txtDB2.Text
    If cdb.GetDBTables(2, False) Then
        For i = 1 To cdb.DB2TableCount
            cboTable2.AddItem cdb.GetDB2TableName(i)
        Next i
    End If

End Sub


Private Sub cmdReport_Click()
'Purpose:   Compare and Reports the table differences
    cdb.DB1Path = txtDB1.Text
    cdb.DB2Path = txtDB2.Text
    cdb.DB1TableName = cboTable1.Text
    cdb.DB2TableName = cboTable2.Text
    Select Case optChoice
        Case 0 'file
            cdb.OutputType = jiFile
            cdb.OutputFileName = txtOutputFile.Text
        Case 1 'print preview
            cdb.OutputType = jiPrintPreview
            frmPP.Show
        Case 2 'Printer
            cdb.OutputType = jiPrinter
    End Select
    cdb.CompareTables
End Sub

Private Sub Form_Load()
    'Set up the object variable
    Set cdb = New clsCompareDB
End Sub
Private Sub optTypeOutput_Click(Index As Integer)
    optChoice = Index
    Select Case Index
        Case 0
            fraOutput.Visible = True
        Case Else
            fraOutput.Visible = False
    End Select
End Sub


