VERSION 5.00
Begin VB.Form frmPP 
   Caption         =   "Really Simple Print Preview"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtText 
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "frmPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    txtText.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()

    Printer.Print txtText.Text
    Printer.EndDoc
End Sub


