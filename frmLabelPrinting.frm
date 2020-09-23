VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLabelPrinting 
   Caption         =   "Label printing"
   ClientHeight    =   4935
   ClientLeft      =   2025
   ClientTop       =   1995
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   10755
   Begin VB.CommandButton Exit 
      Caption         =   "Done"
      Height          =   495
      Left            =   9480
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox LittleFont 
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Text            =   "10"
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox BigFont 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "10"
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox NumberOfLines 
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Text            =   "2"
      Top             =   4320
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "frmLabelPrinting.frx":0000
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\practice\db1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MailingLabelDetails"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdTest2 
      Caption         =   "Print"
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      ToolTipText     =   "Uses class to print Avery 5352 label"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "All other lines font size"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Top line font size"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label NumberLabel 
      Caption         =   "Number of lines per label"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
End
Attribute VB_Name = "frmLabelPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLabel As clsAveryLabels

Private Sub cmdTest2_Click()
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim rc As Integer
    Dim printString As String
    
    Set objLabel = New clsAveryLabels
    objLabel.NumberOfColumns = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 14)
    objLabel.NumberOfRows = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 15)
    objLabel.LabelHeight = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
    objLabel.LabelWidth = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8)
    objLabel.TopMargin = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)
    objLabel.LeftMargin = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
    objLabel.HPitch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 13)
    objLabel.VPitch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)
    objLabel.BigFont = CInt(BigFont.Text)
    objLabel.LittleFont = CInt(LittleFont.Text)
    
    If objLabel.AvailablePointsHigh < objLabel.BigFont + objLabel.LittleFont * CInt(NumberOfLines.Text) Then
       rc = MsgBox("There is only enough space on the label for " + CStr(objLabel.AvailableLinesPerLabel) + " lines using these fonts.  Continue?", vbYesNo + vbExclamation)
       If rc = vbNo Then Exit Sub
       End If
    For i = 1 To objLabel.NumberOfRows
       For j = 1 To objLabel.NumberOfColumns
          For k = 1 To objLabel.AvailableLinesPerLabel
           '  printString = "line " + CStr(k) + " of " + CStr(i) + "'" + CStr(j)
             printString = ""
             If Len(printString) < objLabel.AvailableCharactersPerLabelLine Then
                For l = Len(printString) To objLabel.AvailableCharactersPerLabelLine
                   printString = printString + "X"
                   Next l
                End If
             objLabel.LabelPrint i, j, k, printString
             Next k
          Next j
       Next i
    objLabel.PageFinished
    Set objLabel = Nothing
End Sub

Private Sub Exit_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
Me.Move 0, 0
End Sub
