VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form txtForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A Little NEAT text to BMP Encryptor/ Decryptor"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save As..."
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   7095
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5640
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text Area"
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Text            =   "Form1.frx":0000
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Text            =   "Form1.frx":0010
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Encryption"
      Height          =   1695
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   1935
         TabIndex        =   12
         Top             =   240
         Width           =   1935
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1350
            Left            =   0
            Picture         =   "Form1.frx":0025
            ScaleHeight     =   1350
            ScaleWidth      =   45
            TabIndex        =   13
            Top             =   0
            Width           =   45
         End
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   7095
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":049F
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "txtForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare All Variables
Dim txtIndex As Integer
Dim R As Long, G As Long, B As Long, K, L

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Sub Form_Initialize()
Dim X As Long
X = InitCommonControls
End Sub

'Encrypt-----------------------------------------------------------------------------
Private Sub Command1_Click()
        Frame4.Caption = "Working..."
        Frame4.Refresh
        Picture1.Cls
        Picture1.Picture = Nothing
        Picture1.Width = ((Len(Text1) / 278) * 15) + 15
        ProgressBar1.Max = Len(Text1)
        Text3.ZOrder 0
Picture1.Height = 90 * 15
K = 0
L = 0
txtIndex = 0
For X = 0 To Len(Text1) / 3
L = L + 15
For Y = 1 To 3
Text1.SelStart = txtIndex
Text1.SelLength = 1

Select Case Y
Case 1
R = Ascii(Text1.SelText)
Case 2
G = Ascii(Text1.SelText)
Case 3
B = Ascii(Text1.SelText)
End Select
txtIndex = txtIndex + 1
On Error Resume Next
ProgressBar1.Value = txtIndex
Next
DoEvents
If L >= Picture1.Height Then
L = 0
K = K + 15
End If
Picture1.Line (K, L)-(K - 15, L - 15), RGB(R, G, B)
Next
        Frame4.Caption = ""
        Frame4.Refresh
        Text3.ZOrder 1
        Text1.SelStart = 0
End Sub
'--------------------------------------------------------------------------------------------
Private Function Ascii(cCode As String) As Integer
    'Loop Progress
    For i = 0 To 255
        'Check if Mached and send to End
        If Chr(i) = cCode Then GoTo Fin
    
    Next
    Exit Function
Fin:
'Set Code
Ascii = i

End Function
'Decrypt:------------------------------------------------------------------------------------
Private Sub Command2_Click()
Text3.ZOrder 0
Frame4.Caption = "Working..."
Frame4.Refresh
    ProgressBar1.Max = Picture1.Width
On Error Resume Next
Text1 = ""
For X = 0 To Picture1.Width Step 15
For Y = 0 To Picture1.Height Step 15
Text1.SelText = Chr(Red(Picture1.Point(X, Y)))
Text1.SelText = Chr(Green(Picture1.Point(X, Y)))
Text1.SelText = Chr(Blue(Picture1.Point(X, Y)))
DoEvents
Next
ProgressBar1.Value = X
Next
ProgressBar1.Value = ProgressBar1.Max
Frame4.Caption = ""
Text3.ZOrder 1
Text1.SelStart = 0
End Sub
Private Function Red(ByVal Color As Long) As Integer
Red = Color Mod &H100
End Function
Private Function Green(ByVal Color As Long) As Integer
Green = (Color \ &H100) Mod &H100
End Function
Private Function Blue(ByVal Color As Long) As Integer
Blue = (Color \ &H10000) Mod &H100
End Function
'Save TBMP Format text
Private Sub Command3_Click()
    
    
    
    On Error GoTo 10
    
    CD.Filter = "Text Bmp(*.tbp)|*.tbp"
    CD.ShowSave
    CD.CancelError = True
    Command1_Click
    SavePicture Picture1.Image, CD.FileName
10

End Sub
'Open TBMP Text
Private Sub Command4_Click()

    On Error GoTo 10
    
    CD.Filter = "Text Bmp(*.tbp)|*.tbp"
    CD.ShowOpen
    CD.CancelError = True

    Picture1.Picture = LoadPicture(CD.FileName)
    Picture1.AutoSize = True
    Picture1.Refresh

10
    If Error = "Cancel was selected." Then Exit Sub
    Command2_Click

End Sub

'---------------------------------------------------------------------------------------------
Private Sub Form_Load()
    'Decrypt ReadMe
    Call Command2_Click
    cCrypto = 0
End Sub

