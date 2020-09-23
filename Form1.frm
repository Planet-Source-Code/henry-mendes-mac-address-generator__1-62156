VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MACgen 1.0"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Copy"
      Height          =   465
      Left            =   1350
      TabIndex        =   2
      Top             =   1350
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   465
      Left            =   2625
      TabIndex        =   3
      Top             =   1350
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   0
      Text            =   "00-00-00-00-00-00"
      Top             =   750
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate"
      Height          =   465
      Left            =   150
      TabIndex        =   1
      Top             =   1350
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------
'//////////////////////////////////////
'
'Date: 08/11/2005
'
'MAC address Generator v1.0
'
'Programmed by Henry Mendes
'
'E-mail: hmjbetah@yahoo.com.br
' or
' hmjbetah@gmail.com
'
'You got permissions to modify, copy, redistribiute and etc.
'
'//////////////////////////////////////
'The author has no responsibility for any
'damage caused by MACgen
'--------------------------------------





Private Sub Command1_Click()

Text1 = gen_MAC(Text1)

End Sub
  
 Private Function gen_MAC(oobj As Object) As String
  
 Dim arrhex(15) As String
 Dim sTMP As String
 Dim sMAC As String
 
 
 Dim j
 
 For j = 0 To 9
  
    arrhex(j) = j
 
 Next j
 
    
 For j = 65 To 70
 
 'How it works: 10 + (5 - (70 - j))???
       ' examples
       '70 - j
       '70 - 65 = 5
       '5 - 5   = 0
       'Result 10+0 = 10
       '-------------
       '70 - 66 = 4
       '5 - 4   = 1
       'Result 10 + 1 = 11
       
       arrhex(10 + (5 - (70 - j))) = Chr(j)

 Next j
    
    
    
For i = 1 To 6
sTMP = vbNullString
Randomize

sTMP = arrhex(Int(Rnd * 15)) & arrhex(Int(Rnd * 15))

If i = 1 Then
sMAC = sMAC & sTMP
Else
sMAC = sMAC & "-" & sTMP
End If

Next i
' generate Mac



 
gen_MAC = sMAC
 
 End Function

Private Sub Command2_Click()
MsgBox "MACgen v1.0 Programmed by Henry Mendes.", , "About"
End Sub

Private Sub Command3_Click()
Clipboard.Clear

Clipboard.SetText (Text1.Text)
End Sub

Private Sub Form_Load()
Text1.Locked = True

Command1_Click


End Sub
