VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置窗口置前\正常\置后"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4095
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option3 
      Caption         =   "置后"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "正常"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "置前"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "操作句柄:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1290
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Option1_Click()
On Error Resume Next
  SetWindowPos CLng(Trim(Text1)), -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1: Timer1.Enabled = True
End Sub

Private Sub Option2_Click()
On Error Resume Next
  SetWindowPos CLng(Trim(Text1)), -2, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1: Timer1.Enabled = False
End Sub

Private Sub Option3_Click()
On Error Resume Next
  SetWindowPos CLng(Trim(Text1)), 1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1: Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  If Option1.Value = True Then
    SetWindowPos CLng(Trim(Text1)), -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  ElseIf Option3.Value = True Then
    SetWindowPos CLng(Trim(Text1)), 1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
  End If
End Sub
