VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Gemmas Special Calculator"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command17 
      Caption         =   "About"
      Height          =   375
      Left            =   4800
      TabIndex        =   28
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "."
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "3"
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "2"
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "1"
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "0"
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "7"
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "6"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "5"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "4"
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "9"
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "8"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox showme 
      BackColor       =   &H80000018&
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton cancelbutton 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      ToolTipText     =   "Click this and your sum will be cancled"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      ToolTipText     =   "click this and number 2 will be taken away from number 1"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   "click this and your numbers will divide"
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      ToolTipText     =   "Click this and your numbers will multiply"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox answer 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "here is where your answer will appear"
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      ToolTipText     =   "Click this and it will add your numbers together"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox sum2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "enter a number in here"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox sum1 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "enter a number in here"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "2004"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Answer Box"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Number 2"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Number 1"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "By Gemma Waugh Age 10 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.Label sign 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim isgood, operator As Boolean
Dim a, b, c, d As Integer


Private Sub cancelbutton_Click()
sum1.Text = ""
sum2.Text = ""
answer.Text = ""
sign.Caption = ""
showme.Text = ""
showme.Visible = False
operator = False
End Sub

Private Sub Command1_Click()
operator = True
sign.Caption = "+"
good
If isgood = flase Then
Exit Sub
End If
If sum1.Text > 99999999 Then
MsgBox "That Number Is to Large Make it Smaller Please"
cancelbutton_Click
Exit Sub
End If


End Sub

Private Sub Command10_Click()
If operator = False Then
sum1.Text = sum1.Text & "5"
Else
sum2.Text = sum2.Text & "5"
End If
End Sub

Private Sub Command11_Click()
If operator = False Then
sum1.Text = sum1.Text & "6"
Else
sum2.Text = sum2.Text & "6"
End If
End Sub

Private Sub Command12_Click()
If operator = False Then
sum1.Text = sum1.Text & "7"
Else
sum2.Text = sum2.Text & "7"
End If
End Sub

Private Sub Command13_Click()
If operator = False Then
sum1.Text = sum1.Text & "0"
Else
sum2.Text = sum2.Text & "0"
End If
End Sub

Private Sub Command14_Click()
If operator = False Then
sum1.Text = sum1.Text & "1"
Else
sum2.Text = sum2.Text & "1"
End If
End Sub

Private Sub Command15_Click()
If operator = False Then
sum1.Text = sum1.Text & "2"
Else
sum2.Text = sum2.Text & "2"
End If
End Sub

Private Sub Command16_Click()
If operator = False Then
sum1.Text = sum1.Text & "3"
Else
sum2.Text = sum2.Text & "3"
End If
End Sub

Private Sub Command17_Click()
frmAbout.Show
End Sub

Private Sub Command2_Click()
operator = True
sign.Caption = "*"
good
If isgood = flase Then
Exit Sub
End If

End Sub

Private Sub Command3_Click()
operator = True
sign.Caption = "/"
good
If isgood = flase Then
Exit Sub
End If


End Sub

Private Sub Command4_Click()
operator = True
sign.Caption = "-"
good
If isgood = flase Then
Exit Sub
End If


End Sub

Public Sub good()
isgood = False
If sum1.Text = "" Then
Exit Sub
End If
If sum2.Text = "" Then
Exit Sub
End If
isgood = True
End Sub

Private Sub Command5_Click()
If operator = False Then
sum1.Text = sum1.Text & "8"
Else
sum2.Text = sum2.Text & "8"
End If
End Sub

Private Sub Command6_Click()
If operator = False Then
sum1.Text = sum1.Text & "9"
Else
sum2.Text = sum2.Text & "9"
End If
End Sub

Private Sub Command7_Click()

a = Val(sum1)
b = Val(sum2)
If sign.Caption = "+" Then
c = a + b
answer.Text = c
Exit Sub
End If
If sign.Caption = "*" Then
c = a * b
answer.Text = c
tables
Exit Sub
End If
If sign.Caption = "/" Then
c = a / b
answer.Text = c
Exit Sub
End If
If sign.Caption = "-" Then
c = a - b
answer.Text = c
Exit Sub
End If
If sign.Caption = "" Then
Exit Sub
End If
End Sub

Private Sub Command8_Click()
If operator = False Then
sum1.Text = sum1.Text & "."
Else
sum2.Text = sum2.Text & "."
End If
End Sub

Private Sub Command9_Click()
If operator = False Then
sum1.Text = sum1.Text & "4"
Else
sum2.Text = sum2.Text & "4"
End If
End Sub

Public Sub tables()

showme.Visible = True
showme.Text = ""
For d = 1 To 12
If d < 10 Then
result = "0" & d & " x " & a & " = " & d * a & " "
Else
result = d & " x " & a & " = " & d * a & " "
End If
showme = showme + result & vbCrLf
Next d
End Sub

Private Sub Label1_Click()
frmAbout.Show
End Sub

Private Sub sum1_Change()
On Error GoTo myerror
If sum1.Text = "" Then
Exit Sub
End If
If sum1.Text > 99999999 Then
MsgBox "That Number Is to Large Make it Smaller Please"
cancelbutton_Click
Exit Sub
End If
myerror:
End Sub

Private Sub sum2_Change()
On Error GoTo myerror
If sum2.Text = "" Then
Exit Sub
End If
If sum2.Text > 99999999 Then
MsgBox "That Number Is to Large Make it Smaller Please"
sum2.Text = ""
Exit Sub
End If
myerror:
End Sub
