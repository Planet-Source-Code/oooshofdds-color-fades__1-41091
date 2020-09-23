VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stacked_shit@hotmail.com"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Clear after each effect"
      Height          =   255
      Left            =   3600
      TabIndex        =   26
      Top             =   2520
      Width           =   1935
   End
   Begin VB.OptionButton Option18 
      Caption         =   "Effect 18"
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   2160
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4560
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save to BMP"
      Height          =   495
      Left            =   2160
      TabIndex        =   21
      Top             =   2640
      Width           =   1335
   End
   Begin VB.OptionButton Option17 
      Caption         =   "Effect 17"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   1920
      Width           =   1095
   End
   Begin VB.OptionButton Option16 
      Caption         =   "Effect 16"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   1680
      Width           =   1095
   End
   Begin VB.OptionButton Option15 
      Caption         =   "Effect 15"
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton Option14 
      Caption         =   "Effect 14"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton Option13 
      Caption         =   "Effect 13"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton Option12 
      Caption         =   "Effect 12"
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option11 
      Caption         =   "Effect 11"
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton Option10 
      Caption         =   "Effect 10"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Effect 9"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Effect 8"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Effect 7"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Effect 6"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Effect 5"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Effect 4"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Effect 3"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Effect 2"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Effect 1"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   4200
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   3840
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Effect"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Base color fades"
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   5535
      Begin VB.CommandButton Command3 
         Caption         =   "Start"
         Height          =   975
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   240
         ScaleHeight     =   195
         ScaleWidth      =   3315
         TabIndex        =   23
         Top             =   960
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'stacked_shit@hotmail.com
'Stacked_shit@yahoo.com

Private Sub Command1_Click()
On Error Resume Next
If Check1.Value = 1 Then
Picture1.Cls
End If
For x = 1 To 10000 Step 0.3
DoEvents

If Option1.Value = True Then
Picture1.Circle (0, Picture1.Height / 2), x, RGB(x / 12, 0, 0)
End If

If Option1.Value = True Then
Picture1.Circle (Picture1.Width / Cos(x), 0), x, RGB(x / 12, 0, 0)
End If

If Option2.Value = True Then
Picture1.Line (x, 0)-(x, Picture1.Height), RGB(x / 12, 0, 0)
End If

If Option3.Value = True Then
Picture1.Circle (x, x / 2), x, RGB(x / 10, x / 12, x / 5)
End If

If Option4.Value = True Then
Picture1.Circle (x, x / Sin(x)), x, RGB(x / 10, x / 12, x / 5)
End If

If Option5.Value = True Then
Picture1.Circle (x, x * Sin(x)), x, RGB(x / 12, 0, 0)
End If

If Option6.Value = True Then
Picture1.Circle (x / Cos(x), x / Sin(x)), x, RGB(0, 0, x / 12)
End If

If Option7.Value = True Then
Picture1.Circle (x / Tan(x), x / Sin(x)), x, RGB(x / 12, 0, 0)
End If

If Option8.Value = True Then
Picture1.Circle (x / Tan(x / 2), x / Sin(x)), x, RGB(x / 12, 0, x / 10)
End If

If Option9.Value = True Then
Picture1.Circle (x / (Cos(x) / Sin(x)), x / Sin(x)), x, RGB(x / 12, 0, 0)
End If

If Option10.Value = True Then
Picture1.Circle (x / Tan(x / 2), x / Tan(x)), x, RGB(x / 15, 0, 0)
End If

If Option11.Value = True Then
Picture1.Circle (x / Cos(x / 2), x / 2), x, RGB(x / 12, 0, 0)
End If

If Option12.Value = True Then
Picture1.Circle (x, x / Tan(x * x)), x, RGB(x / 12, 0, 0)
End If

If Option13.Value = True Then
Picture1.Circle (Tan(x) / Sin(x / 2), Sin(x) / Tan(x) + x / 2), x, RGB(x / 12, 0, 0)
End If

If Option14.Value = True Then
Picture1.Circle (x, Tan(x + x / 2)), x, RGB(x / 10, x / 12, x / 5)
End If

If Option15.Value = True Then
Picture1.Circle (Tan(Cos(x / 1.8)), Tan(x + x) / 2), x, RGB(x / 10, x / 12, x / 5)
End If

If Option16.Value = True Then
Picture1.Circle (Tan(x), x * 0.5), x, RGB(x / 10, x / 12, Tan(x / 5))
End If

If Option17.Value = True Then
Picture1.Circle (Int(Tan(x)), ((x * 2) / Tan(x - 100))), x, RGB(x / 10, x / 11, Int(x / 3)): Picture1.DrawWidth = Tan(x / 0.1)
End If

If Option18.Value = True Then
Picture1.Circle (x / Tan(x), x), x, RGB(x / 12, 0, 0)
End If
Next x
End Sub

Private Sub Command2_Click()
On Error Resume Next
CD1.Filter = "Bitmap File (*.bmp)|*.bmp"
CD1.DialogTitle = "Save picture as ..."
CD1.ShowSave
SavePicture Picture1.Image, CD1.FileName
End Sub

Private Sub Command3_Click()
For x = 1 To Picture2.Width Step 0.3
DoEvents
Picture2.Line (x, 0)-(x, Picture2.Height), RGB(0, x / 12, 0)
Next x
For x = 1 To Picture3.Width Step 0.3
DoEvents
Picture3.Line (x, 0)-(x, Picture3.Height), RGB(0, 0, x / 12)
Next x
For x = 1 To Picture4.Width Step 0.3
DoEvents
Picture4.Line (x, 0)-(x, Picture4.Height), RGB(x / 12, 0, 0)
Next x
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Me
Set Form1 = Nothing
End
End Sub
