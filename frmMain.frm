VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form Test"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF00FF&
      Height          =   1095
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I havent .hwnd property but i dont make any error!"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   1890
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   2520
      TabIndex        =   2
      Top             =   5280
      Width           =   4695
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FF00FF&
      Caption         =   "Me too!"
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   3600
      Width           =   6855
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H00FF00FF&
      Caption         =   "Im transparent!"
      Height          =   915
      Left            =   1560
      MaskColor       =   &H00CCCCCC&
      TabIndex        =   0
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   8085
      Left            =   600
      Picture         =   "frmMain.frx":0000
      Top             =   360
      Width           =   10005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim objT As clsTrans
    
    Set objT = New clsTrans
    
    Me.Show
    DoEvents 'remember to do that!
    
    
    objT.DoIT chk
    objT.DoIT opt
    objT.DoIT Frame1
    objT.DoIT Label1 'no error if you pass an object that doesnt allow .hwnd property
    

    Set objT = Nothing
End Sub
