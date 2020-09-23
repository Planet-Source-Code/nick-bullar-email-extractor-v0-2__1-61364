VERSION 5.00
Begin VB.Form frmerror 
   Caption         =   "Error!"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form2"
   ScaleHeight     =   1725
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4680
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "10"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Continue in..."
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Error: Uknown Error Occured!"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmerror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
epp = 0
frmerror.Visible = False
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Label3.Caption - 1
If Label3.Caption = 0 Then Timer1.Enabled = False: Command1_Click
End Sub
