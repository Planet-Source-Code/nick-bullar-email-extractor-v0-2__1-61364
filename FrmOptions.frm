VERSION 5.00
Begin VB.Form FrmOptions 
   Caption         =   "Options"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5520
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check4 
      Caption         =   "Use heavy results filter."
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   1440
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6360
      TabIndex        =   12
      Text            =   "2000"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Exclude files over:"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Automatically remove duplicates."
      Height          =   195
      Left            =   4680
      TabIndex        =   9
      Top             =   720
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton rmfrom1 
      Caption         =   "Remove Selected"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton clrlist1 
      Caption         =   "ClearList"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   4800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   ".txt"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "FrmOptions.frx":08CA
      Left            =   960
      List            =   "FrmOptions.frx":08E3
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "KB"
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Other Options:"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Automatically filter types not mentioned above."
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Files Types To Include:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clrlist1_Click()
If MsgBox("Are you sure?", vbYesNo, "Clear list?") = vbYes Then List1.Clear
End Sub

Private Sub Command1_Click()
List1.AddItem Text1.Text
End Sub

Private Sub Command2_Click()
extmake
FrmOptions.Visible = False
End Sub

Public Sub extmake()
ReDim extarray(List1.ListCount - 1) As String
For a = 0 To List1.ListCount - 1
extarray(a) = List1.List(a)
Next
End Sub


Private Sub rmfrom1_Click()
aa = -1
While aa < List1.ListCount - 1
aa = aa + 1
If List1.Selected(aa) Then List1.RemoveItem aa: aa = aa - 1
Wend
End Sub
