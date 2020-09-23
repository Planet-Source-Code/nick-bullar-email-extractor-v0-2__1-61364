VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Email Extractor v0.2"
   ClientHeight    =   8325
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11310
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton fnish 
      Caption         =   "Stop"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton strt 
      Caption         =   "GO"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton rmfrom2 
      Caption         =   "Remove Selected"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton clrlist2 
      Caption         =   "ClearList"
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   7320
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   7920
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton loadlst 
      Caption         =   "Load List"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton addfoldertosearch 
      Caption         =   "Add Folder To Search"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton clrlist1 
      Caption         =   "ClearList"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton rmfrom1 
      Caption         =   "Remove Selected"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   7320
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   6105
      Left            =   6000
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton savlist 
      Caption         =   "Save List"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton addfiletosearch 
      Caption         =   "Add File To Search"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   240
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "http://planet-source-code.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7440
      TabIndex        =   17
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Like this program? Please vote for it at:"
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Email Adresses Found:"
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Current File: "
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   7680
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Files And Folders that will be scanned:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   3975
   End
   Begin VB.Menu File1 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu explist 
         Caption         =   "Export Search List"
         Index           =   2
      End
      Begin VB.Menu implst 
         Caption         =   "Import Search List"
         Index           =   3
      End
      Begin VB.Menu savemails 
         Caption         =   "Save Email List"
         Index           =   4
      End
      Begin VB.Menu Exit1 
         Caption         =   "Exit"
         Index           =   5
      End
   End
   Begin VB.Menu Tools1 
      Caption         =   "Tools"
      Index           =   7
      Begin VB.Menu Options1 
         Caption         =   "Options..."
         Index           =   7
      End
   End
   Begin VB.Menu About1 
      Caption         =   "About"
      Index           =   8
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim globestop

Private Sub About1_Click(Index As Integer)
MsgBox "Email Extractor v0.2 By Nick Bullar (c)2005."
End Sub

Private Sub addfiletosearch_Click()

InitOpen "*.*", "File To Search...", App.path
temp = Open_File(Me.hWnd)
List1.AddItem temp
End Sub

Private Sub Exit1_Click(Index As Integer)
End
End Sub

Private Sub explist_Click(Index As Integer)
savlist_Click
End Sub

Private Sub fnish_Click()
globestop = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub implst_Click(Index As Integer)
loadlst_Click
End Sub

Private Sub Label5_Click()
CreateObject("wscript.shell").Run "http://planet-source-code.com"
End Sub

Private Sub Options1_Click(Index As Integer)
FrmOptions.Visible = True
End Sub

Private Sub rmfrom2_Click()
aa = -1
While aa < List2.ListCount - 1
aa = aa + 1
If List2.Selected(aa) Then List2.RemoveItem aa: aa = aa - 1
Wend
End Sub

Private Sub savemails_Click(Index As Integer)
InitSave "*.txt", "Save Search List...", App.path
temp = Save_File(Me.hWnd, "txt")
If temp = "" Then Exit Sub
Set tlist = fso.createtextfile(temp)
For a = 0 To List2.ListCount - 1
tlist.WriteLine List2.List(a)
Next
End Sub

Private Sub savlist_Click()
InitSave "*.lst", "Save Search List...", App.path
temp = Save_File(Me.hWnd, "lst")
If temp = "" Then Exit Sub
Set tlist = fso.createtextfile(temp)
tlist.WriteLine "<Email Extractor List Files>"
For a = 0 To List1.ListCount - 1
tlist.WriteLine List1.List(a)
Next
tlist.WriteLine "<End List>"
End Sub

Private Sub rmfrom1_Click()
aa = -1
While aa < List1.ListCount - 1
aa = aa + 1
If List1.Selected(aa) Then List1.RemoveItem aa: aa = aa - 1
Wend
End Sub

Private Sub clrlist1_Click()
If MsgBox("Are you sure?", vbYesNo, "Clear list?") = vbYes Then List1.Clear
End Sub

Private Sub clrlist2_Click()
If MsgBox("Are you sure?", vbYesNo, "Clear list?") = vbYes Then List2.Clear
End Sub

Private Sub addfoldertosearch_Click()
temp = BrowseForFolder(Me.hWnd, "Select a folder to add")
If temp = "" Then Exit Sub
Folderadd temp
End Sub

Private Sub loadlst_Click()
On Error GoTo Errh:
If List1.List(0) <> "" Then
If MsgBox("Would you like to keep the current enteries?", vbYesNo, "Keep Enteries?") = vbNo Then List1.Clear
End If
InitOpen "*.lst", "Load Search List...", App.path
temp = Open_File(Me.hWnd)
If temp = "" Then Exit Sub
Set tlist = fso.opentextfile(temp)
If tlist.ReadLine <> "<Email Extractor List Files>" Then MsgBox "File not a valid export list or corrput!", vbCritical, "Error reading list!": Exit Sub
For a = 1 To 10000000
axx = tlist.ReadLine
If axx = "<End List>" Then GoTo Escc:
List1.AddItem axx ' Expect Error 62
Next
Escc:
Exit Sub
Errh:
If Err.Number <> 62 Then
frmerror.Visible = True
frmerror.Label3.Caption = 10
frmerror.Timer1.Enabled = True
frmerror.Label1.Caption = "Error: " & Err.Number & " occured." & vbNewLine & Err.Description
epp = 1
While epp = 1
DoEvents
Wend
End If
End Sub

Private Sub Form_Load()
Set fso = CreateObject("scripting.filesystemobject")
FrmOptions.extmake
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NumFiles
NumFiles = Data.Files.Count
For a = 1 To NumFiles
If fso.folderexists(Data.Files(a)) Then Folderadd Data.Files(a) Else Fileadd Data.Files(a)
Next
End Sub

Sub Fileadd(thepath)
If searchlist1(thepath) = "yes" Then Exit Sub
List1.AddItem thepath
End Sub

Sub Folderadd(thepath)
Dim extmatch
Set Folder = fso.getfolder(thepath)
For Each Fi In Folder.Files
extmatch = 0
For Each Ext In extarray()
If Right$(Fi, Len(Ext)) = Ext Then extmatch = 1: GoTo brkloop:
Next
brkloop:
If extmatch = 1 Then Fileadd Fi
DoEvents
Next
For Each su In Folder.SubFolders
Folderadd su
Next
End Sub

Function searchlist1(sfor)
For a = 0 To List1.ListCount
If List1.List(a) = sfor Then
searchlist = "yes": GoTo enn:
Else: searchlist = "no"
End If
Next
enn:
searchlist1 = searchlist
End Function


Private Sub strt_Click()
On Error GoTo Errh:
NumFiles = List1.ListCount
For a = 0 To List1.ListCount - 1
Label2.Caption = "Current File: " & List1.List(a)
SearchForEmail List1.List(a)
ProgressBar1.Value = (a / NumFiles) * 100
DoEvents
If globestop = 1 Then globestop = 0: Exit Sub
Next
If FrmOptions.Check2.Value = 1 Then
Label2.Caption = "Removing Duplicates..."
On Error Resume Next

  For X = 1 To 3 ' Filter 3 times

For a = 0 To List2.ListCount - 1
For b = a + 1 To List2.ListCount - 1
If List2.List(b) = List2.List(a) Then List2.RemoveItem (b)
Next
DoEvents
Next

  Next

End If
Label2.Caption = "Done!"
ProgressBar1.Value = 100
Exit Sub
Errh:
If Err.Number <> 62 Then
frmerror.Visible = True
frmerror.Label3.Caption = 10
frmerror.Timer1.Enabled = True
frmerror.Label1.Caption = "Error: " & Err.Number & " occured." & vbNewLine & Err.Description
epp = 1
While epp = 1
DoEvents
Wend
End If
End Sub
