VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail - DUMP"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Sort"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load list"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin MSComctlLib.ListView l1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483643
      BackColor       =   6166313
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NAME"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "E-MAIL"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PHONE"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NOTES"
         Object.Width           =   2610
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InptBx$
Dim Icnt%
Dim HL$, HL1$, HL2$, HL3$, HL4$
Dim V
Dim P%
Dim StrTmp$
Dim StrTmp1$
Dim StrTmp2$
Dim StrTmp3$
Dim StrTmp4$
Dim StrHld$
Private Sub Command1_Click()
If Dir(App.Path & "\" & "l.inj") = "" Then
Open App.Path & "\" & "l.inj" For Output As #1
Close #1
GoTo made
Else
made:
l1.ListItems.Clear
Open App.Path & "\" & "l.inj" For Input As #1
Do Until EOF(1)
Line Input #1, StrHld
Icnt = Icnt + 1
StrTmp = Extract(StrHld, "¢", "¢", "*")
StrTmp1 = Extract(StrHld, "È", "È", "~")
StrTmp2 = Extract(StrHld, "+", "+", "!")
StrTmp3 = Extract(StrHld, "¿", "¿", ":")
StrTmp4 = Extract(StrHld, "Ý", "Ý", ">")
l1.ListItems.Add , , Mid$(StrTmp, 2, Len(StrTmp))
l1.ListItems(Icnt).ListSubItems.Add , , Mid$(StrTmp1, 2, Len(StrTmp1))
l1.ListItems(Icnt).ListSubItems.Add , , Mid$(StrTmp2, 2, Len(StrTmp2))
l1.ListItems(Icnt).ListSubItems.Add , , Mid$(StrTmp3, 2, Len(StrTmp3))
Loop
Close #1
Icnt = 0
Label1 = "DATE BOOK LAST MODIFIED:  " & Mid$(StrTmp4, 2, Len(StrTmp4))
Label2 = l1.ListItems.Count & "   ENTRIES IN BOOK"
If l1.ListItems.Count = 0 Then
Exit Sub
Else
details
End If
End If
End Sub

Private Sub Command2_Click()
Form2.Caption = "ADD"
Form2.Show
End Sub

Private Sub Command3_Click()
If l1.ListItems.Count = 0 Then
MsgBox "No data to save", vbOKOnly, "Alert!"
Exit Sub
End If
Open App.Path & "\" & "l.inj" For Output As #1
For x = 1 To l1.ListItems.Count
P = P + 1
HL = "¢*" & l1.ListItems(P).Text & "¢"
HL1 = "È~" & l1.ListItems(P).ListSubItems(1).Text & "È"
HL2 = "+!" & l1.ListItems(P).ListSubItems(2).Text & "+"
HL3 = "¿:" & l1.ListItems(P).ListSubItems(3).Text & "¿"
HL4 = HL & HL1 & HL2 & HL3
Print #1, HL4 & "Ý>" & Date & "Ý"
Next
P = 0
Text1 = HL4
Close #1
End Sub




Private Sub Command4_Click()
Edit
End Sub

Private Sub Command5_Click()
l1.Sorted = True
End Sub



Private Sub Form_Load()
If App.PrevInstance = True Then End
Icnt = 0
Me.Hide
Form3.Show
App.HelpFile = App.Path & "\help.hlp"
End Sub


Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub l1_Click()
If Form2.Visible = True Then Form2.Hide
details
End Sub



Sub details()
If l1.ListItems.Count = 0 Then Exit Sub
Label3 = ""
Label3 = Label3 & "NAME:  " & l1.SelectedItem.Text & vbCrLf & vbCrLf & vbCrLf
Label3 = Label3 & "EMAIL:  " & l1.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & vbCrLf
Label3 = Label3 & "PHONE: " & l1.SelectedItem.SubItems(2) & vbCrLf & vbCrLf & vbCrLf
Label3 = Label3 & "NOTES: " & l1.SelectedItem.SubItems(3) & vbCrLf & vbCrLf & vbCrLf
End Sub

Sub Edit()
If l1.ListItems.Count = 0 Then Exit Sub
Form2.Text1 = l1.SelectedItem.Text
Form2.Text2 = l1.SelectedItem.SubItems(1)
Form2.Text3 = l1.SelectedItem.SubItems(2)
Form2.Text4 = l1.SelectedItem.SubItems(3)
Form2.Caption = "EDIT"
Form2.Show
End Sub
