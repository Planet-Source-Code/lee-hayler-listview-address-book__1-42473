VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   2970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   765
      Left            =   120
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
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
      TabIndex        =   7
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BOOL As Boolean
Dim R%
Private Sub Command1_Click()
If Text1 = "" Then Text1 = "Empty"
If Text2 = "" Then Text2 = "Empty"
If Text3 = "" Then Text3 = "Empty"
If Text4 = "" Then Text4 = "Empty"
Select Case Me.Caption
Case "ADD"
Command1.Caption = "Add"
BOOL = True
R = Form1.l1.ListItems.Count + 1
Form1.l1.ListItems.Add , , Text1
Form1.l1.ListItems(R).ListSubItems.Add , , Text2
Form1.l1.ListItems(R).ListSubItems.Add , , Text3
Form1.l1.ListItems(R).ListSubItems.Add , , Text4
Form1.Label2 = Form1.l1.ListItems.Count & "   ENTRIES IN BOOK"
Unload Me
Case Else
Form1.l1.SelectedItem.Text = Text1
Form1.l1.SelectedItem.SubItems(1) = Text2
Form1.l1.SelectedItem.SubItems(2) = Text3
Form1.l1.SelectedItem.SubItems(3) = Text4
Command1.Caption = "Save"
Unload Me
End Select
End Sub


