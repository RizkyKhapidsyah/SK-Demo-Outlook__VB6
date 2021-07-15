VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mes Mails"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Send a mail"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   7560
      Width           =   7335
   End
   Begin VB.TextBox txtNewFolder 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create the folder"
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
      Left            =   0
      TabIndex        =   2
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      Height          =   6255
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00808080&
      Height          =   7470
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblSubject 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   600
      Width           =   6375
   End
   Begin VB.Label lblTo 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblFrom 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim item As MailItem
Set mf = ns.Folders(3)

On Error Resume Next
mf.Folders.Add txtNewFolder
End Sub

Private Sub Command2_Click()
Form2.Show vbModal

End Sub

Private Sub Form_Load()

 Set ns = ol.GetNamespace("MAPI")
 
 Set fldr = ns.GetDefaultFolder(olFolderInbox)
 List1.Clear
 
For j = 1 To fldr.Items.Count
    Set mm = fldr.Items(j)
    List1.AddItem mm.SenderName
Next

End Sub

Private Sub List1_Click()
Text1 = fldr.Items(List1.ListIndex + 1).Body
lblFrom = fldr.Items(List1.ListIndex + 1).SenderName
lblTo = fldr.Items(List1.ListIndex + 1).To
lblSubject = fldr.Items(List1.ListIndex + 1).Subject
End Sub
