VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8565
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAbout 
      BackColor       =   &H00808080&
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   7695
   End
   Begin VB.TextBox txtTO 
      BackColor       =   &H00808080&
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   1320
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   8535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   4800
      TabIndex        =   3
      Text            =   "http://gmanouvrier.free.fr/img/gillesmanouvrier.gif"
      Top             =   960
      Width           =   3735
   End
   Begin VB.OptionButton opttype 
      BackColor       =   &H00000000&
      Caption         =   "HTML"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton opttype 
      BackColor       =   &H00000000&
      Caption         =   "Text"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5106
      _Version        =   393217
      BackColor       =   8421504
      Enabled         =   -1  'True
      TextRTF         =   $"Form2.frx":0000
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "TO"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "URL of the picture you want to send"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim img As String


Private Sub Command1_Click()
Dim item As Outlook.MailItem
If txtTO = "" And InStr(txtTO, "@") = 0 And InStr(txtTO, ".") = 0 Then GoTo HDL
If Text1 <> "" Then img = "<img src=" & Text1 & " >"

Set item = ol.CreateItem(olMailItem)
With item
    .Subject = txtAbout
    If opttype(0).Value Then .HTMLBody = "<html><head><title></title></head><body>" & img & rt.Text & "</body></html>" Else .Body = rt.Text
    
    .Recipients.Add (txtTO)
    .Send
End With
Exit Sub
HDL:
MsgBox "Aucune adresse de destinataire" & vbCrLf & "Error, No recipients"
txtTO.SetFocus


End Sub

Private Sub opttype_Click(Index As Integer)
If opttype(Index).Value And Index = 0 Then
    Text1.Enabled = False
Else
    Text1.Enabled = True
End If

End Sub

