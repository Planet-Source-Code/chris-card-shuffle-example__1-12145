VERSION 5.00
Begin VB.Form frmDOC 
   BackColor       =   &H00C06934&
   Caption         =   "Deck Of Cards"
   ClientHeight    =   1110
   ClientLeft      =   1590
   ClientTop       =   3345
   ClientWidth     =   6300
   FillColor       =   &H00C06934&
   ForeColor       =   &H00C06934&
   Icon            =   "frmDOC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   6300
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C06934&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   70
      MaskColor       =   &H00C06934&
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   730
      Width           =   6135
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackColor       =   &H00D3C098&
      Caption         =   $"frmDOC.frx":2DBA
      ForeColor       =   &H00C06934&
      Height          =   630
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   6135
   End
End
Attribute VB_Name = "frmDOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  MsgBox "Thank You For Downloading My Example... Please Vote For This Project!", vbOKOnly
    End
End Sub

Private Sub Form_Load()
Dim MyDeck As New DeckOfCards, I As Integer
Dim YourHand As String, MyHand As String
  MyDeck.Shuffle
  For I = 1 To 5
    YourHand = YourHand & MyDeck.CardValue(I) & "  "
  Next I
  MsgBox "Your hand is: " & YourHand
  For I = 6 To 10
    MyHand = MyHand & MyDeck.CardValue(I) & "  "
  Next I
  MsgBox "My hand is: " & MyHand
  
End Sub
