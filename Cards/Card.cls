VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CardIndex As Integer
'Now we will create a card with a specific value

Public Sub Create(Index As Integer)
  If Index > 51 Or Index < 0 Then
    Err.Raise Number:=vbObjectError + 32144, _
    Description:="Only 52 cards in deck!"
    Exit Sub
  End If
  CardIndex = Index
End Sub
   
Private Function CalculateCard(X As Integer) As String
  Dim Suit As Integer, CardValue As Integer, TheCard As String
  Suit = X \ 13
  Select Case Suit
   Case 0
     TheCard = "Clubs"
   Case 1
     TheCard = "Diamonds"
   Case 2
     TheCard = "Hearts"
   Case 3
     TheCard = "Spades"
   End Select
   
  CardValue = X Mod 13
  Select Case CardValue
   Case 0
    TheCard = "Ace of " + TheCard
   Case 1 To 9
    TheCard = Str$(CardValue + 1) + " of " + TheCard
   Case 10
     TheCard = "Jack of " + TheCard
   Case 11
     TheCard = "Queen of " + TheCard
   Case 12
     TheCard = "King of " + TheCard
   End Select
CalculateCard = TheCard
End Function

'ReadOnly Of Course
Public Property Get CardValue() As String
  CardValue = CalculateCard(CardIndex)
End Property


