VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeckOfCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ACard As Card
Private DeckOfCards(0 To 51) As Card

'creates the deck of cards we will be using
Private Sub Class_Initialize()
  Dim I As Integer
  For I = 0 To 51
    Set ACard = New Card
    ACard.Create I 'creates the card with the specific value
    Set DeckOfCards(I) = ACard
  Next I
End Sub

Public Sub Shuffle()
  Dim X As Integer, I As Integer
  Dim Temp As Card, Place As Integer
  Randomize
  For I = 0 To 5199
    Place = I Mod 52
    X = Int(52 * Rnd)
    Set Temp = DeckOfCards(Place)
    Set DeckOfCards(Place) = DeckOfCards(X)
    Set DeckOfCards(X) = Temp
  Next I
End Sub


Public Property Get CardValue(Position As Integer) As String
   CardValue = DeckOfCards(Position).CardValue
End Property



