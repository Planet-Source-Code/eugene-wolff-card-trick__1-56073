VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Trick - Written by Eugene Wolff 2004"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRow 
      Caption         =   "Command3"
      Height          =   465
      Index           =   0
      Left            =   7140
      TabIndex        =   0
      Top             =   1005
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   255
      TabIndex        =   1
      Top             =   180
      Width           =   6690
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************
'********************************************************************************************
'
' This is a simple card trick to demonstrate how to use the cards.dll
' Written by Eugene Wolff
' Special thanks to Justin Yates for his help
'
'********************************************************************************************
'********************************************************************************************

Dim N(52) As Integer
Dim N1(52) As Integer

Dim times As Integer            ' The number of times the person has selected the row, only needs to be 3 times for this trick

' A routine to get unique numbers in an array
Private Sub ShuffleCards(intMaxNumber As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Randomize
    N(1) = Int(Rnd * intMaxNumber) + 1
    i = 2
    While i <= intMaxNumber
        j = 1
        N(i) = Int(Rnd * intMaxNumber) + 1
        While j < i
            If N(i) = N(j) Then
                N(i) = Int(Rnd * intMaxNumber) + 1
                j = 1
            Else
                j = j + 1
            End If
        Wend
        i = i + 1
    Wend
End Sub

'Find out which row was clicked
Private Sub cmdRow_Click(Index As Integer)
    Dim x As Integer
    Dim z As Integer
    Dim y As Integer
    Dim a As Integer
    Dim i As Integer
    
    times = times + 1
    z = Index - 1
    If z = 0 Then z = 3
    For x = 1 To 3
        For y = 0 To 6
            a = a + 1
            If z > 3 Then
                z = 1
            End If
            i = y * 3 + z
            N1(a) = N(i)
        Next y
        z = z + 1
    Next x
    
    For x = 1 To 21
        N(x) = N1(x)
    Next
    
    If times = 3 Then
        GoTo ThisIsYourCard             ' If the row has been selected 3 times
    End If
    
    DisplayCards
    
    frmMain.Refresh
    
    lblInstructions.Caption = "Select the row that your card is in."
    
    Exit Sub

ThisIsYourCard:

    Me.Cls
    
    cdtInit CardWidth, CardHeight
    cdtDraw Me.hdc, 200, 100, N(11) - 1, 0, vbWhite     ' Display the secret card
    
    For x = 1 To 3
        cmdRow(x).Visible = False
    Next x
    
    frmMain.Refresh
    
    lblInstructions.Caption = "This is your card."
End Sub

' Load the form
Private Sub Form_Load()
    Dim x As Integer
    Dim z As Integer
    
    frmMain.AutoRedraw = True       ' This is very important or you won't see the cards
    
    ShuffleCards (52)               ' Shuffle the cards
    
    DisplayCards                    ' Display the cards
    
    x = 0
    x = x + 1
    
    For z = 1 To 3
        Load cmdRow(z)
        With cmdRow(z)
            .Visible = True
            .Left = 1500 * x
            .Top = 7000
            .Caption = "Row " & z
            x = x + 1
        End With
    Next z
    
    frmMain.Refresh                 ' This is very important or you will not see the changes to your cards
    
    lblInstructions.Caption = "Select a card and tell me what row it is in."
End Sub

' Display the cards on the screen
Sub DisplayCards()
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
    Dim face As Integer
    
    cdtInit CardWidth, CardHeight   ' Initialise the cards
    x = 0
    y = 1
    For z = 1 To 21
        face = N(z) - 1             ' Get the shuffled card from the array
        x = x + 1
        cdtDraw Me.hdc, 100 * x, 50 * y, face, 0, vbRed ' Draw the cards
        If x = 3 Then               ' If 3 cards are across the screen move down and start from the first row again
            x = 0
            y = y + 1
        End If
    Next
    
End Sub

