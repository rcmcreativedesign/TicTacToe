VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   2550
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   8
      Left            =   1680
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   7
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   6
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   5
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   4
      Left            =   960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton btnSquare 
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblCurrentPlayer 
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblPlayerMessage 
      Caption         =   "Current Player: "
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private currentPlayer As Integer


Private Sub Form_Load()
    currentPlayer = 0
    
End Sub

Private Sub cmdNewGame_Click()
    Call InitializeGrid
    currentPlayer = 0
    lblPlayerMessage.Caption = "Current Player: "
    Call UpdateCurrentPlayerLabel
End Sub

Private Sub btnSquare_Click(Index As Integer)
    btnSquare(Index).Caption = lblCurrentPlayer.Caption
    btnSquare(Index).Enabled = False
    
    If Not IsWinner() Then
        Call NextPlayer
    Else
        Call Winner
    End If
End Sub

Private Sub InitializeGrid()
    For intLoop = 0 To 8
        btnSquare(intLoop).Caption = ""
        btnSquare(intLoop).Enabled = True
    Next
End Sub

Private Sub UpdateCurrentPlayerLabel()
    If currentPlayer = 0 Then
        lblCurrentPlayer.Caption = "X"
    Else
        lblCurrentPlayer.Caption = "O"
    End If
End Sub

Private Sub NextPlayer()
    If currentPlayer = 0 Then
        currentPlayer = 1
    Else
        currentPlayer = 0
    End If
    Call UpdateCurrentPlayerLabel
End Sub

Private Function IsWinner() As Boolean
    Dim player As String: player = lblCurrentPlayer.Caption
    Dim grid(8) As Boolean
    
    If (btnSquare(0).Caption = player And btnSquare(1).Caption = player And btnSquare(2).Caption = player) Or _
        (btnSquare(3).Caption = player And btnSquare(4).Caption = player And btnSquare(5).Caption = player) Or _
        (btnSquare(6).Caption = player And btnSquare(7).Caption = player And btnSquare(8).Caption = player) Or _
        (btnSquare(0).Caption = player And btnSquare(3).Caption = player And btnSquare(6).Caption = player) Or _
        (btnSquare(1).Caption = player And btnSquare(4).Caption = player And btnSquare(7).Caption = player) Or _
        (btnSquare(2).Caption = player And btnSquare(5).Caption = player And btnSquare(8).Caption = player) Or _
        (btnSquare(0).Caption = player And btnSquare(4).Caption = player And btnSquare(8).Caption = player) Or _
        (btnSquare(2).Caption = player And btnSquare(4).Caption = player And btnSquare(6).Caption = player) Then
        IsWinner = True
    Else
        IsWinner = False
    End If
End Function

Private Sub Winner()
    For intLoop = 0 To 8
        btnSquare(intLoop).Enabled = False
    Next
    MsgBox "Congradulations! " & lblCurrentPlayer.Caption & " is the winner!", vbOKOnly, "Winner!"
    lblPlayerMessage.Caption = "Winner: "
End Sub

