VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.PictureBox butblue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   285
      ScaleWidth      =   1260
      TabIndex        =   2
      Top             =   0
      Width           =   1260
   End
   Begin VB.PictureBox butblueover 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      Picture         =   "Form1.frx":12F6
      ScaleHeight     =   285
      ScaleWidth      =   1260
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox butgreen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      Picture         =   "Form1.frx":25EE
      ScaleHeight     =   285
      ScaleWidth      =   1260
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label mnuexit 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "   Exit"
      Height          =   255
      Left            =   200
      TabIndex        =   3
      Top             =   450
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Image mnu1 
      Height          =   2655
      Left            =   120
      Picture         =   "Form1.frx":38E4
      Top             =   285
      Visible         =   0   'False
      Width           =   1785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code is written by Naveed Ahmed, 2002
' you are allowed to use this code
' in your programs however you wish to.
' You can edit the code and you do not have
' to give me any credit whatsoever.

Private Sub butblue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' When mouse moves over the menu
butblue.Visible = False
butblueover.Visible = True
butgreen.Visible = False
End Sub
Private Sub butblueover_Click()
' When menu opened
butblueover.Visible = False
butgreen.Visible = True
butblue.Visible = False
mnu1.Visible = True
If mnu1.Visible = True Then mnuexit.Visible = True
End Sub
Private Sub butgreen_Click()
' While the menu is open
mnu1.Visible = False
If mnu1.Visible = False Then mnuexit.Visible = False
butgreen.Visible = False
End Sub

Private Sub butgreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.ForeColor = &H80000012 ' Black
mnuexit.BackStyle = 0
End Sub

Private Sub Form_Click()
' To close menu
butblueover.Visible = False
butgreen.Visible = False
butblue.Visible = False
mnu1.Visible = False
If mnu1.Visible = False Then mnuexit.Visible = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' When the menu is open and mouse moves on form
butblue.Visible = True
butblueover.Visible = False
If butgreen.Visible = True Then butblue.Visible = False
mnuexit.ForeColor = &H80000012 ' Black
mnuexit.BackStyle = 0
End Sub
Private Sub mnu1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.ForeColor = &H80000012 ' Black
mnuexit.BackStyle = 0
End Sub

Private Sub mnuexit_Click()
' Exit the program
End
End Sub

Private Sub mnuexit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuexit.ForeColor = &H80000005 ' White
mnuexit.BackStyle = 1
End Sub
