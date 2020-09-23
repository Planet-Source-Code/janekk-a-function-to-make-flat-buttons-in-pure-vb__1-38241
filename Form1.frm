VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2340
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   2340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Button4"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Button3"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Button2"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox picFlat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   6240
      ScaleHeight     =   585
      ScaleWidth      =   1665
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Button 1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   
   Call MakeFlatButtons(Me)
   
End Sub


