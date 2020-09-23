VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4095
      TabIndex        =   3
      Top             =   4125
      Width           =   645
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   1830
      TabIndex        =   2
      Top             =   3945
      Width           =   1365
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   465
      Left            =   180
      TabIndex        =   1
      Top             =   3945
      Width           =   1365
   End
   Begin VB.TextBox txtCode 
      Height          =   3570
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "engine.frx":0000
      Top             =   255
      Width           =   6405
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
