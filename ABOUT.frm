VERSION 5.00
Begin VB.Form ABOUT 
   BackColor       =   &H00000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de BioClock"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "ABOUT.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"ABOUT.frx":0CCA
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1950
      TabIndex        =   1
      Top             =   2100
      Width           =   5790
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freeware, by Uvchenko Prog. Version: 3.0             - http://www.meetfinder.ar.tc"
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   3300
      Width           =   6210
   End
   Begin VB.Image Image2 
      Height          =   3300
      Left            =   1800
      Top             =   0
      Width           =   6210
   End
   Begin VB.Image Image1 
      Height          =   3570
      Left            =   0
      Picture         =   "ABOUT.frx":0D80
      ToolTipText     =   "Fucci, Leonardo - Programador"
      Top             =   0
      Width           =   1740
   End
End
Attribute VB_Name = "ABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Me.Hide

End Sub

Private Sub Image2_Click()
Me.Hide

End Sub

Private Sub Label1_Click()
Me.Hide

End Sub
