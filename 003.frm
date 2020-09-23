VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "RESULTADO"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   Icon            =   "003.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5745
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6525
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6225
      Top             =   4875
   End
   Begin VB.TextBox Text1 
      Height          =   690
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5850
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   5640
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   7590
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
Me.Show
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()

Print Text1
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Me.PrintForm
Timer2.Enabled = False
End Sub
