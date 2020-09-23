VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1575
      Top             =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando Datos"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ADIWARE Programacion - Fucci Leonardo"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   2925
      Width           =   5940
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   2700
      Width           =   5940
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path + "\BC.BMP")
ABOUT.Image2.Picture = LoadPicture(App.Path + "\BC.BMP")
Load Form2
Timer1.Enabled = True


End Sub

Private Sub Timer1_Timer()
Form2.Show
Me.Hide
Timer1.Enabled = False
End Sub
