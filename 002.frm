VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIOCLOCK"
   ClientHeight    =   5325
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3975
   Icon            =   "002.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Recalcular"
      Height          =   390
      Left            =   375
      TabIndex        =   34
      Top             =   4875
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   2175
      TabIndex        =   33
      Top             =   4875
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Default         =   -1  'True
      Height          =   390
      Left            =   1500
      TabIndex        =   4
      Top             =   975
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1500
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1500
      TabIndex        =   2
      Top             =   300
      Width           =   2415
   End
   Begin VB.Label Label25 
      Caption         =   "7"
      Height          =   240
      Left            =   75
      TabIndex        =   32
      Top             =   4125
      Width           =   165
   End
   Begin VB.Label Label24 
      Caption         =   "8"
      Height          =   240
      Left            =   75
      TabIndex        =   31
      Top             =   4500
      Width           =   165
   End
   Begin VB.Label D13 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   375
      TabIndex        =   30
      Top             =   4125
      Width           =   1665
   End
   Begin VB.Label D15 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   375
      TabIndex        =   29
      Top             =   4500
      Width           =   1665
   End
   Begin VB.Label D14 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   2175
      TabIndex        =   28
      Top             =   4125
      Width           =   1665
   End
   Begin VB.Label D16 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   2175
      TabIndex        =   27
      Top             =   4500
      Width           =   1665
   End
   Begin VB.Label Label19 
      Caption         =   "5"
      Height          =   240
      Left            =   75
      TabIndex        =   26
      Top             =   3375
      Width           =   165
   End
   Begin VB.Label Label18 
      Caption         =   "6"
      Height          =   240
      Left            =   75
      TabIndex        =   25
      Top             =   3750
      Width           =   165
   End
   Begin VB.Label D9 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   375
      TabIndex        =   24
      Top             =   3375
      Width           =   1665
   End
   Begin VB.Label D11 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   375
      TabIndex        =   23
      Top             =   3750
      Width           =   1665
   End
   Begin VB.Label D10 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   2175
      TabIndex        =   22
      Top             =   3375
      Width           =   1665
   End
   Begin VB.Label D12 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   2175
      TabIndex        =   21
      Top             =   3750
      Width           =   1665
   End
   Begin VB.Label Label13 
      Caption         =   "3"
      Height          =   240
      Left            =   75
      TabIndex        =   20
      Top             =   2625
      Width           =   165
   End
   Begin VB.Label Label12 
      Caption         =   "4"
      Height          =   240
      Left            =   75
      TabIndex        =   19
      Top             =   3000
      Width           =   165
   End
   Begin VB.Label D5 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   375
      TabIndex        =   18
      Top             =   2625
      Width           =   1665
   End
   Begin VB.Label D7 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   375
      TabIndex        =   17
      Top             =   3000
      Width           =   1665
   End
   Begin VB.Label D6 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   2175
      TabIndex        =   16
      Top             =   2625
      Width           =   1665
   End
   Begin VB.Label D8 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   2175
      TabIndex        =   15
      Top             =   3000
      Width           =   1665
   End
   Begin VB.Line Line2 
      X1              =   300
      X2              =   300
      Y1              =   4800
      Y2              =   1725
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3900
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Label D4 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   2175
      TabIndex        =   14
      Top             =   2250
      Width           =   1665
   End
   Begin VB.Label D2 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   2175
      TabIndex        =   13
      Top             =   1875
      Width           =   1665
   End
   Begin VB.Label D3 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   375
      TabIndex        =   12
      Top             =   2250
      Width           =   1665
   End
   Begin VB.Label D1 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   375
      TabIndex        =   11
      Top             =   1875
      Width           =   1665
   End
   Begin VB.Label Label7 
      Caption         =   "2"
      Height          =   240
      Left            =   75
      TabIndex        =   10
      Top             =   2250
      Width           =   165
   End
   Begin VB.Label Label6 
      Caption         =   "1"
      Height          =   240
      Left            =   75
      TabIndex        =   9
      Top             =   1875
      Width           =   165
   End
   Begin VB.Label Label5 
      Caption         =   "Fin:"
      Height          =   240
      Left            =   2175
      TabIndex        =   8
      Top             =   1425
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Comienzo:"
      Height          =   240
      Left            =   375
      TabIndex        =   7
      Top             =   1425
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   240
      Left            =   0
      TabIndex        =   6
      Top             =   75
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Nacimiento:"
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   675
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Actual:"
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   1440
   End
   Begin VB.Menu q1 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
   End
   Begin VB.Menu q2 
      Caption         =   "Recalcular"
      Enabled         =   0   'False
   End
   Begin VB.Menu q3 
      Caption         =   "Ayuda"
      Begin VB.Menu Q4 
         Caption         =   "Que es el BIORRITMO"
      End
      Begin VB.Menu SEP1 
         Caption         =   "-"
      End
      Begin VB.Menu a 
         Caption         =   "Acerca de BioClock"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click()
ABOUT.Show

End Sub

Private Sub Command1_Click()
On Error GoTo FLOR
Dim DIASVIVIDOS As Integer
Dim CANTIDADPERIODOS As Integer
Dim ULTIMOPERIODO As Integer
Dim COMIENZO As Date

DIASVIVIDOS = Date - CDate(Text2)
CANTIDADPERIODOS = Int(DIASVIVIDOS / 28)
ULTIMOPERIODO = CANTIDADPERIODOS * 28
COMIENZO = ULTIMOPERIODO + CDate(Text2)
D1 = COMIENZO
D2 = CDate(D1) + 15
D3 = CDate(D2) + 15
D4 = CDate(D3) + 15
D5 = CDate(D4) + 15
D6 = CDate(D5) + 15
D7 = CDate(D6) + 15
D8 = CDate(D7) + 15
D9 = CDate(D8) + 15
D10 = CDate(D9) + 15
D11 = CDate(D10) + 15
D12 = CDate(D11) + 15
D13 = CDate(D12) + 15
D14 = CDate(D13) + 15
D15 = CDate(D14) + 15
D16 = CDate(D15) + 15
Command1.Visible = False
q1.Enabled = True
q2.Enabled = True
Me.Height = 5985
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True

FLOR:
    If Err.Number = 13 Then
        MsgBox "La fecha introducida no es correcta.", vbExclamation, "Error al introducir fecha de nacimiento."
        Text2 = ""
    End If
End Sub

Private Sub Command2_Click()
Form3.Label1 = "Los periodos ideales en los que " + Text3 + " puede tener mas rendimiento son los siguientes." _
+ vbNewLine + "         INICIO    |    FIN" + vbNewLine + D1 + "    |    " + D2 _
+ vbNewLine + D3 + "    |    " + D4 _
+ vbNewLine + D5 + "    |    " + D6 _
+ vbNewLine + D7 + "    |    " + D8 _
+ vbNewLine + D9 + "    |    " + D10 _
+ vbNewLine + D11 + "    |    " + D12 _
+ vbNewLine + D13 + "    |    " + D14 _
+ vbNewLine + D15 + "    |    " + D16 _
+ vbNewLine + Text3 + " lleva " & Date - CDate(Text2) & " dias vividos." _
+ vbNewLine + vbNewLine + "¿En que consiste el Biorritmo.?" + vbNewLine _
+ "El BIORRITMO, es un intervalo de 15 dias en los que el cuerpo humano trabaja al 110%, o sea, toda actividad realizada tendra resultados satisfactorios, como entrenar o adelgazar. Pero de esta misma manera, si se come mucho en este periodo, engordara mucho. Por eso, cuando se encuentre fuera del periodo de BIORRITMO, puede comer normalmente, ya que no tendra mucho efecto, al igual que NO tendra mucho efecto entrenar demasiado, lo recomendable seria entrenar lo suficiente como para no perder la constumbre." + vbNewLine _
+ "O sea, el cuerpo en el periodo de BIORRITMO, tiende a reflejar optimos resultados de las actividades que se hagan en este periodo, como adelgazar, engordar, entrenar, etc. Y fuera de este, lo contrario."

Form3.Show
Form3.Timer2.Enabled = True

End Sub

Private Sub Command3_Click()
Me.Height = 2055
Command1.Visible = True
Text2 = ""
Text3 = ""
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
q1.Enabled = False
q2.Enabled = False

End Sub

Private Sub Command4_Click()
Text4 = Date - CDate(Text2)

End Sub

Private Sub Form_Load()
Text1 = Date
Me.Height = 2055
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload ABOUT
Unload Form1
Unload Form3
Unload Form4
Unload Me
End

End Sub

Private Sub q1_Click()
Command2_Click
End Sub

Private Sub q2_Click()
Command3_Click
End Sub

Private Sub Q4_Click()
Form4.Show

End Sub
