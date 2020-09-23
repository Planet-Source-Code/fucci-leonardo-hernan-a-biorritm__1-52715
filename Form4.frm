VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Que es el Biorritmo"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   3540
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6240
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1 = vbNewLine + "¿En que consiste el Biorritmo.?" + vbNewLine _
+ "El BIORRITMO, es un intervalo de 15 dias en los que el cuerpo humano trabaja al 110%, o sea, toda actividad realizada tendra resultados satisfactorios, como entrenar o adelgazar. Pero de esta misma manera, si se come mucho en este periodo, engordara mucho. Por eso, cuando se encuentre fuera del periodo de BIORRITMO, puede comer normalmente, ya que no tendra mucho efecto, al igual que NO tendra mucho efecto entrenar demasiado, lo recomendable seria entrenar lo suficiente como para no perder la constumbre." + vbNewLine _
+ "O sea, el cuerpo en el periodo de BIORRITMO, tiende a reflejar optimos resultados de las actividades que se hagan en este periodo, como adelgazar, engordar, entrenar, etc. Y fuera de este, lo contrario."

End Sub
