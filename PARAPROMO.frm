VERSION 5.00
Begin VB.Form PARAPROMO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de promoción"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "PARAPROMO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Información de acuerdo a los parámetros de promoción"
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   5175
      Begin VB.CommandButton Command5 
         Caption         =   "&Información predeterminada"
         Height          =   320
         Left            =   2760
         TabIndex        =   20
         ToolTipText     =   "Muestra la información original"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "G&uardar Información"
         Height          =   320
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Guarda la información de los parámetros de promoción"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   4815
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "No promovido:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Pendiente:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Promovido:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Parámetros predeterminados"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      ToolTipText     =   "Vuelve los parámetros a su estado original"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Guardar parámetros"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Guarda los parámetros que se muestran en pantalla"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.VScrollBar VScroll2 
         Height          =   295
         Left            =   3360
         TabIndex        =   4
         Top             =   840
         Width           =   135
      End
      Begin VB.TextBox Text2 
         Height          =   295
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   295
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   135
      End
      Begin VB.TextBox Text1 
         Height          =   295
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "materia(s) perdida(s)."
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "materia(s) perdida(s)."
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No promovido: --->"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pendiente:       --->"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Promovido:      --->"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1290
      End
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   16  'Merge Pen
      X1              =   0
      X2              =   5280
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "PARAPROMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RESP = MsgBox("Desea guardar los parámetros de promoción?", vbYesNo + vbQuestion + vbDefaultButton1, "Guardar")
If RESP = vbYes Then
    I = VScroll1.Value
    J = VScroll2.Value
    NAR = FreeFile
    Open Ruta & "rangpro.txt" For Output As #NAR
    Write #NAR, I, J
    Close #NAR
End If
End Sub

Private Sub Command2_Click()
VALI2 = True
VScroll1.Value = I
VScroll2.Value = J
VALI2 = False
End Sub

Private Sub Command4_Click()
RESP = MsgBox("Desea guardar la información de los parámetros de promoción?", vbYesNo + vbQuestion + vbDefaultButton1, "Guardar")
If RESP = vbYes Then
    SAPO2 = Text3.Text
    SAPO3 = Text4.Text
    SAPO4 = Text5.Text
    NAR = FreeFile
    Open Ruta & "promovido.txt" For Output As #NAR
    Write #NAR, SAPO2, SAPO3, SAPO4
    Close #NAR
End If
End Sub

Private Sub Command5_Click()
Text3.Text = SAPO2
Text4.Text = SAPO3
Text5.Text = SAPO4
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "La información de parámetros de promoción saldrá en el informe final de acuerdo al número de áreas perdidas."
End Sub

Private Sub Form_Load()
If (Dir(Ruta & "promovido.txt") <> "") And (Dir(Ruta & "rangpro.txt") <> "") Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    VScroll1.Enabled = True
    VScroll2.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    VScroll1.Min = 0
    VScroll1.Max = 24
    VScroll2.Min = 0
    VScroll2.Max = 25
    VScroll1.SmallChange = 1
    VScroll2.SmallChange = 1
    NAR = FreeFile
    Open Ruta & "rangpro.txt" For Input As #NAR
    Input #NAR, I, J
    Close #NAR
    Open Ruta & "promovido.txt" For Input As #NAR
    Input #NAR, SAPO2, SAPO3, SAPO4
    Close #NAR
    VALI2 = True
    VScroll1.Value = I
    VScroll2.Value = J
    If VScroll1.Value = 0 Then
        Text1.Text = VScroll1.Value
    End If
    If VScroll2.Value = 0 Then
        Text2.Text = VScroll2.Value
        Label6.Caption = " Desde: (0) Hasta:"
        Label7.Caption = " Desde: (1) área(s) perdida(s) en adelante."
    End If
    Text3.Text = SAPO2
    Text4.Text = SAPO3
    Text5.Text = SAPO4
    Label4.Caption = " Desde: (0) Hasta:"
Else
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    VScroll1.Enabled = False
    VScroll2.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Label4.Caption = " Desde: (0) Hasta:"
    Label6.Caption = " Desde: (0) Hasta:"
    Label7.Caption = " Desde: (0) área(s) perdida(s) en adelante."
End If
VALI2 = False
End Sub

Private Sub VScroll1_Change()
Text1.Text = VScroll1.Value
Label6.Caption = " Desde: (" & Val(Text1.Text) + 1 & ") Hasta:"
If VScroll1.Value >= VScroll2.Value Then
    VScroll2.Value = VScroll1.Value + 1
End If
End Sub

Private Sub VScroll2_Change()
Text2.Text = VScroll2.Value
Label7.Caption = " Desde: (" & Val(Text2.Text) + 1 & ") área(s) perdida(s) en adelante."
If VScroll2.Value <= VScroll1.Value Then
    If VScroll2.Value <> 0 Then
        VScroll1.Value = VScroll2.Value - 1
    Else
        Label6.Caption = " Desde: (0) Hasta:"
    End If
End If
If (VScroll1.Value = 0) And (VScroll2.Value > 0) Then
    Label6.Caption = " Desde: (1) Hasta:"
End If
End Sub
