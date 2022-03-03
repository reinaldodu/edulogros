VERSION 5.00
Begin VB.Form COPIHIST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "COPIHIST.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el periodo a copiar"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "COPIHIST.frx":0442
         Left            =   960
         List            =   "COPIHIST.frx":0455
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PERIODO:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   780
      End
   End
End
Attribute VB_Name = "COPIHIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
Dim LongNom As Single
Clipboard.Clear
cop = ""
Printer.ScaleMode = 7
For I = 1 To (nf - 1)
        RESUFINA.MATI50.Row = I
        RESUFINA.MATI50.Col = 0
        nom = RESUFINA.MATI50.Text
        LongNom = Printer.TextWidth(nom)
        While LongNom < 6
            nom = nom & Chr(9)
            LongNom = Printer.TextWidth(nom)
        Wend
        RESUFINA.MATI50.Col = 1
        cort = RESUFINA.MATI50.Text
        If RTrim(Combo1.Text) = "FINAL" Then
           RESUFINA.MATI50.Col = 6
        End If
        If RTrim(Combo1.Text) = "CUARTO" Then
           RESUFINA.MATI50.Col = 5
        End If
        If RTrim(Combo1.Text) = "TERCERO" Then
           RESUFINA.MATI50.Col = 4
        End If
        If RTrim(Combo1.Text) = "SEGUNDO" Then
           RESUFINA.MATI50.Col = 3
        End If
        If RTrim(Combo1.Text) = "PRIMERO" Then
           RESUFINA.MATI50.Col = 2
        End If
        jho = RESUFINA.MATI50.Text
        cop = cop + nom & cort & Chr(9) & Chr(9) & jho & vbCrLf
Next I
Clipboard.SetText cop
Unload Me
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Copia las notas obtenidas por el alumno de acuerdo al periodo académico seleccionado."
End Sub

Private Sub Form_Load()
Combo1.Text = Combo1.List(0)
End Sub
