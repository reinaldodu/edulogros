VERSION 5.00
Begin VB.Form ELIMNAR_CURSO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar grupo"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "ELIMNAR_CURSO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   630
      End
   End
End
Attribute VB_Name = "ELIMNAR_CURSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text <> Combo1.List(0) Then
    Combo1.Text = Combo1.List(0)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
'Dim icur As inforcur
'Dim alugru As grupoalu
'Dim aluper As pertgrup
'Dim argra As areagr
grune = False
If Dir(Ruta & Combo1.Text & ".gru") <> "" Then
    NAR = FreeFile
    If Dir(Ruta & "AREAGRA.EDU") <> "" Then
        cona = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            cona = cona + 1
            Get #NAR, cona, argra
            If RTrim(argra.nom_grup) = Combo1.Text Then
                MsgBox "GRUPO NO SE PUEDE ELIMINAR, EXISTEN AREAS CREADAS PARA ESTE", 32, "ADVERTENCIA"
                Close #NAR
                Exit Sub
            End If
        Wend
        Close #NAR
    End If
    RESP = MsgBox("DESEA ELIMINAR EL GRUPO " & Combo1.Text & "?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR GRUPO")
    If RESP = vbYes Then
        Open Ruta & "infcur.edu" For Input As #NAR
        While Not EOF(NAR)
            Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
            If RTrim(icur.nom) <> Combo1.Text Then
                grune = True
                NAR = FreeFile
                Open Ruta & "infcur2.edu" For Append As #NAR
                Write #NAR, icur.nom, icur.jornada, icur.grado, icur.director
                Close #NAR
                NAR = NAR - 1
            End If
        Wend
        Close #NAR
        Kill Ruta & "INFCUR.EDU"
        If grune = True Then
            Name Ruta & "INFCUR2.EDU" As Ruta & "INFCUR.EDU"
        End If
        leo = 0
        Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
        While Not EOF(NAR)
            leo = leo + 1
            Get #NAR, leo, alugru
        Wend
        Close #NAR
        Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
        NAR = FreeFile
        Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
        For I = 1 To (leo - 1)
            Get #(NAR - 1), I, alugru
            aluper.grupo = "PENDIENTE"
            Put #NAR, Val(alugru.num_carnet), aluper
        Next I
        Close #(NAR - 1)
        Close #(NAR)
        Kill Ruta & Combo1.Text & ".gru"
        Unload Me
        MsgBox "GRUPO ELIMINADO", 16, "ELIMINAR"
    End If
Else
MsgBox "NO EXISTE ESTE GRUPO", 48, "ELIMINAR GRUPO"
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Elimina un grupo del sistema."
End Sub

Private Sub Form_Load()
'Dim icur As inforcur
If Dir(Ruta & "infcur.edu") <> "" Then
Command1.Enabled = True
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
Combo1.AddItem RTrim(icur.nom)
Wend
Close #NAR
Combo1.Text = Combo1.List(0)
Else
Command1.Enabled = False
End If
End Sub
