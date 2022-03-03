VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form TOTALES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de totales"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9195
   Icon            =   "TOTALES.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   975
      Left            =   7560
      Picture         =   "TOTALES.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "OPCIONES"
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   5535
      Begin VB.OptionButton Option4 
         Caption         =   "ALUMNOS POR GRUPO."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "MATERIAS POR PROFESOR"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "MATERIAS POR GRUPO"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "VARIOS."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin MSFlexGridLib.MSFlexGrid MATI80 
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
   End
End
Attribute VB_Name = "TOTALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim icur As inforcur
'Dim mate As infomater
'Dim alugru As grupoalu
'Dim argra As areagr
'Dim notas As notis
'Dim profe As maestropro
If Option1.Value = True Then
Screen.MousePointer = 11
MATI80.Cols = 2
MATI80.Rows = 7
MATI80.FixedRows = 1
MATI80.Row = 0
MATI80.Col = 0
MATI80.ColWidth(0) = 3000
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(255, 255, 255)
MATI80.CellBackColor = RGB(0, 0, 150)
MATI80.Text = "                 VARIOS"
MATI80.Col = 1
MATI80.ColWidth(1) = 1000
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(255, 255, 255)
MATI80.CellBackColor = RGB(0, 0, 150)
MATI80.Text = "   TOTAL"
MATI80.Col = 0
MATI80.Row = 1
MATI80.Text = "NÑOS EXISTENTES..."
MATI80.Col = 0
MATI80.Row = 2
MATI80.Text = "NIÑAS EXISTENTES..."
MATI80.Col = 0
MATI80.Row = 3
MATI80.Text = "TOTAL ESTUDIANTES..."
'MATI80.Col = 0
'MATI80.Row = 4
'MATI80.Text = "ESTUDIANTES RETIRADO(A)S..."
MATI80.Col = 0
MATI80.Row = 4
MATI80.Text = "PROFESORES EXISTENTES..."
'MATI80.Col = 0
'MATI80.Row = 6
'MATI80.Text = "PROFESORES RETIRADOS..."
MATI80.Col = 0
MATI80.Row = 5
MATI80.Text = "GRUPOS EXISTENTES..."
MATI80.Col = 0
MATI80.Row = 6
MATI80.Text = "MATERIAS EXISTENTES..."
NAR = FreeFile
q = 0
CM = 0
CH = 0
plo = 2
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
While Not EOF(NAR)
q = q + 1
Get #NAR, q, alumno
If RTrim(alumno.sexo) = "M" Then
CH = CH + 1
End If
If RTrim(alumno.sexo) = "F" Then
CM = CM + 1
End If
Wend
Close #NAR
MATI80.Col = 1
MATI80.Row = 1
MATI80.Text = CH
MATI80.Row = plo
MATI80.Text = CM
plo = plo + 1
MATI80.Row = plo
MATI80.Text = CM + CH
'Open Ruta & "contreti.edu" For Input As #NAR
'Input #NAR, zi
'Close #NAR
'Open Ruta & "conelire.edu" For Input As #NAR
'Input #NAR, z
'Close #NAR
'plo = plo + 1
'MATI80.Row = plo
'MATI80.Text = (zi - 1) - z
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
sir = 0
SIRO = 0
Open Ruta & "infcdpro.edu" For Random As #NAR Len = 2
While Not EOF(NAR)
sir = sir + 1
Get #NAR, sir, clat
If clat <> 0 Then
SIRO = SIRO + 1
End If
Wend
Close #NAR
plo = plo + 1
MATI80.Row = plo
MATI80.Text = (r - 1) - SIRO
'Open Ruta & "conrepro.edu" For Input As #NAR
'Input #NAR, zu
'Close #NAR
'plo = plo + 1
'MATI80.Row = plo
'MATI80.Text = zu - 1
cli = 0
If Dir(Ruta & "infcur.edu") <> "" Then
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
cli = cli + 1
Wend
Close #NAR
End If
MATI80.Row = 5
MATI80.Text = cli
h = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
que = 0
While Not EOF(NAR)
que = que + 1
Get #NAR, que, mate
If RTrim(mate.nom) = "" Then
h = h + 1
End If
Wend
Close #NAR
MATI80.Row = 6
MATI80.Text = (que - 1) - h
Screen.MousePointer = 0
End If
If Option4.Value = True Then
Screen.MousePointer = 11
MATI80.Cols = 3
MATI80.Rows = 1
MATI80.Row = 0
MATI80.Col = 0
MATI80.ColWidth(0) = 2000
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(255, 255, 255)
MATI80.CellBackColor = RGB(0, 0, 150)
MATI80.Text = "          GRADO"
MATI80.Col = 1
MATI80.ColWidth(1) = 2500
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(255, 255, 255)
MATI80.CellBackColor = RGB(0, 0, 150)
MATI80.Text = "          GRUPO"
MATI80.Col = 2
MATI80.ColWidth(2) = 1000
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(255, 255, 255)
MATI80.CellBackColor = RGB(0, 0, 150)
MATI80.Text = "   TOTAL"
If Dir(Ruta & "infcur.edu") <> "" Then
NAR = FreeFile
plo = 2
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
MATI80.Rows = plo
If plo = 2 Then
MATI80.FixedRows = 1
End If
MATI80.Row = plo - 1
MATI80.Col = 0
MATI80.Text = RTrim(icur.grado)
MATI80.Col = 1
MATI80.Text = RTrim(icur.nom)
ret = 0
NAR = FreeFile
Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
ret = ret + 1
Get #NAR, ret, alugru
Wend
Close #NAR
NAR = NAR - 1
MATI80.Col = 2
MATI80.Text = ret - 1
plo = plo + 1
Wend
Close #NAR
End If
Screen.MousePointer = 0
End If
If Option2.Value = True Then
Screen.MousePointer = 11
MATI80.Cols = 3
MATI80.Rows = 1
MATI80.Row = 0
MATI80.Col = 0
MATI80.ColWidth(0) = 2500
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(255, 255, 255)
MATI80.CellBackColor = RGB(0, 0, 150)
MATI80.Text = "         GRUPO"
MATI80.Col = 1
MATI80.ColWidth(1) = 5200
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(255, 255, 255)
MATI80.CellBackColor = RGB(0, 0, 150)
MATI80.Text = "                 MATERIAS"
MATI80.Col = 2
MATI80.ColWidth(2) = 4000
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(255, 255, 255)
MATI80.CellBackColor = RGB(0, 0, 150)
MATI80.Text = "                    PROFESOR"
If Dir(Ruta & "infcur.edu") = "" Then
Screen.MousePointer = 0
Exit Sub
End If
If Dir(Ruta & "areagra.edu") = "" Then
Screen.MousePointer = 0
Exit Sub
End If
If Dir(Ruta & "materia.edu") = "" Then
Screen.MousePointer = 0
Exit Sub
End If
NAR = FreeFile
plo = 2
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
MATI80.Rows = plo
If plo = 2 Then
MATI80.FixedRows = 1
End If
MATI80.Row = plo - 1
MATI80.Col = 0
MATI80.CellFontBold = True
MATI80.Text = RTrim(icur.nom)
NAR = FreeFile
cona = 0
CONTAREA = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
cona = cona + 1
Get #NAR, cona, argra
If RTrim(icur.nom) = RTrim(argra.nom_grup) Then
CONTAREA = CONTAREA + 1
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
Get #NAR, argra.num_area, mate
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, argra.num_pro, profe
Close #NAR
NAR = NAR - 1
MATI80.Rows = plo
MATI80.Row = plo - 1
MATI80.Col = 1
'MATI80.Text = RTrim(mate.nom) & " (" & mate.num & ")"
MATI80.Text = RTrim(mate.nom) & " (I.H:" & argra.ih & ")"
MATI80.Col = 2
MATI80.Text = RTrim(profe.nombres) & " " & RTrim(profe.apellidos) & " (" & argra.num_pro & ")"
plo = plo + 1
End If
Wend
Close #NAR
NAR = NAR - 1
MATI80.Rows = plo
MATI80.Row = plo - 1
MATI80.Col = 1
MATI80.CellFontBold = True
MATI80.CellForeColor = RGB(0, 0, 255)
MATI80.Text = "TOTAL MATERIAS..." & CONTAREA
plo = plo + 1
Wend
Close #NAR
Screen.MousePointer = 0
End If
If Option3.Value = True Then
    Screen.MousePointer = 11
    MATI80.Cols = 3
    MATI80.Rows = 1
    MATI80.Row = 0
    MATI80.Col = 0
    MATI80.ColWidth(0) = 3500
    MATI80.CellFontBold = True
    MATI80.CellForeColor = RGB(255, 255, 255)
    MATI80.CellBackColor = RGB(0, 0, 150)
    MATI80.Text = "                  PROFESOR"
    MATI80.Col = 1
    MATI80.ColWidth(1) = 6000
    MATI80.CellFontBold = True
    MATI80.CellForeColor = RGB(255, 255, 255)
    MATI80.CellBackColor = RGB(0, 0, 150)
    MATI80.Text = "      MATERIA"
    MATI80.Col = 2
    MATI80.ColWidth(2) = 500
    MATI80.CellFontBold = True
    MATI80.CellForeColor = RGB(255, 255, 255)
    MATI80.CellBackColor = RGB(0, 0, 150)
    MATI80.Text = "I.H."
    NAR = FreeFile
    cona2 = 0
    plo = 2
    MATI80.Rows = plo
    MATI80.FixedRows = 1
    NAR = FreeFile
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    While Not EOF(NAR)
        cona2 = cona2 + 1
        Get #NAR, cona2, profe
        NAR = FreeFile
        cona = 0
        MATI80.TextMatrix(plo - 1, 0) = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            cona = cona + 1
            Get #NAR, cona, argra
            If RTrim(argra.num_pro) = cona2 Then
                    NAR = FreeFile
                    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                    Get #NAR, argra.num_area, mate
                    Close #NAR
                    NAR = NAR - 1
                    MATI80.TextMatrix(plo - 1, 1) = RTrim(mate.nom) & " (" & RTrim(argra.nom_grup) & ")"
                    MATI80.TextMatrix(plo - 1, 2) = argra.ih
                    plo = plo + 1
                    MATI80.Rows = plo
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
    Wend
    Close #NAR
    Screen.MousePointer = 0
End If
End Sub

Private Sub Command2_Click()
Printer.ScaleMode = 7
If plo = 2 Then
MsgBox "NO HAY INFORMACION PARA IMPRIMIR", 48
Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR CONTROL DE TOTALES")
If RESP = vbYes Then
If Option1.Value = True Then
Printer.CurrentY = 2
Printer.CurrentX = 7
Printer.Print "T O T A L  V A R I O S"
Printer.CurrentY = 4
For I = 1 To plo + 2
MATI80.Col = 0
MATI80.Row = I
Printer.CurrentX = 5
Printer.Print MATI80.Text;
MATI80.Col = 1
Printer.CurrentX = 13
Printer.Print MATI80.Text
Next I
End If
If Option2.Value = True Then
    PAG = 1
    Printer.CurrentY = 1
    Printer.CurrentX = 18
    Printer.Print "Pág." & PAG
    Printer.CurrentY = 2
    Printer.CurrentX = 7
    Printer.Print "M A T E R I A S   E X I S T E N T E S"
    Printer.CurrentY = 4
    For I = 1 To plo - 2
        If (I Mod 48) = 0 Then
            PAG = PAG + 1
            Printer.NewPage
            Printer.CurrentY = 1
            Printer.CurrentX = 18
            Printer.Print "Pág." & PAG
            Printer.CurrentY = 4
        End If
        Printer.Font.Size = 9
        MATI80.Col = 0
        MATI80.Row = I
        If MATI80.Text <> "" Then
            Printer.Print ""
            Printer.CurrentX = 1
            Printer.Print MATI80.Text
            Printer.Print ""
        Else
            Printer.CurrentX = 1
            Printer.Print MATI80.Text;
        End If
        MATI80.Col = 1
        Printer.CurrentX = 1
        Printer.Print MATI80.Text;
        MATI80.Col = 2
        Printer.CurrentX = 11
        Printer.Print MATI80.Text
    Next I
    Printer.Font.Size = 12
End If

If Option3.Value = True Then
    PAG = 1
    Printer.CurrentY = 1
    Printer.CurrentX = 18
    Printer.Print "Pág." & PAG
    Printer.CurrentY = 2
    Printer.CurrentX = 7
    Printer.Print "M A T E R I A S   P O R   P R O F E S O R"
    Printer.CurrentY = 4
    For I = 1 To plo - 2
        If (I Mod 48) = 0 Then
            PAG = PAG + 1
            Printer.NewPage
            Printer.CurrentY = 1
            Printer.CurrentX = 18
            Printer.Print "Pág." & PAG
            Printer.CurrentY = 4
        End If
        Printer.Font.Size = 9
        MATI80.Col = 0
        MATI80.Row = I
        Printer.CurrentX = 1
        Printer.Print MATI80.Text;
        MATI80.Col = 2
        Printer.CurrentX = 9
        Printer.Print MATI80.Text;
        MATI80.Col = 1
        Printer.CurrentX = 10
        Printer.Print MATI80.Text
    Next I
    Printer.Font.Size = 12
End If



If Option4.Value = True Then
PAG = 1
Printer.CurrentY = 1
Printer.CurrentX = 18
Printer.Print "Pág." & PAG
Printer.CurrentY = 2
Printer.CurrentX = 4
Printer.Print "A L U M N O S  E X I S T E N T E S  P O R  G R U P O"
Printer.CurrentY = 4
Printer.CurrentX = 5
Printer.Print "GRADO";
Printer.CurrentX = 10
Printer.Print "GRUPO";
Printer.CurrentX = 15
Printer.Print "TOTAL"
Printer.Print ""
For I = 1 To plo - 2
If (I Mod 46) = 0 Then
PAG = PAG + 1
Printer.NewPage
Printer.CurrentY = 1
Printer.CurrentX = 18
Printer.Print "Pág." & PAG
Printer.CurrentY = 4
Printer.CurrentX = 5
Printer.Print "GRADO";
Printer.CurrentX = 10
Printer.Print "GRUPO";
Printer.CurrentX = 15
Printer.Print "TOTAL"
Printer.Print ""
End If
MATI80.Col = 0
MATI80.Row = I
Printer.CurrentX = 5
Printer.Print MATI80.Text;
MATI80.Col = 1
Printer.CurrentX = 10
Printer.Print MATI80.Text;
MATI80.Col = 2
Printer.CurrentX = 15
Printer.Print MATI80.Text
Next I
End If
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra información de algunos datos estadísticos del sistema."
End Sub

Private Sub Form_Load()
'Dim icur As inforcur
'If Dir(Ruta & "infcur.edu") <> "" Then
'NAR = FreeFile
'Open Ruta & "infcur.edu" For Input As #NAR
'While Not EOF(NAR)
'Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
'Combo2.AddItem RTrim(icur.nom)
'Wend
'Close #NAR
'Combo2.Text = Combo2.List(0)
'End If
Option1.Value = True
plo = 2
End Sub


Private Sub Option1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (Command1.Enabled = True) Then
    Call Command1_Click
End If
End Sub
Private Sub Option2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (Command1.Enabled = True) Then
    Call Command1_Click
End If
End Sub
Private Sub Option3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (Command1.Enabled = True) Then
    Call Command1_Click
End If
End Sub
Private Sub Option4_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (Command1.Enabled = True) Then
    Call Command1_Click
End If
End Sub
