VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form IMP_GRUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir grupo"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6255
   Icon            =   "IMP_GRUP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "IMPRIMIR CONTROL DE LOGROS"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         Height          =   320
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   630
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   320
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupos existentes"
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin MSFlexGridLib.MSFlexGrid MATI4 
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         ForeColor       =   4194368
         BackColorFixed  =   16777215
         ForeColorFixed  =   4194368
         GridColor       =   12582912
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total de grupos..."
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "IMPRIMIR LISTA DETALLADA"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   2400
      TabIndex        =   6
      Top             =   1320
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   320
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   630
      End
   End
End
Attribute VB_Name = "IMP_GRUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim icur As inforcur
'Dim alugru As grupoalu
'Dim profe As maestropro
'Dim ini As inicio
If RTrim(Text1.Text) = "" Then
    MsgBox "ESCRIBA EL NOMBRE DEL GRUPO", 16, "IMPRIMIR GRUPO"
    Exit Sub
End If
Text1.Text = Format(Text1.Text, ">")
If Dir(Ruta & Text1.Text & ".gru") = "" Then
    MsgBox "NO EXISTE ESTE GRUPO", 48, "IMPRIMIR GRUPO"
    Text1.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTE GRUPO?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR GRUPO")
If RESP = vbYes Then
Screen.MousePointer = 11
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
PAG = 1
Printer.ScaleMode = 7
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "Pág." & PAG
Printer.CurrentY = 2.5
Printer.CurrentX = 0.5
Printer.Font.Size = 10
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
If RTrim(icur.nom) = RTrim(Text1.Text) Then
dire = icur.director
YUS = RTrim(icur.jornada)
End If
Wend
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, dire, profe
Close #NAR
Printer.Print "Directora: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
Printer.CurrentY = 1.5
Printer.CurrentX = 6
Printer.Print "ESTUDIANTES JORNADA " & YUS & " GRUPO " & Text1.Text
Printer.CurrentY = 3
Printer.CurrentX = 0.5
Printer.Print ini.nombre
Printer.CurrentY = 4
Printer.CurrentX = 0.5
Printer.Font.Underline = True
Printer.Font.Size = 8
Printer.Print "MATRIC.";
Printer.CurrentX = 2
Printer.Print "CARNET.";
Printer.CurrentX = 3.5
Printer.Print "COD";
Printer.CurrentX = 4.5
Printer.Print "APELLIDOS Y NOMBRES";
Printer.CurrentX = 10.5
Printer.Print "FECH_NACIM";
Printer.CurrentX = 12.7
Printer.Print "EDAD";
Printer.CurrentX = 13.7
Printer.Print "ACUDIENTE";
Printer.CurrentX = 18.7
Printer.Print "TELEFONO"
Printer.Font.Underline = False
Printer.Font.Size = 8
Open Ruta & Text1.Text & ".gru" For Random As #NAR Len = Len(alugru)
leo = 0
While Not EOF(NAR)
leo = leo + 1
Get #NAR, leo, alugru
Wend
Close #NAR
Open Ruta & Text1.Text & ".gru" For Random As #NAR Len = Len(alugru)
NAR = FreeFile
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
For rr = 1 To leo - 1
Get #(NAR - 1), rr, alugru
Get #NAR, (Val(alugru.num_carnet)), alumno
Printer.CurrentX = 0.5
Printer.Print alumno.n_matricula;
Printer.CurrentX = 2
Printer.Print alumno.n_carnet;
Printer.CurrentX = 3.5
Printer.Print rr;
Printer.CurrentX = 4.5
Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
Printer.CurrentX = 10.5
Printer.Print alumno.f_nacimiento;
Printer.CurrentX = 12.7
dd = Val(Left(alumno.f_nacimiento, 2))
mm2 = Right(alumno.f_nacimiento, 7)
mm = Val(Left(mm2, 2))
aaaa = Val(Right(alumno.f_nacimiento, 4))
aaaa = Year(Date) - aaaa
If mm > Month(Date) Then
aaaa = aaaa - 1
End If
If mm = Month(Date) Then
   If dd > Day(Date) Then
   aaaa = aaaa - 1
   End If
End If
Printer.Print aaaa;
Printer.CurrentX = 13.7
Printer.Print alumno.acudiente;
Printer.CurrentX = 18.7
Printer.Print alumno.tel_acu
If (rr Mod 65) = 0 Then
Printer.NewPage
PAG = PAG + 1
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "Pág." & PAG
Printer.CurrentY = 2.5
Printer.CurrentX = 0.5
Printer.Font.Size = 10
Printer.Print "Directora: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
Printer.CurrentY = 1.5
Printer.CurrentX = 6
Printer.Print "ESTUDIANTES JORNADA " & YUS & " GRUPO " & Text1.Text
Printer.CurrentY = 3
Printer.CurrentX = 0.5
Printer.Print ini.nombre
Printer.CurrentY = 4
Printer.CurrentX = 0.5
Printer.Font.Underline = True
Printer.Font.Size = 8
Printer.Print "MATRICULA";
Printer.CurrentX = 2
Printer.Print "CARNET.";
Printer.CurrentX = 3.5
Printer.Print "COD";
Printer.CurrentX = 4.5
Printer.Print "APELLIDOS Y NOMBRES";
Printer.CurrentX = 10.5
Printer.Print "FECH_NACIM";
Printer.CurrentX = 12.7
Printer.Print "EDAD";
Printer.CurrentX = 13.7
Printer.Print "ACUDIENTE";
Printer.CurrentX = 18.7
Printer.Print "TELEFONO"
Printer.Font.Underline = False
Printer.Font.Size = 8
End If
Next rr
Close #(NAR - 1)
Close #NAR
Printer.EndDoc
End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
'Dim icur As inforcur
If RTrim(Text3.Text) = "" Then
    MsgBox "ESCRIBA EL NOMBRE DEL GRUPO", 16, "IMPRIMIR GRUPO"
    Exit Sub
End If
Text3.Text = Format(Text3.Text, ">")
If Dir(Ruta & Text3.Text & ".gru") = "" Then
    MsgBox "NO EXISTE ESTE GRUPO", 48, "IMPRIMIR GRUPO"
    Text3.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
If RTrim(icur.nom) = RTrim(Text3.Text) Then
JOJI = RTrim(icur.jornada)
End If
Wend
Close #NAR
LIST_LOGRO.Label5.Caption = JOJI
LIST_LOGRO.Frame2.Caption = RTrim(Text3.Text)
LIST_LOGRO.Show 1
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Impresión de listas por grupos."
End Sub

Private Sub Form_Load()
'Dim icur As inforcur
MATI4.ColWidth(0) = 1500
Text1.MaxLength = 13
Text3.MaxLength = 13
plo = 0
If Dir(Ruta & "infcur.edu") <> "" Then
    Command1.Enabled = True
    Command2.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        MATI4.Rows = plo + 1
        MATI4.TextMatrix(plo, 0) = icur.nom
        plo = plo + 1
    Wend
    Close #NAR
Else
    Command1.Enabled = False
    Command2.Enabled = False
End If
Text2.Text = plo
End Sub

Private Sub MATI4_DblClick()
Text3.Text = RTrim(MATI4.Text)
Text1.Text = RTrim(MATI4.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Command2.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command2_Click
End If
End Sub
