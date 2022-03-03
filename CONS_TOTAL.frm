VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONS_TTL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por periodos"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "CONS_TOTAL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   7935
      Begin VB.CommandButton Command2 
         Height          =   360
         Left            =   7320
         Picture         =   "CONS_TOTAL.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   315
         Left            =   6120
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "AREA:"
         Height          =   195
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid MATXPER 
         Height          =   5175
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9128
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   45
      End
   End
End
Attribute VB_Name = "CONS_TTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo3.SetFocus
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
MATXPER.Rows = 1
Label1.Caption = ""
Command2.Enabled = False
If Dir(Ruta & Combo2.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
NAR = FreeFile
TN = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    TN = TN + 1
    Get #NAR, TN, mate
    If RTrim(mate.nom) = Combo3.Text Then
        que = mate.num
    End If
Wend
Close #NAR
ret = 0
Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If RTrim(argra.nom_grup) = Combo2.Text And argra.num_area = que Then
        NAR = FreeFile
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, (argra.num_pro), profe
        Close #NAR
        PRO = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        NAR = NAR - 1
        Close #NAR
        GoTo ALTU63
    End If
Wend
Close #NAR
MsgBox "NO SE HA CREADO EL AREA " & Combo3.Text & " PARA ESTE GRUPO", 64, "ADVERTENCIA"
Screen.MousePointer = 0
Exit Sub
ALTU63:
For I = 1 To (ret - 1)
    MATXPER.Rows = I + 1
    MATXPER.TextMatrix(I, 0) = I
    Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
    Get #NAR, I, alugru
    Close #NAR
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, (Val(alugru.num_carnet)), alumno
    Close #NAR
    MATXPER.TextMatrix(I, 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    MATXPER.TextMatrix(I, 2) = alumno.n_carnet
    For lw = 1 To 5
        If Dir(Ruta & Combo2.Text & que & lw & ".obs") <> "" Then
            Open Ruta & Combo2.Text & que & lw & ".obs" For Random As #NAR Len = Len(notas)
            Y = 0
            While Not EOF(NAR)
                Y = Y + 1
                Get #NAR, Y, notas
                If (notas.num_carnet) = Right(alumno.n_carnet, 5) Then
                    'MATXPER.TextMatrix(I, lw + 2) = notas.JV
                    Close #NAR
                    GoTo SALPER
                End If
            Wend
            Close #NAR
        End If
SALPER:
    Next lw
Next I
Command2.Enabled = True
Label1.Caption = "GRUPO: " & Combo2.Text & " - " & " AREA: " & Combo3.Text & " - " & " PROFESOR(A): " & PRO
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
RESP = MsgBox("Desea imprimir esta información?", vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir")
If RESP = vbYes Then
   Screen.MousePointer = 11
   NAR = FreeFile
   Open Ruta & "inicial.edu" For Input As #NAR
   Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
   Close #NAR
   Printer.ScaleMode = 7
   Printer.Font.Size = 10
   Printer.CurrentY = 1
   Printer.CurrentX = 7.5
   Printer.Print "CONSULTA POR PERIODOS"
   Printer.Print ""
   Printer.CurrentX = 0.5
   Printer.Print ini.nombre;
   Printer.CurrentX = 16.5
   Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
   Printer.CurrentX = 0.5
   Printer.Print Label1.Caption
   Printer.Print ""
   Printer.CurrentX = 0.5
   Printer.Print "CD";
   Printer.CurrentX = 1.3
   Printer.Print "APELLIDOS Y NOMBRES";
   Printer.CurrentX = 10.5
   Printer.Print "CARNET";
   Printer.CurrentX = 14
   Printer.Print "P    E    R    I    O    D    O"
   Printer.CurrentX = 14
   Printer.Print "I      II     III    IV       FINAL"
   Printer.Print ""
   For I = 1 To (MATXPER.Rows - 1)
      Printer.CurrentX = 0.5
      Printer.Print I;
      Printer.CurrentX = 1.3
      Printer.Print RTrim(MATXPER.TextMatrix(I, 1));
      Printer.CurrentX = 10.5
      Printer.Print RTrim(MATXPER.TextMatrix(I, 2));
      Printer.CurrentX = 14
      Printer.Print RTrim(MATXPER.TextMatrix(I, 3));
      Printer.CurrentX = 14.7
      Printer.Print RTrim(MATXPER.TextMatrix(I, 4));
      Printer.CurrentX = 15.4
      Printer.Print RTrim(MATXPER.TextMatrix(I, 5));
      Printer.CurrentX = 16.1
      Printer.Print RTrim(MATXPER.TextMatrix(I, 6));
      Printer.CurrentX = 17.5
      Printer.Print RTrim(MATXPER.TextMatrix(I, 7))
   Next I
   Printer.EndDoc
   Printer.Font.Size = 8
   Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta del boletín académico por periodos, de acuerdo con el grupo y área seleccionado."
End Sub

Private Sub Form_Load()
MATXPER.Row = 0
MATXPER.Col = 0
MATXPER.ColWidth(0) = 400
MATXPER.Text = "CD"
MATXPER.Col = 1
MATXPER.ColWidth(1) = 4150
MATXPER.Text = "APELLIDOS Y NOMBRES"
MATXPER.Col = 2
MATXPER.ColWidth(2) = 1000
MATXPER.Text = "CARNET"
MATXPER.Col = 3
MATXPER.ColWidth(3) = 300
MATXPER.CellForeColor = RGB(255, 255, 255)
MATXPER.CellBackColor = RGB(0, 0, 150)
MATXPER.Text = "  I"
MATXPER.Col = 4
MATXPER.ColWidth(4) = 300
MATXPER.CellForeColor = RGB(255, 255, 255)
MATXPER.CellBackColor = RGB(0, 0, 150)
MATXPER.Text = " II"
MATXPER.Col = 5
MATXPER.ColWidth(5) = 300
MATXPER.CellForeColor = RGB(255, 255, 255)
MATXPER.CellBackColor = RGB(0, 0, 150)
MATXPER.Text = " III"
MATXPER.Col = 6
MATXPER.ColWidth(6) = 300
MATXPER.CellForeColor = RGB(255, 255, 255)
MATXPER.CellBackColor = RGB(0, 0, 150)
MATXPER.Text = " IV"
MATXPER.Col = 7
MATXPER.ColWidth(7) = 550
MATXPER.CellForeColor = RGB(255, 255, 255)
MATXPER.CellBackColor = RGB(0, 0, 150)
MATXPER.Text = "FINAL"
If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") Then
    Command1.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo2.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    cona = 0
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, mate
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    For I = 1 To cona - 1
        Get #NAR, I, mate
        If RTrim(mate.nom) <> "" Then
            Combo3.AddItem RTrim(mate.nom)
        End If
    Next I
    Close #NAR
    Combo2.Text = Combo2.List(0)
    Combo3.Text = Combo3.List(0)
Else
    Command1.Enabled = False
End If
Command2.Enabled = False
End Sub
