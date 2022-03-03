VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ARBOL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de grupo"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "ARBOL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Directorio"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Muestra el directorio telefónico por grupos"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Copiar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Copia la lista de estudiantes que conforman el grupo"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Imprime el grupo actual"
      Top             =   5880
      Width           =   855
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00800000&
      Height          =   6255
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid MATI9 
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   10398
         _Version        =   393216
         Rows            =   1
         Cols            =   13
         BackColorBkg    =   12632256
      End
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5775
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10186
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "IMGLIST"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList IMGLIST 
      Left            =   1200
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ARBOL.frx":0442
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ARBOL.frx":075C
            Key             =   "jorna"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ARBOL.frx":0A76
            Key             =   "ok"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ARBOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gra(14) As String
Dim jor(4) As String
Dim nody As String
Dim nodz As String

Private Sub Command1_Click()
'Dim alumno As maestroalum
'Dim icur As inforcur
'Dim alugru As grupoalu
'Dim profe As maestropro
'Dim ini As inicio
If Frame1.Caption <> "" Then
    RESP = MsgBox("DESEA IMPRIMIR EL GRUPO " & nodz & "?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR GRUPO")
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
            If RTrim(icur.nom) = nodz Then
                dire = icur.director
                YUS = RTrim(icur.jornada)
            End If
        Wend
        Close #NAR
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, dire, profe
        Close #NAR
        Printer.Print "Director(a): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        Printer.CurrentY = 1
        Printer.CurrentX = 8
        Printer.Print "GRUPO " & nodz
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
        Open Ruta & nodz & ".gru" For Random As #NAR Len = Len(alugru)
        leo = 0
        While Not EOF(NAR)
            leo = leo + 1
            Get #NAR, leo, alugru
        Wend
        Close #NAR
        Open Ruta & nodz & ".gru" For Random As #NAR Len = Len(alugru)
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
                Printer.Print "Director(a): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
                Printer.CurrentY = 1
                Printer.CurrentX = 8
                Printer.Print "GRUPO " & nodz
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
Else
    MsgBox "ELIJA UN GRUPO PARA IMPRIMIR", vbInformation, "IMPRIMIR"
End If
End Sub

Private Sub Command2_Click()
If Frame1.Caption = "" Then
    MsgBox "NO EXISTE INFORMACION PARA COPIAR", 64, "COPIAR"
    Exit Sub
Else
    COP_INFO.Show 1
End If
End Sub

Private Sub Command3_Click()
Unload Me
DIRECT_TEL.Show
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Para consultar un grupo, de dobleclick en el grado y luego elija la jornada a la que pertenece."
End Sub

Private Sub Form_Load()
Dim nodx As Node
'Dim icur As inforcur
MATI9.Row = 0
MATI9.Col = 0
MATI9.ColWidth(0) = 450
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "CÓD"
MATI9.Col = 1
MATI9.ColWidth(1) = 2200
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "APELLIDOS"
MATI9.Col = 2
MATI9.ColWidth(2) = 2200
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "NOMBRES"
MATI9.Col = 3
MATI9.ColWidth(3) = 800
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "CARNET"
MATI9.Col = 4
MATI9.ColWidth(4) = 1100
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "FECH_NACIM"
MATI9.Col = 5
MATI9.ColWidth(5) = 600
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "EDAD"
MATI9.Col = 6
MATI9.ColWidth(6) = 600
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "RH"
MATI9.Col = 7
MATI9.ColWidth(7) = 1400
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "No.DOCUMENTO"
MATI9.Col = 8
MATI9.ColWidth(8) = 1300
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "AÑO_INGRESO"
MATI9.Col = 9
MATI9.ColWidth(9) = 1200
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "EPS"
MATI9.Col = 10
MATI9.ColWidth(10) = 4000
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "DIRECCION"
MATI9.Col = 11
MATI9.ColWidth(11) = 1300
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "TELEFONO"
MATI9.Col = 12
MATI9.ColWidth(12) = 3300
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "ACUDIENTE"
gra(0) = "PREJARDIN"
gra(1) = "JARDIN"
gra(2) = "TRANSICION"
gra(3) = "PRIMERO"
gra(4) = "SEGUNDO"
gra(5) = "TERCERO"
gra(6) = "CUARTO"
gra(7) = "QUINTO"
gra(8) = "SEXTO"
gra(9) = "SEPTIMO"
gra(10) = "OCTAVO"
gra(11) = "NOVENO"
gra(12) = "DECIMO"
gra(13) = "UNDECIMO"
jor(0) = "UNICA"
jor(1) = "MAÑANA"
jor(2) = "TARDE"
jor(3) = "NOCHE"
NAR = FreeFile
If Dir(Ruta & "infcur.edu") <> "" Then
    For I = 0 To 13
        Set nodx = TreeView1.Nodes.Add(, , gra(I), gra(I), "open")
        For J = 0 To 3
            Set nodx = TreeView1.Nodes.Add(gra(I), tvwChild, gra(I) & jor(J), jor(J), "jorna")
            Open Ruta & "infcur.edu" For Input As #NAR
            While Not EOF(NAR)
                Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
                If (RTrim(icur.jornada) = jor(J)) And (RTrim(icur.grado) = gra(I)) Then _
                Set nodx = TreeView1.Nodes.Add(gra(I) & jor(J), tvwChild, , icur.nom, "ok")
            Wend
            Close #NAR
        Next J
    Next I
Else
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
'Dim alugru As grupoalu
'Dim alumno As maestroalum
'Dim icur As inforcur
'Dim profe As maestropro
nody = Node.Image
If nody = "ok" Then
    nodz = Node.Text
    If Dir(Ruta & nodz & ".gru") = "" Then
        MsgBox "GRUPO INCORRECTO", 48
        Exit Sub
    End If
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        If RTrim(icur.nom) = nodz Then
            GoTo ALTU87
        End If
    Wend
ALTU87:
    Close #NAR
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    Get #NAR, (icur.director), profe
    Close #NAR
    Frame1.Caption = "Director(a): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
    LEO2 = 0
    Open Ruta & nodz & ".gru" For Random As #NAR Len = Len(alugru)
    While Not EOF(NAR)
        LEO2 = LEO2 + 1
        Get #NAR, LEO2, alugru
    Wend
    Close #NAR
    Open Ruta & nodz & ".gru" For Random As #NAR Len = Len(alugru)
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    NAR = FreeFile
    Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
    For TN = 1 To (LEO2 - 1)
        Get #(NAR - 2), TN, alugru
        Get #(NAR - 1), (Val(alugru.num_carnet)), alumno
        Get #NAR, (Val(alugru.num_carnet)), AdiCampo
        MATI9.Rows = TN + 1
        MATI9.TextMatrix(TN, 0) = TN
        MATI9.TextMatrix(TN, 1) = RTrim(alumno.apellidos)
        MATI9.TextMatrix(TN, 2) = RTrim(alumno.nombres)
        MATI9.TextMatrix(TN, 3) = alumno.n_carnet
        MATI9.TextMatrix(TN, 4) = RTrim(alumno.f_nacimiento)
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
        MATI9.TextMatrix(TN, 5) = aaaa
        MATI9.TextMatrix(TN, 6) = RTrim(alumno.rh)
        MATI9.TextMatrix(TN, 7) = RTrim(alumno.documento)
        MATI9.TextMatrix(TN, 8) = RTrim(alumno.año_ingre)
        MATI9.TextMatrix(TN, 9) = RTrim(AdiCampo.salud)
        MATI9.TextMatrix(TN, 10) = RTrim(alumno.direccion)
        MATI9.TextMatrix(TN, 11) = RTrim(AdiCampo.Tel_casa)
        MATI9.TextMatrix(TN, 12) = RTrim(alumno.acudiente)
    Next TN
    Close #NAR
    Close #(NAR - 1)
    Close #(NAR - 2)
    ARBOL.Caption = "Consulta de grupo - " & Format(nodz, "<")
    Screen.MousePointer = 0
End If
End Sub
