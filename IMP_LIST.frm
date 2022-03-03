VERSION 5.00
Begin VB.Form IMP_LIST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de listados por jornada y/o grado"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "IMP_LIST.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "&Imprimir toda la jornada"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "opciones"
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   4095
         Begin VB.ComboBox Combo4 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "IMP_LIST.frx":0442
            Left            =   1440
            List            =   "IMP_LIST.frx":0476
            TabIndex        =   4
            Text            =   "PREKINDER"
            Top             =   960
            Width           =   1935
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "IMP_LIST.frx":04FF
            Left            =   1440
            List            =   "IMP_LIST.frx":050F
            TabIndex        =   3
            Text            =   "UNICA"
            Top             =   600
            Width           =   1935
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "IMP_LIST.frx":0530
            Left            =   1440
            List            =   "IMP_LIST.frx":0543
            TabIndex        =   2
            Text            =   "PRIMERO"
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "GRADO    :"
            Height          =   195
            Left            =   600
            TabIndex        =   11
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "JORNADA:"
            Height          =   195
            Left            =   600
            TabIndex        =   10
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PERIODO:"
            Height          =   195
            Left            =   600
            TabIndex        =   9
            Top             =   360
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4095
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "IMP_LIST.frx":0571
            Left            =   960
            List            =   "IMP_LIST.frx":058D
            TabIndex        =   1
            Text            =   "LISTAS DE GRUPOS"
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "IMPRIMIR:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "IMP_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim swmx() As Byte

Private Sub inimatrix()
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
que = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    que = que + 1
    Get #NAR, que, mate
Wend
Close #NAR
ReDim swmx(1 To (r - 1), 1 To 14, 1 To (que - 1))
For I = 1 To (r - 1)
    For J = 1 To 14
        For z = 1 To (que - 1)
            swmx(I, J, z) = 0
        Next z
    Next J
Next I
End Sub

Private Sub imp_hojamatri()
If (RTrim(alumno.n_carnet) <> "") And (RTrim(alumno.nombres) <> "") And (RTrim(alumno.apellidos) <> "") Then
    IMPOK = True
    Printer.ScaleMode = 7
    Printer.CurrentY = 2.5
    Printer.CurrentX = 8
    Printer.Font.Size = 14
    Printer.Print "HOJA DE MATRICULA"
    Printer.Font.Size = 10
    Printer.CurrentY = 3.5
    Printer.CurrentX = 16.5
    Printer.Print "MATRICULA No." & alumno.n_matricula
    Printer.CurrentX = 16.5
    Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
    Printer.CurrentY = 5
    Printer.CurrentX = 2
    Printer.Print "APELLIDOS: " & alumno.apellidos;
    Printer.CurrentX = 12
    Printer.Print "NOMBRES: " & alumno.nombres
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "FECHA DE NACIMIENTO: " & alumno.f_nacimiento;
    Printer.CurrentX = 12
    Printer.Print "DOCUMENTO ID.: " & alumno.documento
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "R - H: " & alumno.rh;
    Printer.CurrentX = 12
    Printer.Print "SEXO: " & alumno.sexo
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "DIRECCION: " & alumno.direccion;
    Printer.CurrentX = 12
    Printer.Print "TELEFONO: " & alumno.tel_acu
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "No. CARNET: " & alumno.n_carnet;
    Printer.CurrentX = 12
    Printer.Print "JORNADA: " & alumno.jornada
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "AÑO DE INGRESO: " & alumno.año_ingre;
    Printer.CurrentX = 12
    Printer.Print "GRADO: " & alumno.grado
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "ACUDIENTE: " & alumno.acudiente
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2
    Printer.Print "OBSERVACIONES:";
    Printer.Line (5.2, Printer.CurrentY + 0.4)-(20, Printer.CurrentY + 0.4)
    Printer.Print ""
    Printer.Print ""
    Printer.Line (2, Printer.CurrentY)-(20, Printer.CurrentY)
    Printer.Print ""
    Printer.Print ""
    Printer.Line (2, Printer.CurrentY)-(20, Printer.CurrentY)
    Printer.CurrentY = 24
    Printer.Line (2, Printer.CurrentY)-(7, Printer.CurrentY)
    Printer.Line (8.2, Printer.CurrentY)-(13.2, Printer.CurrentY)
    Printer.Line (14.4, Printer.CurrentY)-(19.4, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.CurrentX = 2
    Printer.Print "FIRMA DEL ESTUDIANTE";
    Printer.CurrentX = 8.2
    Printer.Print "FIRMA DEL ACUDIENTE";
    Printer.CurrentX = 14.4
    Printer.Print "FIRMA AUTORIZADA Y SELLO"
    Printer.NewPage
End If
End Sub

Private Sub imp_libronotas()
If (RTrim(alumno.n_carnet) <> "") And (RTrim(alumno.nombres) <> "") And (RTrim(alumno.apellidos) <> "") Then
    IMPOK = True
    Printer.ScaleMode = 7
    Printer.Font.Size = 14
    Printer.CurrentY = 3
    Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
    Printer.FontBold = True
    Printer.Print ini.nombre
    Printer.Font.Size = 11
    Printer.CurrentX = 10.2 - (Len(ini.Rector) / 5.2) / 2
    Printer.Print ini.Rector
    Printer.Print ""
    Printer.Print ""
    Printer.Font.Size = 14
    Printer.CurrentX = 10.2 - ((Len("INFORME FINAL DE CALIFICACIONES") / 3.3) / 2)
    Printer.Print "INFORME FINAL DE CALIFICACIONES"
    Printer.FontBold = False
    Printer.Print ""
    Printer.Print ""
    Printer.Font.Size = 12
    Printer.CurrentX = 3
    Printer.Print "AÑO LECTIVO: " & Year(Date);
    Printer.CurrentX = 16
    Printer.Print "CURSO: " & Right(aluper.grupo, Len(aluper.grupo) - 1)
    Printer.CurrentX = 3
    Printer.Print "ESTUDIANTE: " & RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
    Printer.CurrentX = 16
    Printer.Print "No.carnet: " & alumno.n_carnet
    Printer.Print ""
    Printer.Print ""
    Printer.FontBold = True
    Printer.CurrentX = 3
    Printer.Print "A R E A S";
    Printer.CurrentX = 16
    Printer.Print "VALORACION FINAL"
    Printer.FontBold = False
    Printer.Print ""
    cona = 0
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = RTrim(aluper.grupo) Then
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            NAR = NAR - 1
            Printer.CurrentX = 3
            Printer.Print RTrim(mate.nom);
            If Dir(Ruta & RTrim(aluper.grupo) & argra.num_area & "5.obs") <> "" Then
                    NAR = FreeFile
                    z = 0
                    Open Ruta & RTrim(aluper.grupo) & argra.num_area & "5.obs" For Random As #NAR Len = Len(notas)
                    While Not EOF(NAR)
                        z = z + 1
                        Get #NAR, z, notas
                        If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                            Printer.CurrentX = 17
                            Printer.Print notas.FA
                            GoTo camila2
                        End If
                    Wend
camila2:
                    Close #NAR
                    NAR = NAR - 1
            Else
                Printer.Print ""
            End If
        End If
    Wend
    Close #NAR
    NAR = NAR - 1
    Printer.CurrentY = 20
    Printer.CurrentX = 3
    Printer.Print "OBSERVACIONES:"
    Printer.Line (7, Printer.CurrentY)-(20, Printer.CurrentY)
    Printer.Line (3, Printer.CurrentY + 0.5)-(20, Printer.CurrentY + 0.5)
    Printer.Line (3, Printer.CurrentY + 0.5)-(20, Printer.CurrentY + 0.5)
    Printer.Line (3, 25)-(8.5, 25)
    Printer.Line (13, 25)-(18.5, 25)
    Printer.CurrentY = 25.2
    Printer.CurrentX = 4.3
    Printer.Print "Firma del Directora.";
    Printer.CurrentX = 13.9
    Printer.Print "Firma del Secretario(a)."
    Printer.NewPage
End If
End Sub

Private Sub imp_directel()
Printer.CurrentY = 0
Printer.CurrentX = 1
Printer.Font.Size = 10
Printer.Print "DIRECTORIO TELEFONICO - GRUPO " & RTrim(icur.nom)
Printer.CurrentX = 1
Printer.Print ini.nombre;
NAR = FreeFile
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, icur.director, profe
Close #NAR
Printer.CurrentX = 12.5
Printer.Print "Directora: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
Printer.Font.Size = 8
Printer.Print ""
Printer.CurrentX = 1
Printer.Print "CARNET";
Printer.CurrentX = 2.5
Printer.Print "APELLIDOS Y NOMBRES";
Printer.CurrentX = 7.5
Printer.Print "DIRECCION";
Printer.CurrentX = 12.4
Printer.Print "TELEFONO";
Printer.CurrentX = 14.1
Printer.Print "MADRE";
Printer.CurrentX = 18.8
Printer.Print "TELEFONO";
Printer.CurrentX = 20.5
Printer.Print "PADRE";
Printer.CurrentX = 25.2
Printer.Print "TELEFONO"
Printer.Print ""
Printer.Font.Size = 7
If Dir(Ruta & RTrim(icur.nom) & ".gru") <> "" Then
    ret = 0
    Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
    While Not EOF(NAR)
        ret = ret + 1
        Get #NAR, ret, alugru
    Wend
    Close #NAR
    Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
    For J = 1 To (ret - 1)
        Get #NAR, J, alugru
        NAR = FreeFile
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        NAR = NAR - 1
        Printer.CurrentX = 1
        Printer.Print alumno.n_carnet;
        Printer.CurrentX = 2.5
        Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
        Printer.CurrentX = 7.5
        Printer.Print RTrim(alumno.direccion);
        Printer.CurrentX = 12.4
        Printer.Print RTrim(alumno.tel_acu);
        Printer.CurrentX = 14.1
        Printer.Print RTrim(alumno.madre);
        Printer.CurrentX = 18.8
        Printer.Print RTrim(alumno.tel_ma);
        Printer.CurrentX = 20.5
        Printer.Print RTrim(alumno.padre);
        Printer.CurrentX = 25.2
        Printer.Print RTrim(alumno.tel_pa)
    Next J
    Close #NAR
    NAR = NAR - 1
    Printer.Font.Size = 8
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print "TOTAL ESTUDIANTES..." & (ret - 1)
End If
Printer.NewPage
End Sub

Private Sub imp_bolgrab()
If Dir(Ruta & RTrim(argra.nom_grup) & (argra.num_area) & lw & ".obs") <> "" Then
    If FileLen(Ruta & RTrim(argra.nom_grup) & (argra.num_area) & lw & ".obs") <> 0 Then
        IMPOK = True
        Printer.ScaleMode = 7
        Printer.Font.Size = 10
        Printer.CurrentY = 1
        Printer.CurrentX = 6.5
        Printer.Print "CONTROL DE LOGROS PERIODO " & Combo2.Text
        Printer.Print ""
        Printer.CurrentX = 0.5
        Printer.Print ini.nombre;
        Printer.CurrentX = 16.5
        Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
        NAR = FreeFile
        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
        Get #NAR, argra.num_area, mate
        Close #NAR
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, argra.num_pro, profe
        Close #NAR
        Printer.CurrentX = 0.5
        Printer.Print "GRUPO: " & RTrim(argra.nom_grup) & " - AREA: " & RTrim(mate.nom) & " - PROFESOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        Printer.Print ""
        Printer.CurrentX = 0.5
        Printer.Print "CD";
        Printer.CurrentX = 1.3
        Printer.Print "APELLIDOS Y NOMBRES";
        Printer.CurrentX = 10.5
        Printer.Print "LOGROS Y/O DIFICULTADES"
        Printer.CurrentX = 10.5
        Printer.Print "LG1 LG2 LG3 LG4 LG5 LG6 LG7 LG8 LG9 LG10";
        Printer.CurrentX = 18.4
        Printer.Print "JV";
        Printer.CurrentX = 19.4
        Printer.Print "FA"
        Y = 0
        Open Ruta & RTrim(argra.nom_grup) & (argra.num_area) & lw & ".obs" For Random As #NAR Len = Len(notas)
        While Not EOF(NAR)
            Y = Y + 1
            Get #NAR, Y, notas
        Wend
        Close #NAR
        Open Ruta & RTrim(argra.nom_grup) & (argra.num_area) & lw & ".obs" For Random As #NAR Len = Len(notas)
        For I = 1 To (Y - 1)
            Get #NAR, I, notas
            NAR = FreeFile
            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
            Get #NAR, (Val(notas.num_carnet)), alumno
            Close #NAR
            NAR = NAR - 1
            Printer.CurrentX = 0.5
            Printer.Print I;
            Printer.CurrentX = 1.3
            Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
            CX = 10.6
            For J = 1 To 10
                If notas.area(J) <> 0 Then
                    Printer.CurrentX = CX
                    Printer.Print notas.area(J);
                End If
                If J = 9 Then
                    CX = CX + 0.8
                Else
                    CX = CX + 0.7
                End If
            Next J
            Printer.CurrentX = 18.5
            Printer.Print notas.FA;
            Printer.CurrentX = 19.5
            Printer.Print notas.FA
        Next I
        Close #NAR
        NAR = NAR - 1
        Printer.NewPage
    End If
End If
End Sub

Private Sub imp_obser()
If Dir(Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> "" Then
    If FileLen(Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> 0 Then
        If swmx(argra.num_pro, s, argra.num_area) <> 1 Then
            swmx(argra.num_pro, s, argra.num_area) = 1
            IMPOK = True
            PAG = 1
            Printer.ScaleMode = 7
            Printer.Font.Size = 9
            Printer.CurrentY = 1
            Printer.CurrentX = 8
            Printer.Print "REPORTE DE OBSERVACIONES"
            Printer.CurrentX = 19
            Printer.Print "Pág." & PAG
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            Printer.CurrentX = 1
            Printer.Print "JORNADA: " & Combo3.Text & " - GRADO: " & RTrim(argra.grado) & " - AREA: " & RTrim(mate.nom) & " (" & argra.num_area & ") - PERIODO: " & Combo2.Text
            Printer.CurrentX = 1
            Printer.Print ini.nombre;
            Printer.CurrentX = 17
            Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
            Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
            Get #NAR, argra.num_pro, profe
            Close #NAR
            Printer.CurrentX = 1
            Printer.Print "PROFESOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
            Printer.Print ""
            Printer.CurrentX = 1
            Printer.Print "CD";
            Printer.CurrentX = 2
            Printer.Print "IND";
            Printer.CurrentX = 3
            Printer.Print "OBSERVACION"
            Printer.Print ""
            L = 0
            Open Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr" For Random As #NAR Len = Len(logru)
            J = 0
            While Not EOF(NAR)
                J = J + 1
                Get #NAR, J, logru
            Wend
            For I = 1 To (J - 1)
                Get #NAR, I, logru
                Printer.CurrentX = 1
                Printer.Print I;
                Printer.CurrentX = 2
                Printer.Print logru.indicador;
                X = 100
                L1 = Left(logru.observ, X)
                While Right(L1, 1) <> " "
                X = X - 1
                L1 = Left(L1, X)
                Wend
                Printer.CurrentX = 3
                Printer.Print L1
                L = L + 1
                If (L Mod 60) = 0 Then
                    PAG = PAG + 1
                    Printer.NewPage
                    Printer.CurrentY = 1
                    Printer.CurrentX = 8
                    Printer.Print "REPORTE DE OBSERVACIONES"
                    Printer.CurrentX = 19
                    Printer.Print "Pág." & PAG
                    Printer.CurrentX = 1
                    Printer.Print "JORNADA: " & Combo3.Text & " - GRADO: " & RTrim(argra.grado) & " - AREA: " & RTrim(mate.nom) & " (" & argra.num_area & ") - PERIODO: " & Combo2.Text
                    Printer.CurrentX = 1
                    Printer.Print ini.nombre;
                    Printer.CurrentX = 17
                    Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
                    Printer.CurrentX = 1
                    Printer.Print "PROFESOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
                    Printer.Print ""
                    Printer.CurrentX = 1
                    Printer.Print "CD";
                    Printer.CurrentX = 2
                    Printer.Print "IND";
                    Printer.CurrentX = 3
                    Printer.Print "OBSERVACION"
                    Printer.Print ""
                End If
                Y = Len(L1)
                Y = 200 - Y
                L2 = Right(logru.observ, Y)
                If RTrim(L2) <> "" Then
                    Printer.CurrentX = 3
                    Printer.Print L2
                    L = L + 1
                    If (L Mod 60) = 0 Then
                        PAG = PAG + 1
                        Printer.NewPage
                        Printer.CurrentY = 1
                        Printer.CurrentX = 8
                        Printer.Print "REPORTE DE OBSERVACIONES"
                        Printer.CurrentX = 19
                        Printer.Print "Pág." & PAG
                        Printer.CurrentX = 1
                        Printer.Print "JORNADA: " & Combo3.Text & " - GRADO: " & RTrim(argra.grado) & " - AREA: " & RTrim(mate.nom) & " (" & argra.num_area & ") - PERIODO: " & Combo2.Text
                        Printer.CurrentX = 1
                        Printer.Print ini.nombre;
                        Printer.CurrentX = 17
                        Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
                        Printer.CurrentX = 1
                        Printer.Print "PROFESOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
                        Printer.Print ""
                        Printer.CurrentX = 1
                        Printer.Print "CD";
                        Printer.CurrentX = 2
                        Printer.Print "IND";
                        Printer.CurrentX = 3
                        Printer.Print "OBSERVACION"
                        Printer.Print ""
                    End If
                End If
                Printer.Print ""
                L = L + 1
                If (L Mod 60) = 0 Then
                    PAG = PAG + 1
                    Printer.NewPage
                    Printer.CurrentY = 1
                    Printer.CurrentX = 8
                    Printer.Print "REPORTE DE OBSERVACIONES"
                    Printer.CurrentX = 19
                    Printer.Print "Pág." & PAG
                    Printer.CurrentX = 1
                    Printer.Print "JORNADA: " & Combo3.Text & " - GRADO: " & RTrim(argra.grado) & " - AREA: " & RTrim(mate.nom) & " (" & argra.num_area & ") - PERIODO: " & Combo2.Text
                    Printer.CurrentX = 1
                    Printer.Print ini.nombre;
                    Printer.CurrentX = 17
                    Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
                    Printer.CurrentX = 1
                    Printer.Print "PROFESOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
                    Printer.Print ""
                    Printer.CurrentX = 1
                    Printer.Print "CD";
                    Printer.CurrentX = 2
                    Printer.Print "IND";
                    Printer.CurrentX = 3
                    Printer.Print "OBSERVACION"
                    Printer.Print ""
                End If
            Next I
            Close #NAR
            NAR = NAR - 1
            Printer.NewPage
        End If
    End If
End If
End Sub

Private Sub imp_pendxgrupo()
Printer.CurrentY = 0.5
Printer.CurrentX = 12
Printer.Print "LISTADO DE LOGROS PENDIENTES - GENERALES " & "(" & RTrim(icur.nom) & ")"
Printer.CurrentY = 2
Printer.CurrentX = 1
Printer.Print ini.nombre;
NAR = FreeFile
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, icur.director, profe
Close #NAR
Printer.CurrentX = 12
Printer.Print "Directora: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
Printer.CurrentX = 24
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy");
Printer.CurrentX = 28
Printer.Print "PERIODO: " & Combo2.Text
Printer.CurrentY = 3
Printer.CurrentX = 1
Printer.Print "CD";
Printer.CurrentX = 1.5
Printer.Print "APELLIDOS Y NOMBRES";
cona = 0
CX = 8
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If RTrim(argra.nom_grup) = RTrim(icur.nom) Then
        Printer.CurrentX = CX
        NAR = FreeFile
        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
        Get #NAR, argra.num_area, mate
        Close #NAR
        NAR = NAR - 1
        If argra.num_area < 10 Then
            Printer.Print Left(mate.nom, 3) & " (" & mate.num & ")";
        Else
            Printer.Print Left(mate.nom, 3) & "(" & mate.num & ")";
        End If
        CX = CX + 1.15
    End If
Wend
Close #NAR
Printer.CurrentX = CX
Printer.Print "TTL"
Printer.Print ""
ret = 0
Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
For J = 1 To (ret - 1)
    Get #NAR, J, alugru
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, (Val(alugru.num_carnet)), alumno
    Close #NAR
    NAR = NAR - 1
    Printer.CurrentX = 1
    Printer.Print J;
    Printer.CurrentX = 1.5
    Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
    EXISALU = False
    z = 0
    h = 0
    CX = 8
    cona = 0
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = RTrim(icur.nom) Then
            k = 0
            For ww = 1 To lw
                CP = 0
                If Dir(Ruta & RTrim(argra.nom_grup) & (argra.num_area) & ww & ".obp") <> "" Then
                    r = 0
                    NAR = FreeFile
                    Open Ruta & RTrim(argra.nom_grup) & (argra.num_area) & ww & ".obp" For Random As #NAR Len = Len(notas)
                    While Not EOF(NAR)
                        r = r + 1
                        Get #NAR, r, notas
                        If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                            EXISALU = True
                            NAR = FreeFile
                            Open Ruta & fl & Left(argra.grado, 3) & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
                            For I = 1 To 10
                                If notas.area(I) <> 0 Then
                                    Get #NAR, notas.area(I), logru
                                    If (logru.indicador = "D") Or (logru.indicador = "N") Or (logru.indicador = "I") Then
                                        CP = CP + 1
                                    End If
                                End If
                            Next I
                            Close #NAR
                            NAR = NAR - 1
                            k = k + CP
                            GoTo LPA2
                        End If
                    Wend
LPA2:
                    Close #NAR
                    NAR = NAR - 1
                End If
                If ww = lw Then
                    If k <> 0 Then
                        Printer.CurrentX = CX
                        Printer.Print k & "(" & CP & ")";
                        h = h + CP
                        CX = CX + 1.15
                    Else
                        CX = CX + 1.15
                    End If
                End If
            Next ww
            z = z + k
        End If
    Wend
    Close #NAR
    NAR = NAR - 1
    If EXISALU = True Then
        Printer.CurrentX = CX
        Printer.Print z & "(" & h & ")"
    Else
        Printer.Print ""
    End If
Next J
Close #NAR
NAR = NAR - 1
Printer.NewPage
End Sub

Private Sub imp_pendxarea()
NAR = FreeFile
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If RTrim(icur.nom) = RTrim(argra.nom_grup) Then
        If Dir(Ruta & RTrim(icur.nom) & argra.num_area & lw & ".obp") <> "" Then
            PAG = 1
            Printer.CurrentY = 1
            Printer.CurrentX = 6
            Printer.Print "LISTADO DE LOGROS PENDIENTES PERIODO " & Combo2.Text
            Printer.CurrentY = 1.5
            Printer.CurrentX = 19
            Printer.Print "Pág." & PAG
            Printer.Print ""
            Printer.CurrentX = 0.5
            Printer.Print ini.nombre;
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            NAR = NAR - 1
            Printer.CurrentX = 11
            Printer.Print "AREA: " & RTrim(mate.nom)
            NAR = FreeFile
            Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
            Get #NAR, argra.num_pro, profe
            Close #NAR
            NAR = NAR - 1
            Printer.CurrentX = 0.5
            Printer.Print "PROFESOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
            Printer.CurrentX = 11
            Printer.Print "GRUPO: " & RTrim(icur.nom);
            Printer.CurrentX = 17
            Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
            Printer.Print ""
            Printer.CurrentX = 0.5
            Printer.Print "APELLIDOS Y NOMBRES";
            Printer.CurrentX = 7
            Printer.Print "No.";
            Printer.CurrentX = 7.5
            Printer.Print "LOGROS PENDIENTES"
            Printer.Print ""
            k = 0
            NAR = FreeFile
            Open Ruta & RTrim(icur.nom) & argra.num_area & lw & ".obp" For Random As #NAR Len = Len(notas)
            While Not EOF(NAR)
                k = k + 1
                Get #NAR, k, notas
            Wend
            Close #NAR
            z = 1
            Open Ruta & RTrim(icur.nom) & argra.num_area & lw & ".obp" For Random As #NAR Len = Len(notas)
            For I = 1 To (k - 1)
                Get #NAR, I, notas
                NAR = FreeFile
                Open Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr" For Random As #NAR Len = Len(logru)
                J = 0
                For r = 1 To 10
                    If notas.area(r) <> 0 Then
                        Get #NAR, notas.area(r), logru
                        If (z Mod 66) = 0 Then
                            Printer.NewPage
                            PAG = PAG + 1
                            Printer.CurrentY = 1
                            Printer.CurrentX = 6
                            Printer.Print "LISTADO DE LOGROS PENDIENTES PERIODO " & Combo2.Text
                            Printer.CurrentY = 1.5
                            Printer.CurrentX = 19
                            Printer.Print "Pág." & PAG
                            Printer.Print ""
                            Printer.CurrentX = 0.5
                            Printer.Print ini.nombre;
                            NAR = FreeFile
                            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                            Get #NAR, argra.num_area, mate
                            Close #NAR
                            NAR = NAR - 1
                            Printer.CurrentX = 11
                            Printer.Print "AREA: " & RTrim(mate.nom)
                            NAR = FreeFile
                            Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
                            Get #NAR, argra.num_pro, profe
                            Close #NAR
                            NAR = NAR - 1
                            Printer.CurrentX = 0.5
                            Printer.Print "PROFESOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
                            Printer.CurrentX = 11
                            Printer.Print "GRUPO: " & RTrim(icur.nom);
                            Printer.CurrentX = 17
                            Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
                            Printer.Print ""
                            Printer.CurrentX = 0.5
                            Printer.Print "APELLIDOS Y NOMBRES";
                            Printer.CurrentX = 7
                            Printer.Print "No.";
                            Printer.CurrentX = 7.5
                            Printer.Print "LOGROS PENDIENTES"
                            Printer.Print ""
                        End If
                        If (logru.indicador = "D") Or (logru.indicador = "N") Or (logru.indicador = "I") Then
                            If J = 0 Then
                                NAR = FreeFile
                                Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
                                Get #NAR, (Val(notas.num_carnet)), alumno
                                Close #NAR
                                NAR = NAR - 1
                                Printer.CurrentX = 0.5
                                Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres);
                            End If
                            J = J + 1
                            Printer.CurrentX = 7
                            Printer.Print notas.area(r);
                            Printer.CurrentX = 7.5
                            Printer.Print RTrim(logru.observ)
                            z = z + 1
                        End If
                    End If
                Next r
                Close #NAR
                NAR = NAR - 1
                If J <> 0 Then
                    Printer.Print ""
                    z = z + 1
                End If
            Next I
            Close #NAR
            NAR = NAR - 1
            Printer.NewPage
            IMPOK = True
        End If
    End If
Wend
Close #NAR
NAR = NAR - 1
End Sub

Private Sub imp_listgrup()
Printer.CurrentY = 1
Printer.CurrentX = 7.5
Printer.Font.Size = 12
Printer.Print "LISTA GRUPO " & RTrim(icur.nom)
Printer.Font.Size = 10
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print ini.nombre;
Printer.CurrentX = 11
Printer.Print "JORNADA: " & RTrim(icur.jornada);
Printer.CurrentX = 16.5
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
NAR = FreeFile
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, icur.director, profe
Close #NAR
NAR = NAR - 1
Printer.CurrentX = 0.5
Printer.Print "Directora: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
Printer.CurrentX = 11
Printer.Print "GRADO: " & RTrim(icur.grado)
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print "CD";
Printer.CurrentX = 1.3
Printer.Print "APELLIDOS Y NOMBRES"
If Dir(Ruta & RTrim(icur.nom) & ".gru") <> "" Then
    ret = 0
    NAR = FreeFile
    Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
    While Not EOF(NAR)
        ret = ret + 1
        Get #NAR, ret, alugru
    Wend
    Close #NAR
    Open Ruta & RTrim(icur.nom) & ".gru" For Random As #NAR Len = Len(alugru)
    For J = 1 To (ret - 1)
        Get #NAR, J, alugru
        Printer.CurrentX = 0.5
        Printer.Print J;
        NAR = FreeFile
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        NAR = NAR - 1
        Printer.CurrentX = 1.3
        Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    Next J
    Close #NAR
    NAR = NAR - 1
End If
Printer.NewPage
End Sub

Private Sub imp_traba()
Printer.CurrentY = 1
Printer.CurrentX = 5.5
Printer.Font.Size = 12
Printer.Print "CONTROL DE LOGROS  PERIODO: " & Combo2.Text
Printer.Font.Size = 10
Printer.Print ""
Printer.CurrentX = 0.5
Printer.Print ini.nombre;
Printer.CurrentX = 16.5
Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
NAR = FreeFile
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, argra.num_pro, profe
Close #NAR
NAR = NAR - 1
Printer.CurrentX = 0.5
Printer.Print "PROFESOR(A): " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
Get #NAR, argra.num_area, mate
Close #NAR
NAR = NAR - 1
Printer.CurrentX = 11
Printer.Print "AREA: " & RTrim(mate.nom);
Printer.CurrentX = 19
Printer.Print "IH: " & argra.ih
Printer.CurrentX = 0.5
Printer.Print "JORNADA: " & Combo3.Text;
Printer.CurrentX = 11
Printer.Print "GRUPO: " & RTrim(argra.nom_grup)
Printer.Print ""
If Combo1.ListIndex = 1 Then
    Printer.CurrentX = 0.5
    Printer.Print "CD";
    Printer.CurrentX = 1.3
    Printer.Print "APELLIDOS Y NOMBRES"
Else
    Printer.CurrentX = 0.5
    Printer.Print "CD";
    Printer.CurrentX = 1.3
    Printer.Print "APELLIDOS Y NOMBRES";
    Printer.CurrentX = 10.5
    Printer.Print "LOGROS Y/O DIFICULTADES";
    Printer.CurrentX = 18.4
    Printer.Print "JV";
    Printer.CurrentX = 19.4
    Printer.Print "FA"
    Printer.CurrentX = 10.5
    Printer.Print "LG1 LG2 LG3 LG4 LG5 LG6 LG7 LG8 LG9 LG10"
End If
If Dir(Ruta & RTrim(argra.nom_grup) & ".gru") <> "" Then
    ret = 0
    NAR = FreeFile
    Open Ruta & RTrim(argra.nom_grup) & ".gru" For Random As #NAR Len = Len(alugru)
    While Not EOF(NAR)
        ret = ret + 1
        Get #NAR, ret, alugru
    Wend
    Close #NAR
    Open Ruta & RTrim(argra.nom_grup) & ".gru" For Random As #NAR Len = Len(alugru)
    For J = 1 To (ret - 1)
        Get #NAR, J, alugru
        Printer.CurrentX = 0.5
        Printer.Print J;
        NAR = FreeFile
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Get #NAR, (Val(alugru.num_carnet)), alumno
        Close #NAR
        NAR = NAR - 1
        Printer.CurrentX = 1.3
        Printer.Print RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
    Next J
    Close #NAR
    NAR = NAR - 1
End If
Printer.NewPage
End Sub

Private Sub list_hojamatri()
If (Dir(Ruta & "prinalu.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    J = 0
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    While Not EOF(NAR)
        J = J + 1
        Get #NAR, J, alumno
        If Check1.Value = 1 Then
            If RTrim(alumno.jornada) = Combo3.Text Then
                Call imp_hojamatri
            End If
        Else
            If (RTrim(alumno.jornada) = Combo3.Text) And (RTrim(alumno.grado) = Combo4.Text) Then
                Call imp_hojamatri
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Hojas de matrícula"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub list_libronotas()
If (Dir(Ruta & "prinalu.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    J = 0
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    While Not EOF(NAR)
        J = J + 1
        Get #NAR, J, alumno
        NAR = FreeFile
        Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
        Get #NAR, J, aluper
        Close #NAR
        NAR = NAR - 1
        If Check1.Value = 1 Then
            If RTrim(alumno.jornada) = Combo3.Text Then
                Call imp_libronotas
            End If
        Else
            If (RTrim(alumno.jornada) = Combo3.Text) And (RTrim(alumno.grado) = Combo4.Text) Then
                Call imp_libronotas
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Libro de notas"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub list_directel()
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "prinpro.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    Printer.Orientation = 2
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        If Check1.Value = 1 Then
            If Left((icur.nom), 1) = fl Then
                Call imp_directel
                IMPOK = True
            End If
        Else
            If (Left((icur.nom), 1) = fl) And (RTrim(icur.grado) = Combo4.Text) Then
                Call imp_directel
                IMPOK = True
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Directorio telefónico"
        Screen.MousePointer = 0
        Printer.Orientation = 1
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Orientation = 1
    Printer.Font.Size = 8
End If
End Sub

Private Sub list_grabolet()
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "prinpro.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    cona = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If Check1.Value = 1 Then
            If Left((argra.nom_grup), 1) = fl Then
                Call imp_bolgrab
            End If
        Else
            If (Left((argra.nom_grup), 1) = fl) And (RTrim(argra.grado) = Combo4.Text) Then
                Call imp_bolgrab
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Imprimir grabación de boletines"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub list_obser()
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "prinpro.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    Call inimatrix
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    cona = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        For k = 0 To (Combo4.ListCount - 1)
            If RTrim(argra.grado) = Combo4.List(k) Then
                s = k + 1
            End If
        Next k
        If Check1.Value = 1 Then
            If Left((argra.nom_grup), 1) = fl Then
                Call imp_obser
            End If
        Else
            If (Left((argra.nom_grup), 1) = fl) And (RTrim(argra.grado) = Combo4.Text) Then
                Call imp_obser
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Impresión de observaciones"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub list_pendi()
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "prinpro.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        If Check1.Value = 1 Then
            If Left((icur.nom), 1) = fl Then
                Call imp_pendxarea
            End If
        Else
            If (Left((icur.nom), 1) = fl) And (RTrim(icur.grado) = Combo4.Text) Then
                Call imp_pendxarea
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Listas de logros pendientes"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub list_grupend()
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "prinpro.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    Printer.Orientation = 2
    Printer.PaperSize = 5
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        If Check1.Value = 1 Then
            If Left((icur.nom), 1) = fl Then
                Call imp_pendxgrupo
                IMPOK = True
            End If
        Else
            If (Left((icur.nom), 1) = fl) And (RTrim(icur.grado) = Combo4.Text) Then
                Call imp_pendxgrupo
                IMPOK = True
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Listas de logros pendientes"
        Screen.MousePointer = 0
        Printer.Orientation = 1
        Printer.PaperSize = 1
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Orientation = 1
    Printer.PaperSize = 1
    Printer.Font.Size = 8
End If
End Sub

Private Sub List_grupo()
If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "prinpro.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        If Check1.Value = 1 Then
            If Left((icur.nom), 1) = fl Then
                Call imp_listgrup
                IMPOK = True
            End If
        Else
            If (Left((icur.nom), 1) = fl) And (RTrim(icur.grado) = Combo4.Text) Then
                Call imp_listgrup
                IMPOK = True
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Imprimir listas de grupo"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub

Private Sub list_traba()
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "prinpro.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    IMPOK = False
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    cona = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If Check1.Value = 1 Then
            If Left((argra.nom_grup), 1) = fl Then
                Call imp_traba
                IMPOK = True
            End If
        Else
            If (Left((argra.nom_grup), 1) = fl) And (RTrim(argra.grado) = Combo4.Text) Then
                Call imp_traba
                IMPOK = True
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para imprimir", 48, "Impresión de listas"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.Font.Size = 8
End If
End Sub


Private Sub Combo1_Change()
If Combo1.Text <> Combo1.List(0) Then
    Combo1.Text = Combo1.List(0)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo3.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo4.SetFocus
End If
End Sub

Private Sub Combo4_Change()
If Combo4.Text <> Combo4.List(0) Then
    Combo4.Text = Combo4.List(0)
End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0
If Check1.Value = 1 Then
    If Combo1.ListIndex = 0 Then
        MS1 = "Desea imprimir las listas de grupos de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 1 Then
        MS1 = "Desea imprimir todas las listas de trabajo de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 2 Then
        MS1 = "Desea imprimir todas las listas finales de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 3 Then
        MS1 = "Desea imprimir todas las listas de logros pendientes por área de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 4 Then
        MS1 = "Desea imprimir todas las listas de logros pendientes por grupo de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 5 Then
        MS1 = "Desea imprimir todas las listas de observaciones de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 6 Then
        MS1 = "Desea imprimir todas las listas de grabación de boletines de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 7 Then
        MS1 = "Desea imprimir todos los directorios telefónicos de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 8 Then
        MS1 = "Desea imprimir todas las hojas de matrícula de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 9 Then
        MS1 = "Desea imprimir el libro de notas de la jornada " & Format(Combo3.Text, "<") & "?"
    End If
Else
    If Combo1.ListIndex = 0 Then
        MS1 = "Desea imprimir las listas de grupos del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 1 Then
        MS1 = "Desea imprimir todas las listas de trabajo del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 2 Then
        MS1 = "Desea imprimir todas las listas finales del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 3 Then
        MS1 = "Desea imprimir todas las listas de logros pendientes por área del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 4 Then
        MS1 = "Desea imprimir todas las listas de logros pendientes por grupo del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 5 Then
        MS1 = "Desea imprimir todas las listas de observaciones del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 6 Then
        MS1 = "Desea imprimir todas las listas de grabación de boletines del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 7 Then
        MS1 = "Desea imprimir todos los directorios telefónicos del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 8 Then
        MS1 = "Desea imprimir todas las hojas de matrícula del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
    If Combo1.ListIndex = 9 Then
        MS1 = "Desea imprimir el libro de notas del grado " & Format(Combo4.Text, "<") & ", jornada " & Format(Combo3.Text, "<") & "?"
    End If
End If
If Combo2.Text = "PRIMERO" Then
    lw = 1
End If
If Combo2.Text = "SEGUNDO" Then
    lw = 2
End If
If Combo2.Text = "TERCERO" Then
    lw = 3
End If
If Combo2.Text = "CUARTO" Then
    lw = 4
End If
If Combo2.Text = "FINAL" Then
    lw = 5
End If
If Combo3.Text = "UNICA" Then
    fl = "1"
End If
If Combo3.Text = "MAÑANA" Then
    fl = "2"
End If
If Combo3.Text = "TARDE" Then
    fl = "3"
End If
If Combo3.Text = "NOCHE" Then
    fl = "4"
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "Impresión de listados")
If RESP = vbYes Then
    If Combo1.ListIndex = 0 Then
        Call List_grupo
    End If
    If Combo1.ListIndex = 1 Then
        Call list_traba
    End If
    If Combo1.ListIndex = 2 Then
        Call list_traba
    End If
    If Combo1.ListIndex = 3 Then
        Call list_pendi
    End If
    If Combo1.ListIndex = 4 Then
        Call list_grupend
    End If
    If Combo1.ListIndex = 5 Then
        Call list_obser
    End If
    If Combo1.ListIndex = 6 Then
        Call list_grabolet
    End If
    If Combo1.ListIndex = 7 Then
        Call list_directel
    End If
    If Combo1.ListIndex = 8 Then
        Call list_hojamatri
    End If
    If Combo1.ListIndex = 9 Then
        Call list_libronotas
    End If
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Impresión de listados por jornada y/o grado."
End Sub
