VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Competencias 
   Caption         =   "Competencias"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   11175
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   9840
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   3735
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Competencias.frx":0000
         Left            =   3240
         List            =   "Competencias.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Competencias.frx":00AE
         Left            =   960
         List            =   "Competencias.frx":00BE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Materia:"
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grado:"
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11175
      Begin VB.CommandButton Command7 
         Caption         =   "Ver Logros"
         Height          =   375
         Left            =   8040
         TabIndex        =   15
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9840
         TabIndex        =   14
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   4680
         TabIndex        =   13
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MTComp 
         Height          =   3015
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedRows       =   0
      End
   End
End
Attribute VB_Name = "Competencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If VALI180 = False Then
   Call Command5_Click
End If
Frame2.Caption = ""
MTComp.Rows = 1
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command7.Enabled = False
Y = 0
NAR = FreeFile
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    Y = Y + 1
    Get #NAR, Y, mate
    If RTrim(mate.nom) = Combo3.Text Then
        que = mate.num
    End If
Wend
Close #NAR
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If (RTrim(argra.grado) = RTrim(Combo2.Text) And (argra.num_area = que)) Then
        Close #NAR
        GoTo intel
    End If
Wend
Close #NAR
MsgBox "ESTA AREA NO ESTA CREADA PARA ESTE GRADO O NO LE CORRESPONDE", 16, "OBSERVACIONES"
Exit Sub
intel:


If RTrim(Combo1.Text) = "PRIMERO" Then
lw = 1
End If
If RTrim(Combo1.Text) = "SEGUNDO" Then
lw = 2
End If
If RTrim(Combo1.Text) = "TERCERO" Then
lw = 3
End If
If RTrim(Combo1.Text) = "CUARTO" Then
lw = 4
End If

fl = "1"
ser = Left(Combo2.Text, 3)


CROA = 0
ListaLogros = False
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    CROA = CROA + 1
    Get #NAR, CROA, logru
    If Trim(logru.indicador) = "L" Then
        ListaLogros = True
    End If
Wend
Close #NAR
If ListaLogros = False Then
     MsgBox "NO EXISTEN LOGROS PARA ESTE PERIODO, AGREGUELOS ANTES DE INGRESAR LAS COMPETENCIAS", 16, "ADVERTENCIA"
     Exit Sub
End If

CROA = 0
Open Ruta & fl & ser & que & lw & ".cpt" For Random As #NAR Len = Len(semanal_competencias)
While Not EOF(NAR)
    CROA = CROA + 1
    Get #NAR, CROA, semanal_competencias
Wend
Close #NAR

Open Ruta & fl & ser & que & lw & ".cpt" For Random As #NAR Len = Len(semanal_competencias)
For J = 1 To CROA - 1
    Get #NAR, J, semanal_competencias
    MTComp.Rows = J + 1
    MTComp.TextMatrix(J, 0) = J
    MTComp.TextMatrix(J, 1) = Trim(semanal_competencias.cod_comp)
    MTComp.TextMatrix(J, 2) = Trim(semanal_competencias.txt_comp)
    MTComp.TextMatrix(J, 3) = Trim(semanal_competencias.num_logro)
Next J
Close #NAR

Frame2.Caption = "GRADO: " & Combo2.Text & " - " & "MATERIA: " & Combo3.Text & " (" & que & ")" & " - " & "PERIODO: " & Combo1.Text
' Se verifica si está bloqueado el periodo para no habilitar el botón guardar
'If VeriPeriodo(lw, "periodosL") = False Then
'    Command5.Enabled = False
'    MsgBox "EL PERIODO " & Combo1 & " SOLO ESTA DISPONIBLE PARA CONSULTA", 32, "Grabar logros y observaciones"
'Else
'    Command5.Enabled = True
'End If
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command7.Enabled = True

End Sub

Private Sub Command2_Click()
ValiModifica = False
New_Competencia.Show 1
End Sub

Private Sub Command3_Click()
TTT = InputBox("Escriba el número de la competencia que desea modificar", "modificar competencia")
If TTT = "" Then
    MsgBox "No escribió el número de la competencia", 64, "Modificar"
    Exit Sub
End If
If Val(TTT) > Val(MTComp.Rows - 1) Or (Val(TTT) < 1) Then
    MsgBox "No existe este número de competencia", 32, "Modificar"
    Exit Sub
End If

ValiModifica = True
New_Competencia.Show 1
End Sub

Private Sub Command4_Click()
If Val(MTComp.Rows - 1) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA ELIMINAR", 64
    Exit Sub
End If
If Val(MTComp.Rows - 1) = 1 Then
    MsgBox "No se puede Eliminar la última competencia de la lista", 32, "Eliminar competencia"
    Exit Sub
End If
TTT = InputBox("Escriba el número de la competencia que desea eliminar", "Eliminar competencia")
If TTT = "" Then
    MsgBox "No escribió el número de la competencia", 64, "Eliminar competencia"
    Exit Sub
End If
If Val(TTT) > Val(MTComp.Rows - 1) Or (Val(TTT) < 1) Then
    MsgBox "No existe este número de competencia", 32, "Eliminar"
    Exit Sub
End If
MTComp.RemoveItem Val(TTT)
For I = 1 To Val(MTComp.Rows - 1)
    MTComp.TextMatrix(I, 0) = I
Next I
'Text1.Text = Val(Text1.Text) - 1
'Text1.Text = ""
'Text3.Text = ""
VALI180 = False
End Sub

Private Sub Command5_Click()
If Val(MTComp.Rows - 1) = 0 Then
    MsgBox "NO HAY INFORMACION PARA GUARDAR", 64
    Exit Sub
End If
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
'VERIFICAMOS SI ESTÁ BLOQUEADO EL PERIODO ACADÉMICO
'If VeriPeriodo(lw, "periodosL") = False Then
'    Exit Sub
'End If
MS1 = "Desea guardar estas competencias?"
'If FileLen(Label2.Caption) <> 0 Then
'   MS1 = "DESEA GUARDAR LOS CAMBIOS EFECTUADOS?"
'End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
   Screen.MousePointer = 11
   Kill Ruta & fl & ser & que & lw & ".cpt"
   NAR = FreeFile
   Open Ruta & fl & ser & que & lw & ".cpt" For Random As #NAR Len = Len(semanal_competencias)
   For X = 1 To Val(MTComp.Rows - 1)
        semanal_competencias.cod_comp = MTComp.TextMatrix(X, 1)
        semanal_competencias.txt_comp = MTComp.TextMatrix(X, 2)
        semanal_competencias.num_logro = MTComp.TextMatrix(X, 3)
       Put #NAR, X, semanal_competencias
   Next X
   Close #NAR
   Screen.MousePointer = 0
End If
VALI180 = True
End Sub

Private Sub Command6_Click()
If VALI180 = False Then
   Call Command5_Click
   Unload Me
Else
   Unload Me
End If
End Sub

Private Sub Command7_Click()
SWobserv = False
Ver_Obser.Show
End Sub

Private Sub Form_Load()
MTComp.Row = 0
MTComp.Col = 0
MTComp.ColWidth(0) = 400
MTComp.CellForeColor = RGB(255, 255, 255)
MTComp.CellBackColor = RGB(0, 0, 150)
MTComp.Text = "No."
MTComp.Col = 1
MTComp.ColWidth(1) = 500
MTComp.CellForeColor = RGB(255, 255, 255)
MTComp.CellBackColor = RGB(0, 0, 150)
MTComp.Text = "CÓD"
MTComp.Col = 2
MTComp.ColWidth(2) = 8000
MTComp.CellForeColor = RGB(255, 255, 255)
MTComp.CellBackColor = RGB(0, 0, 150)
MTComp.Text = "COMPETENCIA"
MTComp.Col = 3
MTComp.ColWidth(3) = 1500
MTComp.CellForeColor = RGB(255, 255, 255)
MTComp.CellBackColor = RGB(0, 0, 150)
MTComp.Text = "LOGROS"

If Dir(Ruta & "materia.edu") <> "" Then
    Command5.Enabled = True
    NAR = FreeFile
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    que = 0
    While Not EOF(NAR)
        que = que + 1
        Get #NAR, que, mate
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    For I = 1 To que - 1
        Get #NAR, I, mate
        If RTrim(mate.nom) <> "" Then
            Combo3.AddItem RTrim(mate.nom)
        End If
    Next I
    Close #NAR
    Combo3.Text = Combo3.List(0)
Else
    Command5.Enabled = False
End If
Combo1 = Combo1.List(0)
Combo2 = Combo2.List(0)
'Text1.Text = ""
'Text3.Text = ""
'Text3.MaxLength = 750
'Text3.Enabled = False
'Combo5.Enabled = False
'Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command7.Enabled = False
'Command8.Enabled = False
'Command9.Enabled = False
'Label1.Caption = 0
VALI180 = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VALI180 = False Then
   Call Command5_Click
   Unload Me
Else
   Unload Me
End If
End Sub
