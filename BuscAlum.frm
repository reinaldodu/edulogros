VERSION 5.00
Begin VB.Form BuscAlum 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton CanBusq 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   320
         Left            =   3960
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton NewBusq 
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   320
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox NomBusq 
         Height          =   320
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox ApeBusq 
         Height          =   320
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Caption         =   "Buscar por:"
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
         Begin VB.OptionButton PorNom 
            Caption         =   "N&ombres"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton PorApe 
            Caption         =   "Ap&ellidos"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Nombres:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Apellidos:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   675
      End
   End
End
Attribute VB_Name = "BuscAlum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CanBusq_Click()
Unload Me
End Sub


Private Sub NewBusq_Click()
Dim BusqOk As Boolean
If PorApe.Value = True Then
    If RTrim(ApeBusq.Text) = "" Then
        MsgBox "No ha escrito los apellidos para la búsqueda", 64, "Advertencia"
        ApeBusq.SetFocus
        Exit Sub
    End If
End If
If PorNom.Value = True Then
    If RTrim(NomBusq.Text) = "" Then
        MsgBox "No ha escrito los nombres para la búsqueda", 64, "Advertencia"
        NomBusq.SetFocus
        Exit Sub
    End If
End If
ApeBusq.Text = RTrim(Format(ApeBusq.Text, ">"))
NomBusq.Text = RTrim(Format(NomBusq.Text, ">"))
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, k
Close #NAR
i = 0
BusqOk = False
Screen.MousePointer = 11
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
While i < (k - 1)
    i = i + 1
    Get #NAR, i, alumno
    If PorApe.Value = True Then
        If (RTrim(alumno.apellidos) = ApeBusq.Text) Then
            BusqOk = True
            GoTo finbusq
        End If
    End If
    If PorNom.Value = True Then
        If (RTrim(alumno.nombres) = NomBusq.Text) Then
            BusqOk = True
            GoTo finbusq
        End If
    End If
Wend
finbusq:
Close #NAR
Screen.MousePointer = 0
If BusqOk = False Then
    MsgBox "No se encontraron registros", 64
Else
    Open Ruta & "informe.edu" For Random As #NAR Len = Len(detalle)
    Get #NAR, i, detalle
    Close #NAR
    Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
    Get #NAR, i, aluper
    Close #NAR
    Modifico_info = False
    info_adicional.Command2.Enabled = True
    info_adicional.Command4.Enabled = True
    info_adicional.Command6.Enabled = True
    info_adicional.informe.Enabled = True
    info_adicional.Caption = "Información adicional - Carnet No." & alumno.n_carnet
    info_adicional.apellido.Caption = RTrim(alumno.apellidos)
    info_adicional.nombre.Caption = RTrim(alumno.nombres)
    info_adicional.grupo.Caption = RTrim(aluper.grupo)
    info_adicional.informe.Text = RTrim(detalle.info)
    If Dir(Ruta & "FOTOALU\" & i & ".jpg") <> "" Then
        info_adicional.Image1.Picture = LoadPicture(Ruta & "FOTOALU\" & i & ".jpg")
    Else
        info_adicional.Image1.Picture = LoadPicture()
    End If
    info_adicional.Text2.Text = ""
    Unload Me
End If
End Sub
