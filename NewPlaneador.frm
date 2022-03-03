VERSION 5.00
Begin VB.Form NewPlaneador 
   Caption         =   "Agregar ítem - Planeación semanal"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Fecha (dd/mm/aaaa)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   11415
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   4
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   2
         Top             =   360
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   480
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Día:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   315
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   11415
      Begin VB.Frame Frame5 
         Caption         =   "Logros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   5880
         TabIndex        =   7
         Top             =   3480
         Width           =   5295
         Begin VB.ListBox List_Logros 
            Height          =   2595
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   11
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Competencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   240
         TabIndex        =   6
         Top             =   3480
         Width           =   5295
         Begin VB.ListBox List_Competencias 
            Height          =   2595
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contenidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   5880
         TabIndex        =   5
         Top             =   240
         Width           =   5300
         Begin VB.ListBox List_Contenidos 
            Height          =   2595
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   9
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ejes temáticos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   5300
         Begin VB.ListBox List_Ejes 
            Height          =   2595
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   5055
         End
      End
   End
End
Attribute VB_Name = "NewPlaneador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command9_Click()
Dim kk1 As String, kk2 As String, ColOculta As Long

If Trim(Text1.Text) = "" Then
    MsgBox "No ha escrito el día", 16, "ADVERTENCIA"
    Exit Sub
End If
If Trim(Text2.Text) = "" Then
    MsgBox "No ha escrito el mes", 16, "ADVERTENCIA"
    Exit Sub
End If
If Trim(Text3.Text) = "" Then
    MsgBox "No ha escrito el año", 16, "ADVERTENCIA"
    Exit Sub
End If

If Val(Text1.Text) < 1 Or Val(Text1.Text) > 31 Then
    MsgBox "Día inválido", 16, "ADVERTENCIA"
    Exit Sub
End If
If Val(Text2.Text) < 1 Or Val(Text2.Text) > 12 Then
    MsgBox "Mes inválido", 16, "ADVERTENCIA"
    Exit Sub
End If
If Val(Text3.Text) < 2010 Or Val(Text3.Text) > 2100 Then
    MsgBox "Año inválido", 16, "ADVERTENCIA"
    Exit Sub
End If

For I = 1 To planeacion_semanal.MTPlan.Rows - 1
    If planeacion_semanal.MTPlan.TextMatrix(I, 0) = Format(Text1.Text, "0#") & "/" & Format(Text2.Text, "0#") & "/" & Text3.Text Then
        MsgBox "Esta fecha ya existe en la planeación", 16, "ADVERTENCIA"
        Exit Sub
    End If
Next I
If List_Ejes.SelCount = 0 Then
    MsgBox "Debe seleccionar un eje temático", 16, "ADVERTENCIA"
    Exit Sub
End If
If List_Contenidos.SelCount = 0 Then
    MsgBox "Debe seleccionar un contenido", 16, "ADVERTENCIA"
    Exit Sub
End If
If List_Competencias.SelCount = 0 Then
    MsgBox "Debe seleccionar una competencia", 16, "ADVERTENCIA"
    Exit Sub
End If
If List_Logros.SelCount = 0 Then
    MsgBox "Debe seleccionar un logro", 16, "ADVERTENCIA"
    Exit Sub
End If

kk1 = ""
kk2 = ""
h = 0
NAR = FreeFile
Open Ruta & planeacion_semanal.Label4.Caption & que & lw & ".pln" For Random As #NAR Len = Len(semanal_planeacion)
While Not EOF(NAR)
    h = h + 1
    Get #NAR, h, semanal_planeacion
Wend
Close #NAR

For I = 0 To List_Contenidos.ListCount - 1
    If List_Contenidos.Selected(I) = True Then
        kk1 = kk1 & Val(Left(List_Contenidos.List(I), 2)) & ","
    End If
Next I

For I = 0 To List_Logros.ListCount - 1
    If List_Logros.Selected(I) = True Then
        kk2 = kk2 & Val(Left(List_Logros.List(I), 2)) & ","
    End If
Next I
semanal_planeacion.fecha = Format(Text1.Text, "0#") & "/" & Format(Text2.Text, "0#") & "/" & Text3.Text
ColOculta = Val(Text3.Text & Format(Text2.Text, "0#") & Format(Text1.Text, "0#"))
semanal_planeacion.eje = List_Ejes.ListIndex + 1
semanal_planeacion.competencia = List_Competencias.ListIndex + 1
semanal_planeacion.contenidos = Trim(kk1)
semanal_planeacion.logros = Trim(kk2)
Open Ruta & planeacion_semanal.Label4.Caption & que & lw & ".pln" For Random As #NAR Len = Len(semanal_planeacion)
Put #NAR, h, semanal_planeacion
Close #NAR
planeacion_semanal.MTPlan.Rows = planeacion_semanal.MTPlan.Rows + 1
If List_Contenidos.SelCount <= List_Logros.SelCount Then
    planeacion_semanal.MTPlan.RowHeight(planeacion_semanal.MTPlan.Rows - 1) = 240 * (Val(List_Logros.SelCount) + 3)
Else
    planeacion_semanal.MTPlan.RowHeight(planeacion_semanal.MTPlan.Rows - 1) = 240 * (Val(List_Contenidos.SelCount) + 3)
End If
planeacion_semanal.MTPlan.TextMatrix(planeacion_semanal.MTPlan.Rows - 1, 0) = Trim(semanal_planeacion.fecha)
planeacion_semanal.MTPlan.TextMatrix(planeacion_semanal.MTPlan.Rows - 1, 1) = List_Ejes.List(List_Ejes.ListIndex)
For I = 0 To List_Contenidos.ListCount - 1
    If List_Contenidos.Selected(I) = True Then
        planeacion_semanal.MTPlan.TextMatrix(planeacion_semanal.MTPlan.Rows - 1, 2) = planeacion_semanal.MTPlan.TextMatrix(planeacion_semanal.MTPlan.Rows - 1, 2) & Right(List_Contenidos.List(I), Len(List_Contenidos.List(I)) - 5) & vbCrLf
    End If
Next I
planeacion_semanal.MTPlan.TextMatrix(planeacion_semanal.MTPlan.Rows - 1, 3) = List_Competencias.List(List_Competencias.ListIndex)

For I = 0 To List_Logros.ListCount - 1
    If List_Logros.Selected(I) = True Then
        planeacion_semanal.MTPlan.TextMatrix(planeacion_semanal.MTPlan.Rows - 1, 4) = planeacion_semanal.MTPlan.TextMatrix(planeacion_semanal.MTPlan.Rows - 1, 4) & List_Logros.List(I) & vbCrLf
        z = z + 1
    End If
Next I
planeacion_semanal.MTPlan.TextMatrix(planeacion_semanal.MTPlan.Rows - 1, 5) = ColOculta
planeacion_semanal.MTPlan.Col = 5
planeacion_semanal.MTPlan.Sort = 3
Unload Me
End Sub

Private Sub Form_Load()
List_Ejes.Clear
List_Contenidos.Clear
List_Competencias.Clear
List_Logros.Clear

h = 0
NAR = FreeFile
Open Ruta & fl & ser & que & lw & ".eje" For Random As #NAR Len = Len(semanal_ejetematico)
While Not EOF(NAR)
    h = h + 1
    Get #NAR, h, semanal_ejetematico
Wend
Close #NAR

Open Ruta & fl & ser & que & lw & ".eje" For Random As #NAR Len = Len(semanal_ejetematico)
For r = 1 To h - 1
    Get #NAR, r, semanal_ejetematico
    If Trim(semanal_ejetematico.txt_eje) <> "" Then
        List_Ejes.AddItem Trim(semanal_ejetematico.txt_eje)
    End If
Next r
Close #NAR

h = 0
Open Ruta & fl & ser & que & lw & ".cpt" For Random As #NAR Len = Len(semanal_competencias)
While Not EOF(NAR)
    h = h + 1
    Get #NAR, h, semanal_competencias
Wend
Close #NAR

Open Ruta & fl & ser & que & lw & ".cpt" For Random As #NAR Len = Len(semanal_competencias)
For r = 1 To h - 1
    Get #NAR, r, semanal_competencias
    If Trim(semanal_competencias.txt_comp) <> "" Then
        List_Competencias.AddItem Trim(semanal_competencias.txt_comp)
    End If
Next r
Close #NAR


End Sub

Private Sub List_Competencias_Click()
Dim ListCompLog As String
List_Logros.Clear
ListCompLog = ""
h = 0
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    h = h + 1
    Get #NAR, h, logru
Wend
Close #NAR

Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
For r = 1 To h - 1
    Get #NAR, r, logru
    If Trim(logru.indicador) = "L" Then
        ListCompLog = ListCompLog & Trim(logru.observ) & ","
    End If
Next r
Close #NAR
ArrLogros = Split(ListCompLog, ",")

Open Ruta & fl & ser & que & lw & ".cpt" For Random As #NAR Len = Len(semanal_competencias)
Get #NAR, List_Competencias.ListIndex + 1, semanal_competencias
Close #NAR

ArrLogros2 = Split(semanal_competencias.num_logro, ",")

For r = 0 To UBound(ArrLogros2) - 1
    If Val(ArrLogros2(r)) <= UBound(ArrLogros) Then
        List_Logros.AddItem Format(ArrLogros2(r), "00") & " - " & ArrLogros(ArrLogros2(r) - 1)
    End If
Next r

End Sub

Private Sub List_Ejes_Click()
List_Contenidos.Clear
h = 0
Open Ruta & fl & ser & que & lw & ".ctd" For Random As #NAR Len = Len(semanal_contenidos)
While Not EOF(NAR)
    h = h + 1
    Get #NAR, h, semanal_contenidos
Wend
Close #NAR

Open Ruta & fl & ser & que & lw & ".ctd" For Random As #NAR Len = Len(semanal_contenidos)
For r = 1 To h - 1
    Get #NAR, r, semanal_contenidos
    If semanal_contenidos.num_eje = List_Ejes.ListIndex + 1 Then
        List_Contenidos.AddItem Format(r, "00") & " - " & Trim(semanal_contenidos.txt_cont)
    End If
Next r
Close #NAR
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub
