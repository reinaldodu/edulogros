VERSION 5.00
Begin VB.Form HISTORIAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación del historial"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ForeColor       =   &H00400040&
   Icon            =   "HISTORIAL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3135
         Begin VB.CommandButton Command1 
            Caption         =   "CREAR HISTORIAL"
            Height          =   735
            Left            =   600
            Picture         =   "HISTORIAL.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "CREA EL HISTORIAL DEL AÑO SELECCIONADO"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3135
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "HISTORIAL.frx":0884
            Left            =   1320
            List            =   "HISTORIAL.frx":0886
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   2640
            Picture         =   "HISTORIAL.frx":0888
            Top             =   240
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   0
            Picture         =   "HISTORIAL.frx":0CCA
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "AÑO:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   840
            TabIndex        =   5
            Top             =   360
            Width           =   390
         End
      End
   End
End
Attribute VB_Name = "HISTORIAL"
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
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
'Dim notas As notis
'Dim alugru As grupoalu
'Dim leyfin As leyenfin
'Dim hisnotas As hisnotis
'Dim hisalugru As hisgrupoalu
'Dim hisleyfin As hisleyenfin
'Dim alumno As maestroalum
If Dir(Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & ".fin") <> "" Then
    MsgBox "NO SE PUEDE MODIFICAR EL HISTORIAL DE ESTE AÑO", 16
    Exit Sub
End If
If Dir(Ruta & "historia\" & Combo1.Text & "\*.*") <> "" Then
    MS1 = "YA EXISTE EL HISTORIAL PARA EL AÑO " & Combo1.Text & ", DESEA REMPLAZARLO?"
Else
    MS1 = "DESEA CREAR EL HISTORIAL DEL AÑO " & Combo1.Text & "?"
    If Dir(Ruta & "historia\" & Combo1.Text & "\", vbDirectory) = "" Then
        MkDir Ruta & "historia\" & Combo1.Text
    End If
End If
NGRA = Dir(Ruta & "*.obs")
If NGRA = "" Then
    MsgBox "NO HAY INFORMACION PARA CREAR HISTORIAL", 48
    Exit Sub
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton1, "CREAR HISTORIAL")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    
    NGRA = Dir(Ruta & "*.*")
    Do While NGRA <> ""
        FileCopy Ruta & NGRA, Ruta & "historia\" & Combo1.Text & "\" & NGRA
        NGRA = Dir
    Loop

'    Do While NGRA <> ""
'        VV = 0
'        Open Ruta & NGRA For Random As #NAR Len = Len(notas)
'        While Not EOF(NAR)
'            VV = VV + 1
'            Get #NAR, VV, notas
'        Wend
'        Close #NAR
'        Open Ruta & NGRA For Random As #NAR Len = Len(notas)
'        For J = 1 To VV - 1
'            Get #NAR, J, notas
'            If RTrim(notas.num_carnet) = "" Then
'                GoTo sablaba
'            End If
'            NAR = FreeFile
'            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'            Get #NAR, Val(notas.num_carnet), alumno
'            Close #NAR
'            Open Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & NGRA For Random As #NAR Len = Len(hisnotas)
'            hisnotas.nombres = RTrim(alumno.nombres)
'            hisnotas.apellidos = RTrim(alumno.apellidos)
'            hisnotas.FA = notas.FA
'            hisnotas.FA = notas.FA
'            For I = 1 To 10
'                hisnotas.nota(I) = notas.area(I)
'            Next I
'            Put #NAR, J, hisnotas
'            Close #NAR
'            NAR = NAR - 1
'sablaba:
'        Next J
'        Close #NAR
'        NGRA = Dir
'    Loop
'    NGRA = Dir(Ruta & "*.gru")
'    Do While NGRA <> ""
'        VV = 0
'        Open Ruta & NGRA For Random As #NAR Len = Len(alugru)
'        While Not EOF(NAR)
'            VV = VV + 1
'            Get #NAR, VV, alugru
'        Wend
'        Close #NAR
'        Open Ruta & NGRA For Random As #NAR Len = Len(alugru)
'        For J = 1 To VV - 1
'            Get #NAR, J, alugru
'            NAR = FreeFile
'            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'            Get #NAR, Val(alugru.num_carnet), alumno
'            Close #NAR
'            Open Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & NGRA For Random As #NAR Len = Len(hisalugru)
'            hisalugru.nombres = RTrim(alumno.nombres)
'            hisalugru.apellidos = RTrim(alumno.apellidos)
'            Put #NAR, J, hisalugru
'            Close #NAR
'            NAR = NAR - 1
'        Next J
'        Close #NAR
'        NGRA = Dir
'    Loop
'    NGRA = Dir(Ruta & "*.lrf")
'    Do While NGRA <> ""
'        VV = 0
'        Open Ruta & NGRA For Random As #NAR Len = Len(leyfin)
'        While Not EOF(NAR)
'            VV = VV + 1
'            Get #NAR, VV, leyfin
'        Wend
'        Close #NAR
'        Open Ruta & NGRA For Random As #NAR Len = Len(leyfin)
'        For J = 1 To VV - 1
'            Get #NAR, J, leyfin
'            NAR = FreeFile
'            Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
'            Get #NAR, Val(leyfin.num_carnet), alumno
'            Close #NAR
'            Open Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & NGRA For Random As #NAR Len = Len(hisleyfin)
'            hisleyfin.nombres = RTrim(alumno.nombres)
'            hisleyfin.apellidos = RTrim(alumno.apellidos)
'            For I = 1 To 5
'                hisleyfin.fnob(I) = leyfin.fnob(I)
'            Next I
'            Put #NAR, J, hisleyfin
'            Close #NAR
'            NAR = NAR - 1
'        Next J
'        Close #NAR
'        NGRA = Dir
'    Loop
'
'    NGRA = Dir(Ruta & "*.orf")
'    Do While NGRA <> ""
'        FileCopy Ruta & NGRA, Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & NGRA
'        NGRA = Dir
'    Loop
'
'    FileCopy Ruta & "infcur.edu", Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & "infcur.edu"
'    FileCopy Ruta & "areagra.edu", Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & "areagra.edu"
'    FileCopy Ruta & "materia.edu", Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & "materia.edu"
'    If (Dir(Ruta & "promovido.txt") <> "") And (Dir(Ruta & "rangpro.txt") <> "") Then
'        FileCopy Ruta & "promovido.txt", Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & "promovido.txt"
'        FileCopy Ruta & "rangpro.txt", Ruta & "historia\" & Combo1.Text & "\" & Combo1.Text & "rangpro.txt"
'    End If
    Screen.MousePointer = 0
    MsgBox "EL HISTORIAL PARA EL AÑO " & Combo1.Text & " SE CREO CON EXITO", 64
    Unload Me
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Crea el historial de los resultados académicos de cada alumno de un año determinado."
End Sub

Private Sub Form_Load()
For I = 1998 To 2100
Combo1.AddItem I
Next I
Combo1.Text = Combo1.List(0)
End Sub
