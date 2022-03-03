VERSION 5.00
Begin VB.Form INI_ANO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicio de año académico"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "INI_ANO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AÑO QUE FINALIZA:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1530
      End
   End
End
Attribute VB_Name = "INI_ANO"
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
NGRA = Dir(Ruta & "historia\" & Combo1.Text & "\prinalu.edu")
If NGRA = "" Then
    MsgBox "DEBE CREAR PRIMERO EL HISTORIAL PARA EL AÑO " & Combo1.Text, 48
    Unload Me
    HISTORIAL.Show 1
    Exit Sub
End If
RESP = MsgBox("2. DESEA INICIAR LOS ARCHIVOS DEL AÑO " & Combo1.Text & "? (ADVERTENCIA: SE BORRARÁ TODA LA INFORMACIÓN ACTUAL)", vbYesNo + vbQuestion + vbDefaultButton1, "INICIALIZACIÓN DE AÑO ACADEMICO")
     If RESP = vbYes Then
'        On Error Resume Next
'        Err.Clear
'        If (Dir("a:\setup.lst") <> "") And (Dir("a:\edulog1.cab") <> "") Then
'           Screen.MousePointer = 11
'           If Err.Number <> 0 Then
'                If Err.Number = 52 Then
'                    MsgBox "INSERTE EL DISKETTE #1 DE INSTALACION", 16, "ADVERTENCIA"
'                    Screen.MousePointer = 0
'                    Exit Sub
'                Else
'                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
'                    Screen.MousePointer = 0
'                    Exit Sub
'                End If
'           End If
           If Dir(Ruta & "logs.edu") <> "" Then
              Kill Ruta & "logs.edu"
           End If
           If Dir(Ruta & "pensi.edu") <> "" Then
              Kill Ruta & "pensi.edu"
           End If
           If Dir(Ruta & "clapro.edu") <> "" Then
              Kill Ruta & "clapro.edu"
           End If
           If Dir(Ruta & "infnota.edu") <> "" Then
              Kill Ruta & "infnota.edu"
           End If
           If Dir(Ruta & "infosub.edu") <> "" Then
              Kill Ruta & "infosub.edu"
           End If
           If Dir(Ruta & "infcur.edu") <> "" Then
              Kill Ruta & "infcur.edu"
           End If
           If Dir(Ruta & "areagra.edu") <> "" Then
              Kill Ruta & "areagra.edu"
           End If
           If Dir(Ruta & "*.gru") <> "" Then
              Kill Ruta & "*.gru"
           End If
           If Dir(Ruta & "*.obs") <> "" Then
              Kill Ruta & "*.obs"
           End If
           If Dir(Ruta & "*.obp") <> "" Then
              Kill Ruta & "*.obp"
           End If
           If Dir(Ruta & "*.lgr") <> "" Then
              Kill Ruta & "*.lgr"
           End If
           If Dir(Ruta & "*.ptj") <> "" Then
              Kill Ruta & "*.ptj"
           End If
           If Dir(Ruta & "*.lrf") <> "" Then
              Kill Ruta & "*.lrf"
           End If
           If Dir(Ruta & "*.orf") <> "" Then
              Kill Ruta & "*.orf"
           End If
           If Dir(Ruta & "*.dsp") <> "" Then
              Kill Ruta & "*.dsp"
           End If
           ' Eliminar planeadores -competencias, contenidos y ejes-
           If Dir(Ruta & "*.cpt") <> "" Then
              Kill Ruta & "*.cpt"
           End If
           If Dir(Ruta & "*.ctd") <> "" Then
              Kill Ruta & "*.ctd"
           End If
           If Dir(Ruta & "*.eje") <> "" Then
              Kill Ruta & "*.eje"
           End If
           If Dir(Ruta & "*.pln") <> "" Then
              Kill Ruta & "*.pln"
           End If
           
           
           Screen.MousePointer = 0
           ELIMINAUTO.Text1.Text = INI_ANO.Combo1.Text
           Unload Me
           MsgBox "INICIACION DE AÑO ACADEMICO TERMINO CON EXITO", 64, "OPERACION EXITOSA"
           ELIMINAUTO.Show 1
'        Else
'           Screen.MousePointer = 0
'           MsgBox "VIOLACION DE SEGURIDAD, DISCO INCORRECTO; NO SE CREARON LOS ARCHIVOS DE INICIO DE AÑO", 16
'           Unload Me
'        End If
     Else
        Unload Me
     End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Inicialización de año académico.  Elija el año que finaliza."
End Sub

Private Sub Form_Load()
For I = 2000 To 2100
Combo1.AddItem I
Next I
Combo1.Text = Combo1.List(0)
End Sub
