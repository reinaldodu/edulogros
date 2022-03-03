VERSION 5.00
Begin VB.Form HELP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda rápida"
   ClientHeight    =   3285
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5850
   Icon            =   "AYUDA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Siguiente"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "AYUDA.frx":0442
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "AYUDA RAPIDA DE MENUS:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   540
         TabIndex        =   4
         Top             =   180
         Width           =   2145
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "HELP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' La base de datos de sugerencias.
Dim Tips As New Collection

' Nombre del archivo de sugerencias
Const TIP_FILE = "HELP.TXT"

' Índice de la colección con la sugerencia visualizada actualmente.
Dim CurrentTip As Long


Private Sub DoNextTip()

    'recorre las sugerencias por orden

    CurrentTip = CurrentTip + 1
    If Tips.Count < CurrentTip Then
       CurrentTip = 1
    End If
    
    ' Muestra la sugerencia.
    HELP.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Se lee cada sugerencia desde el archivo.
    Dim InFile As Integer   ' Descriptor del archivo.
    
    ' Obtiene el siguiente descriptor libre.
    InFile = FreeFile
    
    ' Comprueba que se ha especificado un archivo.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Comprueba que existe un archivo antes de abrirlo.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Lee la colección desde el archivo de texto.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Muestra la primera sugerencia.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Ayuda rápida de menús. Para avanzar de click en Siguiente."
End Sub

Private Sub Form_Load()
          
     ' Lee el archivo de ayuda
    If LoadTips(Ruta & TIP_FILE) = False Then
        lblTipText.Caption = "No se encontró el archivo " & TIP_FILE & "."
    End If
    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
