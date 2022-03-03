VERSION 5.00
Begin VB.Form DriveCSV 
   Caption         =   "Seleccionar archivo CSV"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   4560
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   5775
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5775
      End
   End
End
Attribute VB_Name = "DriveCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RutaCSV = Dir1.Path & "\" & File1.FileName
If Dir(RutaCSV) = "" Then
    MsgBox "Seleccione primero el archivo a importar", 48, "Importar CSV"
    Exit Sub
End If
Extension = Split(File1.FileName, ".")
If (Extension(1) <> "csv") And (Extension(1) <> "CSV") Then
    MsgBox "El archivo seleccionado no es de tipo CSV", 48, "Importar CSV"
    Exit Sub
End If

RESP = MsgBox("Desea seleccionar el archivo CSV: " & RutaCSV, vbYesNo + vbQuestion + vbDefaultButton1, "Importar CSV")
If RESP = vbYes Then
    Unload Me
    Import_CSV.Command1.Enabled = True
    Import_CSV.Show
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Frame1.Caption = Dir1.Path & "\" & File1.FileName
End Sub
