VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form LIST_PRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de profesores existentes"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "LIST_PRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin MSFlexGridLib.MSFlexGrid MATI17 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   12632256
      End
   End
End
Attribute VB_Name = "LIST_PRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim profe As maestropro
'Dim ini As inicio
PAG = 1
RESP = MsgBox("DESEA IMPRIMIR LA LISTA DE PROFESORES EXISTENTES?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
NAR = FreeFile
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
Printer.ScaleMode = 7
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "Pág." & PAG
Printer.CurrentY = 2
Printer.CurrentX = 8.5
Printer.Font.Size = 10
Printer.Print "LISTA DE PROFESORES"
Printer.CurrentY = 3
Printer.CurrentX = 1.5
Printer.Print ini.nombre
Printer.CurrentY = 4
Printer.CurrentX = 1.5
Printer.Print "No.";
Printer.CurrentX = 3
Printer.Print "NOMBRES Y APELLIDOS";
Printer.CurrentX = 10
Printer.Print "DIRECCION";
Printer.CurrentX = 17
Printer.Print "TELEFONO";
Printer.CurrentY = 5
term = 0
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
For ww = 1 To (MATI17.Rows - 1)
Get #NAR, ww, profe
If ((RTrim(profe.nombres) = "") And (RTrim(profe.especiali) = "")) Then
term = term + 1
GoTo samsu
End If
Printer.CurrentX = 1.5
Printer.Font.Size = 8
Printer.Print ww;
Printer.CurrentX = 3
Printer.Print RTrim(profe.nombres) & " " & RTrim(profe.apellidos);
Printer.CurrentX = 10
Printer.Print RTrim(profe.direccion);
Printer.CurrentX = 17
Printer.Print RTrim(profe.Telefono)
If ((ww - term) Mod 58) = 0 Then
PAG = PAG + 1
Printer.NewPage
Printer.CurrentY = 1.5
Printer.CurrentX = 19
Printer.Print "Pág." & PAG
Printer.CurrentY = 2
Printer.CurrentX = 8.5
Printer.Font.Size = 10
Printer.Print "LISTA DE PROFESORES"
Printer.CurrentY = 3
Printer.CurrentX = 1.5
Printer.Print ini.nombre
Printer.CurrentY = 4
Printer.CurrentX = 1.5
Printer.Print "No.";
Printer.CurrentX = 3
Printer.Print "NOMBRES Y APELLIDOS";
Printer.CurrentX = 10
Printer.Print "DIRECCION";
Printer.CurrentX = 17
Printer.Print "TELEFONO";
Printer.CurrentY = 5
End If
samsu:
Next ww
Close #NAR
Printer.EndDoc
Printer.Font.Size = 8
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Muestra la lista de profesores existentes."
End Sub

Private Sub MATI17_DblClick()
'Dim profe As maestropro
If (MATI17.Col = 0) And (MATI17.Row > 0) Then
    NAR = FreeFile
    w = Val(MATI17.Text)
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    Get #NAR, w, profe
    Close #NAR
    If ((RTrim(profe.nombres) = "") And (RTrim(profe.especiali) = "")) Then
        MsgBox "REGISTRO NO EXISTE", 16, "CONSULTAR"
        Exit Sub
    End If
    CONS_PRO.Text1.Text = RTrim(profe.nombres)
    CONS_PRO.Text2.Text = RTrim(profe.apellidos)
    CONS_PRO.Text3.Text = RTrim(profe.documento)
    CONS_PRO.Text11.Text = RTrim(profe.fech_nacim)
    CONS_PRO.Text4.Text = RTrim(profe.rh)
    CONS_PRO.Text5.Text = RTrim(profe.direccion)
    CONS_PRO.Text6.Text = RTrim(profe.Telefono)
    CONS_PRO.Text7.Text = RTrim(profe.año_ingre)
    CONS_PRO.Text8.Text = RTrim(profe.especiali)
    CONS_PRO.Text10.Text = RTrim(profe.escalafon)
    CONS_PRO.Text9.Text = MATI17.Text
    If Dir(Ruta & "FOTOPRO\" & w & ".jpg") <> "" Then
        CONS_PRO.picture2.Picture = LoadPicture(Ruta & "FOTOPRO\" & w & ".jpg")
    End If
    CONS_PRO.Show
End If
End Sub
