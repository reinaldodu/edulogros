VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Ejes_Contenidos 
   Caption         =   "Ejes temáticos y contenidos"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11415
      Begin VB.CommandButton Command6 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9720
         TabIndex        =   14
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   5040
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MTEjes 
         Height          =   4575
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8070
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   9840
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Ejes_Contenidos.frx":0000
         Left            =   3600
         List            =   "Ejes_Contenidos.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Ejes_Contenidos.frx":00AE
         Left            =   960
         List            =   "Ejes_Contenidos.frx":00BE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Materia:"
         Height          =   195
         Left            =   5640
         TabIndex        =   6
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grado:"
         Height          =   195
         Left            =   3000
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
End
Attribute VB_Name = "Ejes_Contenidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If VALI380 = False Then
   Call Command5_Click
End If
Frame1.Caption = ""
MTEjes.Rows = 1
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
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
Open Ruta & fl & ser & que & lw & ".eje" For Random As #NAR Len = Len(semanal_ejetematico)
While Not EOF(NAR)
    CROA = CROA + 1
    Get #NAR, CROA, semanal_ejetematico
Wend
Close #NAR


For Y = 1 To CROA - 1
    Open Ruta & fl & ser & que & lw & ".eje" For Random As #NAR Len = Len(semanal_ejetematico)
    Get #NAR, Y, semanal_ejetematico
    Close #NAR
    MTEjes.Rows = MTEjes.Rows + 1
    MTEjes.TextMatrix(Y, 0) = Y
    MTEjes.Row = Y
    MTEjes.Col = 1
    MTEjes.CellFontBold = True
    MTEjes.Text = Trim(semanal_ejetematico.txt_eje)
    'MTEjes.TextMatrix(Y, 0) = semanal_ejetematico.txt_eje
    t = 0
    Open Ruta & fl & ser & que & lw & ".ctd" For Random As #NAR Len = Len(semanal_contenidos)
    While Not EOF(NAR)
        t = t + 1
        Get #NAR, t, semanal_contenidos
        If semanal_contenidos.num_eje = Y Then
            MTEjes.TextMatrix(Y, 2) = MTEjes.TextMatrix(Y, 2) & Trim(semanal_contenidos.txt_cont) & vbCrLf
            MTEjes.RowHeight(Y) = MTEjes.RowHeight(Y) + 240
        End If
    Wend
    Close #NAR
Next Y
Frame1.Caption = "GRADO: " & Combo2.Text & " - " & "MATERIA: " & Combo3.Text & " (" & que & ")" & " - " & "PERIODO: " & Combo1.Text
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True

End Sub

Private Sub Command2_Click()
ValiModifica = False
New_Eje.Show 1
End Sub

Private Sub Command3_Click()
If Val(MTEjes.Rows - 1) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA ELIMINAR", 64
    Exit Sub
End If
TTT = InputBox("Escriba el número del eje temático a eliminar (tenga en cuenta que se eliminaran también los contenidos del eje temático)", "Eliminar eje temático")
If TTT = "" Then
    MsgBox "No escribió el No. del eje temático", 64, "Eliminar eje temático"
    Exit Sub
End If
If (Val(TTT) > Val(MTEjes.Rows - 1)) Or (Val(TTT) < 1) Then
    MsgBox "Número de eje temático no existe", 64, "Eliminar eje temático"
    Exit Sub
End If
MTEjes.RemoveItem Val(TTT)
For I = 1 To Val(MTEjes.Rows - 1)
    MTEjes.TextMatrix(I, 0) = I
Next I
VALI380 = False
End Sub

Private Sub Command4_Click()
TTT = InputBox("Escriba el número de eje temático que desea modificar", "modificar eje temático")
If TTT = "" Then
    MsgBox "No escribió el número del eje temático", 64, "Modificar"
    Exit Sub
End If
If Val(TTT) > Val(MTEjes.Rows - 1) Or (Val(TTT) < 1) Then
    MsgBox "No existe este número de eje temático", 32, "Modificar"
    Exit Sub
End If

ValiModifica = True
New_Eje.Show 1
End Sub

Private Sub Command5_Click()
If Val(MTEjes.Rows - 1) = 0 Then
    MsgBox "NO HAY INFORMACION PARA GUARDAR", 64
    Exit Sub
End If
On Error Resume Next
Err.Clear
If Dir(Ruta & "inicial.edu") = "" Then
    MsgBox "DATOS DEL SISTEMA NO ENCONTRADOS", 16, "ADVERTENCIA"
    End
End If
MS1 = "Desea guardar esta información?"
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR")
If RESP = vbYes Then
   Screen.MousePointer = 11
   Kill Ruta & fl & ser & que & lw & ".eje"
   Kill Ruta & fl & ser & que & lw & ".ctd"
   z = 0
   For h = 1 To Val(MTEjes.Rows - 1)
        If Trim(MTEjes.TextMatrix(h, 1)) <> "" Then
            
            semanal_ejetematico.txt_eje = Trim(MTEjes.TextMatrix(h, 1))
            NAR = FreeFile
            Open Ruta & fl & ser & que & lw & ".eje" For Random As #NAR Len = Len(semanal_ejetematico)
            Put #NAR, h, semanal_ejetematico
            Close #NAR
            
            If Trim(MTEjes.TextMatrix(h, 2)) <> "" Then
                ArrCont = Split(MTEjes.TextMatrix(h, 2), vbCrLf)
                
                For J = 0 To UBound(ArrCont)
                    If Trim(ArrCont(J)) <> "" Then
                        semanal_contenidos.txt_cont = Trim(ArrCont(J))
                        semanal_contenidos.num_eje = h
                        z = z + 1
                        NAR = FreeFile
                        Open Ruta & fl & ser & que & lw & ".ctd" For Random As #NAR Len = Len(semanal_contenidos)
                        Put #NAR, z, semanal_contenidos
                        Close #NAR
                    End If
                Next
            End If
        End If
   Next h
   Screen.MousePointer = 0
End If
VALI380 = True
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
MTEjes.Row = 0
MTEjes.Col = 0
MTEjes.ColWidth(0) = 400
MTEjes.CellForeColor = RGB(255, 255, 255)
MTEjes.CellBackColor = RGB(0, 0, 150)
MTEjes.Text = "No."
MTEjes.Col = 1
MTEjes.ColWidth(1) = 5000
MTEjes.CellForeColor = RGB(255, 255, 255)
MTEjes.CellBackColor = RGB(0, 0, 150)
MTEjes.Text = "EJES TEMÁTICOS"
MTEjes.Col = 2
MTEjes.ColWidth(2) = 5300
MTEjes.CellForeColor = RGB(255, 255, 255)
MTEjes.CellBackColor = RGB(0, 0, 150)
MTEjes.Text = "CONTENIDOS"


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


'If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") Then
'    Command5.Enabled = True
'    NAR = FreeFile
'    cona = 0
'    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
'    While Not EOF(NAR)
'        cona = cona + 1
'        Get #NAR, cona, argra
'        If argra.num_pro = Val(MENUPROFE.LBLNumProfe.Caption) Then
'            VALI2 = False
'            For I = 0 To (Combo2.ListCount - 1)
'                If Combo2.List(I) = RTrim(argra.grado) Then
'                    VALI2 = True
'                    Exit For
'                End If
'            Next I
'            If VALI2 = False Then
'                Combo2.AddItem RTrim(argra.grado)
'            End If
'            NAR = FreeFile
'            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
'            Get #NAR, argra.num_area, mate
'            Close #NAR
'            NAR = NAR - 1
'            VALI2 = False
'            For I = 0 To (Combo3.ListCount - 1)
'                If Combo3.List(I) = RTrim(mate.nom) Then
'                    VALI2 = True
'                    Exit For
'                End If
'            Next I
'            If VALI2 = False Then
'                Combo3.AddItem RTrim(mate.nom)
'            End If
'        End If
'    Wend
'    Close #NAR
'    Combo1.Text = Combo1.List(0)
'    Combo2.Text = Combo2.List(0)
'    Combo3.Text = Combo3.List(0)
'    If (RTrim(Combo2.Text) = "") Or (RTrim(Combo2.Text) = "") Then
'        Command1.Enabled = False
'    End If
'Else
'    Command5.Enabled = False
'End If

Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
VALI380 = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If VALI380 = False Then
   Call Command5_Click
   Unload Me
Else
   Unload Me
End If

End Sub
