VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONS_NOTA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta e impresión de boletines"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10365
   Icon            =   "CONS_NOTA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir Grupo"
      Height          =   735
      Left            =   5640
      Picture         =   "CONS_NOTA.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Impresión de boletines por grupo"
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   5640
      Picture         =   "CONS_NOTA.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Impresión del boletín de acuerdo al código ingresado"
      Top             =   5160
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "CONS_NOTA.frx":0646
      Left            =   8760
      List            =   "CONS_NOTA.frx":0659
      TabIndex        =   0
      Text            =   "PRIMERO"
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONSULTAR BOLETIN"
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   4935
      Begin VB.CommandButton Command4 
         Caption         =   "Copiar datos"
         Height          =   375
         Left            =   3360
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   360
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   840
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid MATI20 
         Height          =   3855
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   12632256
         GridColor       =   12582912
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3240
         TabIndex        =   12
         Top             =   4680
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "IMPRIMIR BOLETIN"
      ForeColor       =   &H00800000&
      Height          =   2175
      Left            =   5400
      TabIndex        =   19
      Top             =   4800
      Width           =   4815
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   3120
         Stretch         =   -1  'True
         ToolTipText     =   "Copiar imagen head2.jpg en el directorio de datos."
         Top             =   240
         Width           =   1545
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   1440
         Stretch         =   -1  'True
         ToolTipText     =   "Copiar imagen head1.jpg en el directorio de datos."
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Imagen derecha"
         Height          =   195
         Left            =   3360
         TabIndex        =   21
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Imagen izquierda"
         Height          =   195
         Left            =   1560
         TabIndex        =   20
         Top             =   1800
         Width           =   1200
      End
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4800
      Width           =   4935
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4320
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3120
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1560
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PERIODO..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7680
      TabIndex        =   9
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "CONS_NOTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OkObs As Boolean, OkDes As Boolean

' ******* FUNCION QUE OBTIENE EL TOTAL DE LOGROS ACUMULADOS  Y EL TOTAL DE LOGROS ALCANZADOS POR PERIODOS (EJ: 5,4)*******
Private Function Total_Logros(DSPeriodo2 As Integer) As String
Dim Cont_Lgr2 As Integer, Lgr_Ttl2 As Integer, Cp2 As Integer
Lgr_Ttl = 0
Cp2 = 0
For lwe3 = 1 To DSPeriodo2
    If Dir(Ruta & fl & seri & argra.num_area & lwe3 & ".lgr") <> "" Then
    'OBTENER TOTAL DE LOGROS
        CROA2 = 0
        Cont_Lgr2 = 0
        NAR = FreeFile
        Open Ruta & fl & seri & argra.num_area & lwe3 & ".lgr" For Random As #NAR Len = Len(logru)
        While Not EOF(NAR)
            CROA2 = CROA2 + 1
            Get #NAR, CROA2, logru
            If Trim(logru.indicador) = "L" Then
                Cont_Lgr2 = Cont_Lgr2 + 1
                Lgr_Ttl2 = Lgr_Ttl2 + 1
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
        
        If Dir(Ruta & Combo4.Text & argra.num_area & lwe3 & ".dsp") <> "" Then
            'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
            NAR = FreeFile
            VV2 = 0
            Open Ruta & Combo4.Text & argra.num_area & lwe3 & ".dsp" For Random As #NAR Len = Len(notas_desemp)
            While Not EOF(NAR)
                VV2 = VV2 + 1
                Get #NAR, VV2, notas_desemp
                If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                    GoTo PorcentEncontrado3
                End If
            Wend
PorcentEncontrado3:
            Close #NAR
            NAR = NAR - 1
            
            For w2 = 1 To Cont_Lgr2
                'Se verifican los logros mayores al 69% (o rango menor)
                If notas_desemp.porcentaje(w2) > confdesemp.rango(3) Then
                    Cp2 = Cp2 + 1
                End If
            Next w2
        End If
        
    End If
Next lwe3
Total_Logros = Lgr_Ttl2 & "," & Cp2
End Function

' ******* FUNCION QUE OBTIENE EL DESEMPEÑO ACUMULADO POR PERIODOS Y LA DEFINITIVA (FINAL) *******
Private Function DefinitivaAcum(DsPeriodo As Integer) As String
Dim VeriManual As Boolean
Dim w As Integer, CP As Integer, PorcentLogro As Single, PromLogros As Single, SumDesemp As Long, PorcentManual(10) As Integer, ContPorcent As Integer
DefinitivaAcum = ""

' VERIFICAR PORCENTAJES DE LOGROS AUTOMATICOS O MANUALES
VeriManual = False
If Dir(Ruta & "conf_logro.edu") <> "" Then
    NAR = FreeFile
    Open Ruta & "conf_logro.edu" For Input As #NAR
    Input #NAR, ConfLgr
    Close #NAR
    NAR = NAR - 1
    If ConfLgr = 1 Then
        VeriManual = True
    End If
End If

'OBTENER ARREGLOS RANGOS DE DESEMPEÑOS Y CONVENSIONES DE DESEMPEÑOS.
NAR = FreeFile
Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
For h = 1 To 14
    Get #NAR, h, confdesemp
    If Trim(argra.grado) = Trim(confdesemp.grado) Then
        Exit For
    End If
Next h
Close #NAR
NAR = NAR - 1

CP = 0
SumDesemp = 0
Lgr_Ttl = 0

For lwe2 = 1 To DsPeriodo
'If DsPeriodo <> lwe2 Or DsPeriodo <> 5 Then
    If Dir(Ruta & Combo4.Text & argra.num_area & lwe2 & ".dsp") <> "" Then
        
    
        'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
        NAR = FreeFile
        VV2 = 0
        Open Ruta & Combo4.Text & argra.num_area & lwe2 & ".dsp" For Random As #NAR Len = Len(notas_desemp)
        While Not EOF(NAR)
            VV2 = VV2 + 1
            Get #NAR, VV2, notas_desemp
            If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                GoTo PorcentEncontrado2
            End If
        Wend
PorcentEncontrado2:
        Close #NAR
        NAR = NAR - 1
        
        'OBTENER TOTAL DE LOGROS
        CROA = 0
        Cont_Lgr = 0
        NAR = FreeFile
        Open Ruta & fl & seri & argra.num_area & lwe2 & ".lgr" For Random As #NAR Len = Len(logru)
        While Not EOF(NAR)
            CROA = CROA + 1
            Get #NAR, CROA, logru
            If Trim(logru.indicador) = "L" Then
                Cont_Lgr = Cont_Lgr + 1
                Lgr_Ttl = Lgr_Ttl + 1
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
       
        ContPorcent = 0
        If VeriManual = False Then
            ' ******* SI LOS PORCENTAJES SON AUTOMATICOS SE HACE LO SIGUIENTE: *******
            For w = 1 To Cont_Lgr
                'Se verifican los logros mayores al 69% (o rango menor)
                If notas_desemp.porcentaje(w) > confdesemp.rango(3) Then
                    CP = CP + 1
                End If
                If notas_desemp.porcentaje(w) = 0 Then
                    ContPorcent = ContPorcent + 1
                End If
                SumDesemp = SumDesemp + notas_desemp.porcentaje(w)
            Next w
            'Se verifica que tenga grabado porcentajes para obtener el desempeño final
            If ContPorcent = Cont_Lgr Then
                'DefinitivaAcum = ""
                'Exit Function
                'Exit For
                GoTo SaltaDesemp2
            End If
            PorcentLogro = (100 / Lgr_Ttl) * CP
            If PorcentLogro <> 0 Then
                PorcentLogro = Format(PorcentLogro, "#")
            End If
            PromLogros = SumDesemp / Lgr_Ttl
            If PromLogros >= 0.5 Then
                PromLogros = Format(PromLogros, "#")
            End If
            'REPORTE POR DESEMPEÑOS
            If PorcentLogro <= confdesemp.rango(3) Then
                    ' ***Pierde si el porcentaje total de logros alcanzados es menor o igual al Rango inferior
                DefinitivaAcum = Trim(confdesemp.desemp(4))
            Else
                ' ***Si alcanza el porcentaje mínimo de logros, se promedia para obtener el desempeño
                If PromLogros <= confdesemp.rango(3) Then
                    DefinitivaAcum = Trim(confdesemp.desemp(3))
                End If
                If (PromLogros > confdesemp.rango(3)) And (PromLogros <= confdesemp.rango(2)) Then
                     DefinitivaAcum = Trim(confdesemp.desemp(3))
                End If
                
                If (PromLogros > confdesemp.rango(2)) And (PromLogros <= confdesemp.rango(1)) Then
                    DefinitivaAcum = Trim(confdesemp.desemp(2))
                End If
                If PromLogros > confdesemp.rango(1) Then
                    DefinitivaAcum = Trim(confdesemp.desemp(1))
                End If
            End If
        
        Else
            If Dir(Ruta & fl & seri & argra.num_area & lwe2 & ".ptj") = "" Then
                DefinitivaAcum = "ERROR"
                Exit Function
            Else
                NAR = FreeFile
                Open Ruta & fl & seri & argra.num_area & lwe2 & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                For h = 1 To Cont_Lgr
                    Get #NAR, h, porcent_manual
                    PorcentManual(h) = porcent_manual.porcent_logro
                Next h
                Close #NAR
                NAR = NAR - 1
            End If
                  
            ' ******* SI LOS PORCENTAJES SON MANUALES SE HACE LO SIGUIENTE: *******
            For w = 1 To Cont_Lgr
                If notas_desemp.porcentaje(w) = 0 Then
                    ContPorcent = ContPorcent + 1
                End If
                SumDesemp = SumDesemp + (notas_desemp.porcentaje(w) * PorcentManual(w))
            Next w
            'Se verifica que tenga grabado porcentajes para obtener el desempeño final
            If ContPorcent = Cont_Lgr Then
                'DefinitivaAcum = ""
                'Exit Function
                'Exit For
                GoTo SaltaDesemp2
            End If
            If DsPeriodo <> 5 Then
                PorcentLogro = SumDesemp / 100
                If PorcentLogro <> 0 Then
                    PorcentLogro = Format(PorcentLogro, "#")
                End If
            Else
                PorcentLogro = SumDesemp / (100 * lwe2)
                If PorcentLogro <> 0 Then
                    PorcentLogro = Format(PorcentLogro, "#")
                End If
            End If
            'REPORTE POR DESEMPEÑOS
            If PorcentLogro <= confdesemp.rango(3) Then
                DefinitivaAcum = Trim(confdesemp.desemp(4))
            End If
            If (PorcentLogro > confdesemp.rango(3)) And (PorcentLogro <= confdesemp.rango(2)) Then
                 DefinitivaAcum = Trim(confdesemp.desemp(3))
            End If
            
            If (PorcentLogro > confdesemp.rango(2)) And (PorcentLogro <= confdesemp.rango(1)) Then
                DefinitivaAcum = Trim(confdesemp.desemp(2))
            End If
            If PorcentLogro > confdesemp.rango(1) Then
                DefinitivaAcum = Trim(confdesemp.desemp(1))
            End If
        End If
    
    End If
'Else
'    Exit For
SaltaDesemp2:
'End If
Next lwe2
End Function

' *******FUNCION QUE OBTIENE EL DESEMPEÑO DE UN PERIODO ESPECÍFICO DE LA MATERIA*******
Private Function Definitiva(Rangos() As Byte, Desempenos() As String * 5, Porcentajes() As Byte, TtLogros As Integer, VeriLgr As Boolean) As String
Dim w As Integer, CP As Integer, PorcentLogro As Single, PromLogros As Single, SumDesemp As Long, PorcentManual(10) As Integer, ContPorcent As Integer
CP = 0
ContPorcent = 0
SumDesemp = 0
If VeriLgr = False Then
    ' ******* SI LOS PORCENTAJES SON AUTOMATICOS SE HACE LO SIGUIENTE: *******
    For w = 1 To TtLogros
        'Se verifican los logros mayores al 69% (o rango menor)
        If Porcentajes(w) > Rangos(3) Then
            CP = CP + 1
        End If
        If Porcentajes(w) = 0 Then
            ContPorcent = ContPorcent + 1
        End If
        SumDesemp = SumDesemp + Porcentajes(w)
    Next w
    'Se verifica que tenga grabado porcentajes para obtener el desempeño final
    If ContPorcent = TtLogros Then
        Definitiva = ""
        Exit Function
    End If
    PorcentLogro = (100 / TtLogros) * CP
    If PorcentLogro <> 0 Then
        PorcentLogro = Format(PorcentLogro, "#")
    End If
    PromLogros = SumDesemp / TtLogros
    If PromLogros >= 0.5 Then
        PromLogros = Format(PromLogros, "#")
    End If
    'REPORTE POR DESEMPEÑOS
    If PorcentLogro <= Rangos(3) Then
            ' ***Pierde si el porcentaje total de logros alcanzados es menor o igual al Rango inferior
        Definitiva = Trim(Desempenos(4))
    Else
        ' ***Si alcanza el porcentaje mínimo de logros, se promedia para obtener el desempeño
        If PromLogros <= Rangos(3) Then
            Definitiva = Trim(Desempenos(3))
        End If
        If (PromLogros > Rangos(3)) And (PromLogros <= Rangos(2)) Then
             Definitiva = Trim(Desempenos(3))
        End If
        
        If (PromLogros > Rangos(2)) And (PromLogros <= Rangos(1)) Then
            Definitiva = Trim(Desempenos(2))
        End If
        If PromLogros > Rangos(1) Then
            Definitiva = Trim(Desempenos(1))
        End If
    End If

Else
    If Dir(Ruta & fl & seri & argra.num_area & lwe & ".ptj") = "" Then
        Definitiva = "ERROR"
        Exit Function
    Else
        NAR = FreeFile
        Open Ruta & fl & seri & argra.num_area & lwe & ".ptj" For Random As #NAR Len = Len(porcent_manual)
        For h = 1 To Cont_Lgr
            Get #NAR, h, porcent_manual
            PorcentManual(h) = porcent_manual.porcent_logro
        Next h
        Close #NAR
        NAR = NAR - 1
    End If
          
    ' ******* SI LOS PORCENTAJES SON MANUALES SE HACE LO SIGUIENTE: *******
    For w = 1 To TtLogros
        If Porcentajes(w) = 0 Then
            ContPorcent = ContPorcent + 1
        End If
        SumDesemp = SumDesemp + (Porcentajes(w) * PorcentManual(w))
    Next w
    'Se verifica que tenga grabado porcentajes para obtener el desempeño final
    If ContPorcent = TtLogros Then
        Definitiva = ""
        Exit Function
    End If
    PorcentLogro = SumDesemp / 100
    If PorcentLogro <> 0 Then
        PorcentLogro = Format(PorcentLogro, "#")
    End If
    'REPORTE POR DESEMPEÑOS
    If PorcentLogro <= Rangos(3) Then
        Definitiva = Trim(Desempenos(4))
    End If
    If (PorcentLogro > Rangos(3)) And (PorcentLogro <= Rangos(2)) Then
         Definitiva = Trim(Desempenos(3))
    End If
    
    If (PorcentLogro > Rangos(2)) And (PorcentLogro <= Rangos(1)) Then
        Definitiva = Trim(Desempenos(2))
    End If
    If PorcentLogro > Rangos(1) Then
        Definitiva = Trim(Desempenos(1))
    End If
End If

End Function

Private Function Definitiva_Imp() As String
Dim AcumulaPorcent As Byte, NotAcumula As Single

Definitiva_Imp = ""
AcumulaPorcent = 0
NotAcumula = 0
'Cont_Lgr = 0

For ww = 1 To lwe
    If Dir(Ruta & Combo4.Text & argra.num_area & ww & ".dsp") <> "" Then
    
        'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
        NAR = FreeFile
        VV2 = 0
        Open Ruta & Combo4.Text & argra.num_area & ww & ".dsp" For Random As #NAR Len = Len(notas_desemp)
        While Not EOF(NAR)
            VV2 = VV2 + 1
            Get #NAR, VV2, notas_desemp
            If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                GoTo PorcentEncontrado
            End If
        Wend
PorcentEncontrado:
        Close #NAR
        NAR = NAR - 1
        
        'OBTENER TOTAL DE LOGROS
        CROA = 0
        Cont_Lgr = 0
        NAR = FreeFile
        Open Ruta & fl & seri & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
        While Not EOF(NAR)
            CROA = CROA + 1
            Get #NAR, CROA, logru
            If Trim(logru.indicador) = "L" Then
                Cont_Lgr = Cont_Lgr + 1
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
        
        'AcumulaPorcent = 0
        'NotAcumula = 0
        For I = 1 To Cont_Lgr
            'NAR = FreeFile
            'Open Ruta & fl & seri & argra.num_area & ww & ".lgr" For Random As #NAR Len = Len(logru)
            'Get #NAR, notas_desemp.logro(I), logru
            If notas_desemp.porcentaje(I) <> 0 Then
                NAR = FreeFile
                Open Ruta & fl & seri & argra.num_area & ww & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                Get #NAR, I, porcent_manual
                Close #NAR
                NAR = NAR - 1
                AcumulaPorcent = AcumulaPorcent + porcent_manual.porcent_logro
                NotAcumula = NotAcumula + (Val(porcent_manual.porcent_logro) * Val(notas_desemp.porcentaje(I)))
            End If
            'Close #NAR
            'NAR = NAR - 1
        Next I
        Definitiva_Imp = AcumulaPorcent & "," & NotAcumula
    End If
Next ww

End Function

Private Function CortaObs(Observacion As String)
Dim Recorrer As Integer, Cortar() As String, XSuma As Single
Cortar = Split(Observacion, " ")
XSuma = 0
Printer.CurrentX = 0.5
For Recorrer = 0 To UBound(Cortar)
    XSuma = XSuma + Printer.TextWidth(Cortar(Recorrer))
    If XSuma <= 16 Then
        Printer.FontSize = 8
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    Else
        XSuma = Printer.TextWidth(Cortar(Recorrer))
        Printer.Print ""
        Call SaltaLinea
        Printer.FontSize = 8
        Printer.CurrentX = 0.5
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    End If
Next Recorrer
Printer.Print ""
End Function

Private Function SaltaLinea()
'If (Printer.CurrentY > 26) Then
If (Printer.CurrentY > 31) Then
   'Printer.Line (15.7, 5)-(15.7, Printer.CurrentY)
   Printer.Line (18.9, 4.5)-(18.9, Printer.CurrentY)
   Printer.Line (19.9, 4.5)-(19.9, Printer.CurrentY)
   Printer.Line (21, 4.5)-(21, Printer.CurrentY)
   Printer.NewPage
   Call Encabezado
End If
End Function

Private Function Encabezado()
Printer.ScaleMode = 7
'Hoja oficio
Printer.PaperSize = 5
If Printer.Page = 1 Then
    'Imprimir logotipos
    If Image1.Picture <> 0 Then
        Printer.PaintPicture Image1.Picture, 0.5, 1, 2, 2
    End If
    If Image2.Picture <> 0 Then
        Printer.PaintPicture Image2.Picture, 19, 1, 2, 2
    End If
    'Imprimir encabezado
    Printer.CurrentY = 1
    Printer.FontName = "Monotype corsiva"
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.nombre))) / 2
    Printer.Print ini.nombre
    
    Printer.FontSize = 10
    Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.ciudad))) / 2
    Printer.Print ini.ciudad
    
    Printer.FontName = ""
    Printer.Print ""
    If (Combo3.Text <> "FINAL") Then
        Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.modalidad) & " PERIODO " & Combo3.Text & " - " & Trim(ini.Telefono))) / 2
        Printer.Print Trim(ini.modalidad) & " PERIODO " & Combo3.Text & " - " & Trim(ini.Telefono)
        Printer.Print ""
    Else
        Printer.CurrentX = (21 - Printer.TextWidth("PROMOCIÓN FINAL")) / 2
        Printer.Print "PROMOCIÓN FINAL"
        Printer.Print ""
        Printer.CurrentX = (21 - Printer.TextWidth("AÑO LECTIVO " & Trim(ini.Telefono))) / 2
        Printer.Print "AÑO LECTIVO " & Trim(ini.Telefono)
        'Printer.CurrentX = 17.5
        'Printer.Print "FOLIO No."
        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
        Printer.Line (1, Printer.CurrentY + 0.07)-(21, Printer.CurrentY + 0.07)
        Printer.CurrentY = Printer.CurrentY + 0.07
        'Printer.Print ""
    End If
    If (Combo3.Text <> "FINAL") Then
        'Imprimir el nombre del estudiante
        Printer.CurrentX = (21 - Printer.TextWidth(Frame1.Caption)) / 2
        Printer.Print Frame1.Caption
        
        'Imprimir el nombre del grupo
        Printer.CurrentX = (21 - Printer.TextWidth(Combo4.Text)) / 2
        Printer.Print Right(Combo4.Text, Len(Combo4.Text) - 1)
    Else
        'Imprimir el nombre del estudiante
        Printer.FontSize = 12
        Printer.CurrentX = (21 - Printer.TextWidth(Frame1.Caption)) / 2
        Printer.Print Frame1.Caption
        
        'Imprimir el nombre del grupo
        Printer.FontSize = 10
        Printer.CurrentX = (21 - Printer.TextWidth(Combo4.Text)) / 2
        Printer.Print Right(Combo4.Text, Len(Combo4.Text) - 1)
        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
        Printer.Line (1, Printer.CurrentY + 0.07)-(21, Printer.CurrentY + 0.07)
    End If
    Printer.FontBold = False
Else
    Printer.CurrentY = 2.5
    Printer.FontSize = 10
    Printer.FontBold = True
    If (Combo3.Text <> "FINAL") Then
        Printer.CurrentX = (21 - Printer.TextWidth(Trim(ini.modalidad) & " PERIODO " & Combo3.Text & " - " & Trim(ini.Telefono))) / 2
        Printer.Print Trim(ini.modalidad) & " PERIODO " & Combo3.Text & " - " & Trim(ini.Telefono)
        Printer.Print ""
    Else
        Printer.CurrentX = (21 - Printer.TextWidth("PROMOCIÓN FINAL")) / 2
        Printer.Print "PROMOCIÓN FINAL"
        Printer.Print ""
        Printer.CurrentX = (21 - Printer.TextWidth("AÑO LECTIVO " & Trim(ini.Telefono))) / 2
        Printer.Print "AÑO LECTIVO " & Trim(ini.Telefono)
        'Printer.CurrentX = 17.5
        'Printer.Print "FOLIO No."
        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
        Printer.Line (1, Printer.CurrentY + 0.07)-(21, Printer.CurrentY + 0.07)
        Printer.CurrentY = Printer.CurrentY + 0.07
        'Printer.Print ""
    End If
    If (Combo3.Text <> "FINAL") Then
        'Imprimir el nombre del estudiante
        Printer.CurrentX = (21 - Printer.TextWidth(Frame1.Caption)) / 2
        Printer.Print Frame1.Caption
        
        'Imprimir el nombre del grupo
        Printer.CurrentX = (21 - Printer.TextWidth(Combo4.Text)) / 2
        Printer.Print Right(Combo4.Text, Len(Combo4.Text) - 1)
    Else
        'Imprimir el nombre del estudiante
        Printer.FontSize = 12
        Printer.CurrentX = (21 - Printer.TextWidth(Frame1.Caption)) / 2
        Printer.Print Frame1.Caption
        
        'Imprimir el nombre del grupo
        Printer.FontSize = 10
        Printer.CurrentX = (21 - Printer.TextWidth(Combo4.Text)) / 2
        Printer.Print Right(Combo4.Text, Len(Combo4.Text) - 1)
        Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
        Printer.Line (1, Printer.CurrentY + 0.07)-(21, Printer.CurrentY + 0.07)
    End If
    
    Printer.FontBold = False
End If
Printer.Print ""
If (Combo3.Text <> "FINAL") Then
    Printer.FontBold = False
    Printer.CurrentX = 19
    Printer.Print "  %";
    Printer.CurrentX = 20
    Printer.Print " CAL"
Else
    Printer.FontSize = 9
    Printer.FontBold = True
    Printer.CurrentX = 1
    Printer.Print "ASIGNATURAS / CONCEPTOS";
    Printer.CurrentX = 16
    Printer.Print "I.H.";
    Printer.CurrentX = 17
    Printer.Print "VAL";
    Printer.CurrentX = 18.5
    Printer.Print "DESEMPEÑO"
End If
If (Combo3.Text <> "FINAL") Then
    Printer.Line (0.5, Printer.CurrentY)-(21, Printer.CurrentY)
Else
    Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
End If
End Function

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
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Command1_Click()
Dim VeriManual As Boolean, RowDesemp As Integer, DsFinal As String
'Dim AcumulaPorcent As Byte, NotAcumula As Single

If Dir(Ruta & Combo4.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
MATI20.Rows = 1
Screen.MousePointer = 11
If RTrim(Text2.Text) = "" Then
    MsgBox "ESCRIBA EL CODIGO DEL ESTUDIANTE", 16, "CONSULTAR OBSERVACIONES"
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
Frame1.Caption = ""
Label12.Caption = ""
ret = 0
NAR = FreeFile
Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
    'If Val(Text2.Text) = ret Then
    '    Close #NAR
    '    GoTo MIJO
    'End If
Wend
Close #NAR
If Val(Text2.Text) > ret - 1 Or Val(Text2.Text) < 1 Then
    MsgBox "CODIGO NO EXISTE EN ESTE GRUPO", 64, "CONSULTAR"
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
Else
    Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
    Get #NAR, Val(Text2.Text), alugru
    Close #NAR
End If

'MIJO:
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, Val(alugru.num_carnet), alumno
Close #NAR
Frame1.Caption = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & Text2.Text & ")."
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo4.Text Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
    End If
Wend
Close #NAR
Label9.Caption = RE22
'Label10.Caption = RTrim(Combo3.Text)
'Label11.Caption = Combo4.Text
'Label17.Caption = alugru.num_carnet
seri = Left(RE22, 3)
If JOJI = "UNICA" Then
    fl = "1"
End If
If JOJI = "MAÑANA" Then
    fl = "2"
End If
If JOJI = "TARDE" Then
    fl = "3"
End If
If JOJI = "NOCHE" Then
    fl = "4"
End If
If Combo3.Text = "PRIMERO" Then
    lwe = 1
End If
If Combo3.Text = "SEGUNDO" Then
    lwe = 2
End If
If Combo3.Text = "TERCERO" Then
    lwe = 3
End If
If Combo3.Text = "CUARTO" Then
    lwe = 4
End If
If Combo3.Text = "FINAL" Then
    lwe = 5
End If

' VERIFICAR PORCENTAJES DE LOGROS AUTOMATICOS O MANUALES
VeriManual = False
If Dir(Ruta & "conf_logro.edu") <> "" Then
    NAR = FreeFile
    Open Ruta & "conf_logro.edu" For Input As #NAR
    Input #NAR, ConfLgr
    Close #NAR
    If ConfLgr = 1 Then
        VeriManual = True
    End If
End If

If (lwe <> 5) Then
    
    cona = 0
    'h = 1
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Combo4.Text Then
            OkObs = False
            OkDes = False
            If Dir(Ruta & Combo4.Text & argra.num_area & lwe & ".obs") <> "" Then
                NAR = FreeFile
                Y = 0
                Open Ruta & Combo4.Text & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
                While Not EOF(NAR)
                    Y = Y + 1
                    Get #NAR, Y, notas
                    If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                        OkObs = True
                        NAR = FreeFile
                        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                        Get #NAR, argra.num_area, mate
                        Close #NAR
                        NAR = NAR - 1
                        MATI20.Rows = MATI20.Rows + 1
                        MATI20.Col = 0
                        MATI20.Row = MATI20.Rows - 1
                        MATI20.CellFontBold = True
                        MATI20.CellForeColor = RGB(0, 0, 255)
                        MATI20.TextMatrix(MATI20.Rows - 1, 0) = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:" & notas.FA & ")"
                        'RowDesemp = MATI20.Rows - 1
                        ValPromAcu = Definitiva_Imp
                        If ValPromAcu <> "" Then
                            ValPromAcu = Split(ValPromAcu, ",")
                            If ValPromAcu(0) <> 0 Then
                                MATI20.Col = 1
                                MATI20.Row = MATI20.Rows - 1
                                MATI20.CellFontBold = True
                                MATI20.TextMatrix(MATI20.Rows - 1, 1) = ValPromAcu(0)
                                MATI20.Col = 2
                                MATI20.Row = MATI20.Rows - 1
                                MATI20.CellFontBold = True
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = Format(ValPromAcu(1) / ValPromAcu(0), "#.00")
                            Else
                                MATI20.TextMatrix(MATI20.Rows - 1, 1) = ""
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = ""
                            End If
                        End If
                        GoTo encontrar
                    
                    End If
                Wend
encontrar:
                Close #NAR
                NAR = NAR - 1
            'SI NO EXISTE EL ARCHIVO DE OBSERVACIONES LAS FALLAS SON IGUAL A CERO.
            Else
                notas.FA = 0
            End If
            If Dir(Ruta & Combo4.Text & argra.num_area & lwe & ".dsp") <> "" Then
                NAR = FreeFile
                VV = 0
                Open Ruta & Combo4.Text & argra.num_area & lwe & ".dsp" For Random As #NAR Len = Len(notas_desemp)
                While Not EOF(NAR)
                    VV = VV + 1
                    Get #NAR, VV, notas_desemp
                    If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                        OkDes = True
                        If OkObs = False Then
                            NAR = FreeFile
                            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                            Get #NAR, argra.num_area, mate
                            Close #NAR
                            NAR = NAR - 1
                            MATI20.Rows = MATI20.Rows + 1
                            MATI20.Col = 0
                            MATI20.Row = MATI20.Rows - 1
                            MATI20.CellFontBold = True
                            MATI20.CellForeColor = RGB(0, 0, 255)
                            'MATI20.TextMatrix(MATI20.Rows - 1, 0) = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:0)"
                            MATI20.TextMatrix(MATI20.Rows - 1, 0) = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:" & notas.FA & ")"
                            'RowDesemp = MATI20.Rows - 1
                            ValPromAcu = Definitiva_Imp
                            If ValPromAcu <> "" Then
                                ValPromAcu = Split(ValPromAcu, ",")
                                If ValPromAcu(0) <> 0 Then
                                    MATI20.Col = 1
                                    MATI20.Row = MATI20.Rows - 1
                                    MATI20.CellFontBold = True
                                    MATI20.TextMatrix(MATI20.Rows - 1, 1) = ValPromAcu(0)
                                    MATI20.Col = 2
                                    MATI20.Row = MATI20.Rows - 1
                                    MATI20.CellFontBold = True
                                    MATI20.TextMatrix(MATI20.Rows - 1, 2) = Format(ValPromAcu(1) / ValPromAcu(0), "#.00")
                                Else
                                    MATI20.TextMatrix(MATI20.Rows - 1, 1) = ""
                                    MATI20.TextMatrix(MATI20.Rows - 1, 2) = ""
                                End If
                            End If
                            
                        End If
                        GoTo encontrar2
                    End If
                Wend
encontrar2:
                Close #NAR
                NAR = NAR - 1
            End If
            If OkDes = True Then
                CROA = 0
                Cont_Lgr = 0
                NAR = FreeFile
                Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                While Not EOF(NAR)
                    CROA = CROA + 1
                    Get #NAR, CROA, logru
                    If Trim(logru.indicador) = "L" Then
                        Cont_Lgr = Cont_Lgr + 1
                        
                    End If
                Wend
                Close #NAR
                'AcumulaPorcent = 0
                'NotAcumula = 0
                For I = 1 To Cont_Lgr
                    
                        NAR = FreeFile
                        Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                        Get #NAR, notas_desemp.logro(I), logru
                        If notas_desemp.porcentaje(I) <> 0 Then
                            NAR = FreeFile
                            Open Ruta & fl & seri & argra.num_area & lwe & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                            Get #NAR, I, porcent_manual
                            Close #NAR
                            NAR = NAR - 1
                        
                        
                            MATI20.Rows = MATI20.Rows + 1
                            MATI20.TextMatrix(MATI20.Rows - 1, 0) = Trim(logru.indicador) & " - " & Trim(logru.observ)
                            MATI20.TextMatrix(MATI20.Rows - 1, 1) = porcent_manual.porcent_logro
                            'AcumulaPorcent = AcumulaPorcent + porcent_manual.porcent_logro
                            MATI20.TextMatrix(MATI20.Rows - 1, 2) = notas_desemp.porcentaje(I)
                            'NotAcumula = NotAcumula + (Val(porcent_manual.porcent_logro) * Val(notas_desemp.porcentaje(I)))
                        End If
                        Close #NAR
                        NAR = NAR - 1
                        
                Next I
            
            End If
            If OkObs = True Then
                For I = 1 To 10
                    If notas.area(I) <> 0 Then
                        NAR = FreeFile
                        Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                        Get #NAR, notas.area(I), logru
                        If Trim(logru.indicador) <> "L" Then
                            MATI20.Rows = MATI20.Rows + 1
                            MATI20.TextMatrix(MATI20.Rows - 1, 0) = Trim(logru.indicador) & " - " & Trim(logru.observ)
                        End If
                        Close #NAR
                        NAR = NAR - 1
                    End If
                Next I
            End If
            If OkDes = True Or OkObs = True Then
                MATI20.Rows = MATI20.Rows + 1
            End If
                  
        End If
    Wend
    Close #NAR
    Screen.MousePointer = 0

Else
' **********  CREACIÓN DEL REPORTE FINAL  ****************
    lwe = 4
    cona = 0
    MateriaX = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Combo4.Text Then
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
            For h = 1 To 14
                Get #NAR, h, confdesemp
                If Trim(argra.grado) = Trim(confdesemp.grado) Then
                    Exit For
                End If
            Next h
            Close #NAR
            NAR = NAR - 1
            '****CONVIVENCIA ESCOLAR(27) SE DISCRIMINA EN DISCIPLINA Y CONDUCTA******
            If (mate.num <> 27) Then
                MATI20.Rows = MATI20.Rows + 1
                MATI20.Col = 0
                MATI20.Row = MATI20.Rows - 1
                MATI20.CellFontBold = True
                MATI20.CellForeColor = RGB(0, 0, 255)
                MATI20.TextMatrix(MATI20.Rows - 1, 0) = RTrim(mate.nom) & "  (I.H:" & argra.ih & ")"
                ValPromAcu = Definitiva_Imp
                If ValPromAcu <> "" Then
                    ValPromAcu = Split(ValPromAcu, ",")
                    If ValPromAcu(0) <> 0 Then
                        MATI20.Col = 1
                        MATI20.Row = MATI20.Rows - 1
                        MATI20.CellFontBold = True
                        MATI20.TextMatrix(MATI20.Rows - 1, 1) = Format(ValPromAcu(1) / ValPromAcu(0), "#.00")
                        MATI20.Col = 2
                        'VALIDA RANGOS POR DESEMPEÑO
                        If MATI20.TextMatrix(MATI20.Rows - 1, 1) <> "" Then
                            If MATI20.TextMatrix(MATI20.Rows - 1, 1) <= confdesemp.rango(3) Then
                                MATI20.CellFontBold = True
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = "BAJO"
                                
                                MateriaX = MateriaX + 1
                                
                                '**********SE VERIFICA SI LA MATERIA FUE NIVELADA*********
                                If Dir(Ruta & Combo4.Text & argra.num_area & "5.dsp") <> "" Then
                                    'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
                                    NAR = FreeFile
                                    VV44 = 0
                                    Open Ruta & Combo4.Text & argra.num_area & "5.dsp" For Random As #NAR Len = Len(notas_desemp)
                                    While Not EOF(NAR)
                                        VV44 = VV44 + 1
                                        Get #NAR, VV44, notas_desemp
                                        If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                                            If (notas_desemp.porcentaje(1) <> 0) Or (notas_desemp.porcentaje(2) <> 0) Then
                                                
                                                
                                                If (notas_desemp.porcentaje(1) >= 70) Then
                                                    MateriaX = MateriaX - 1
                                                Else
                                                    If (notas_desemp.porcentaje(1) <> 0) Then
                                                        MateriaX = MateriaX + 1
                                                    End If
                                                End If
                                                If (notas_desemp.porcentaje(2) >= 70) Then
                                                    MateriaX = MateriaX - 2
                                                Else
                                                    If (notas_desemp.porcentaje(2) <> 0) Then
                                                        MateriaX = MateriaX + 2
                                                    End If
                                                End If
                                            End If
                                            GoTo PorcentEncontrado15
                                        End If
                                    Wend
PorcentEncontrado15:
                                    Close #NAR
                                    NAR = NAR - 1
                                End If
                                
                            End If
                            If (MATI20.TextMatrix(MATI20.Rows - 1, 1) > confdesemp.rango(3)) And (MATI20.TextMatrix(MATI20.Rows - 1, 1) <= confdesemp.rango(2)) Then
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = "BÁSICO"
                            End If
                    
                            If (MATI20.TextMatrix(MATI20.Rows - 1, 1) > confdesemp.rango(2)) And (MATI20.TextMatrix(MATI20.Rows - 1, 1) <= confdesemp.rango(1)) Then
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = "ALTO"
                            End If
                            If MATI20.TextMatrix(MATI20.Rows - 1, 1) > confdesemp.rango(1) Then
                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = "SUPERIOR"
                            End If
                        End If
                    Else
                        MATI20.TextMatrix(MATI20.Rows - 1, 1) = ""
                        MATI20.TextMatrix(MATI20.Rows - 1, 2) = ""
                    End If
                End If
            Else
           '****** SE IMPRIME DISCIPLINA Y CONDUCTA******
                If Dir(Ruta & Combo4.Text & argra.num_area & "4.dsp") <> "" Then
                    NAR = FreeFile
                    VV5 = 0
                    Open Ruta & Combo4.Text & argra.num_area & "4.dsp" For Random As #NAR Len = Len(notas_desemp)
                    While Not EOF(NAR)
                        VV5 = VV5 + 1
                        Get #NAR, VV5, notas_desemp
                        If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                            If (Dir(Ruta & fl & seri & argra.num_area & "4.lgr") <> "") Then
                                NAR = FreeFile
                                Open Ruta & fl & seri & argra.num_area & "4.lgr" For Random As #NAR Len = Len(logru)
                                For TT = 1 To 2
                                    Get #NAR, notas_desemp.logro(TT), logru
                                    If (notas_desemp.porcentaje(TT) <> 0) Then
                                        MATI20.Rows = MATI20.Rows + 1
                                        MATI20.Col = 0
                                        MATI20.Row = MATI20.Rows - 1
                                        MATI20.CellFontBold = True
                                        MATI20.CellForeColor = RGB(0, 0, 255)
                                        MATI20.TextMatrix(MATI20.Rows - 1, 0) = Format(RTrim(logru.observ), ">")
                                        MATI20.Col = 1
                                        MATI20.Row = MATI20.Rows - 1
                                        MATI20.CellFontBold = True
                                        MATI20.TextMatrix(MATI20.Rows - 1, 1) = Format(notas_desemp.porcentaje(TT), "#.00")
                                        MATI20.Col = 2
                                        'VALIDA RANGOS POR DESEMPEÑO
                                        If MATI20.TextMatrix(MATI20.Rows - 1, 1) <> "" Then
                                            If MATI20.TextMatrix(MATI20.Rows - 1, 1) <= confdesemp.rango(3) Then
                                                MATI20.CellFontBold = True
                                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = "BAJO"
                                            End If
                                            If (MATI20.TextMatrix(MATI20.Rows - 1, 1) > confdesemp.rango(3)) And (MATI20.TextMatrix(MATI20.Rows - 1, 1) <= confdesemp.rango(2)) Then
                                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = "BÁSICO"
                                            End If
                                    
                                            If (MATI20.TextMatrix(MATI20.Rows - 1, 1) > confdesemp.rango(2)) And (MATI20.TextMatrix(MATI20.Rows - 1, 1) <= confdesemp.rango(1)) Then
                                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = "ALTO"
                                            End If
                                            If MATI20.TextMatrix(MATI20.Rows - 1, 1) > confdesemp.rango(1) Then
                                                MATI20.TextMatrix(MATI20.Rows - 1, 2) = "SUPERIOR"
                                            End If
                                        End If
                                    End If
                                Next TT
                                Close #NAR
                                NAR = NAR - 1
                            End If
                        End If
                    Wend
                    Close #NAR
                    NAR = NAR - 1
                End If
            End If
        End If
    Wend
    Close #NAR
    ' Abrir parámetros de promoción
    If (Dir(Ruta & "rangpro.txt") <> "") And (Dir(Ruta & "promovido.txt") <> "") Then
        Open Ruta & "rangpro.txt" For Input As #NAR
        Input #NAR, rus, fis
        Close #NAR
        Open Ruta & "promovido.txt" For Input As #NAR
        Input #NAR, SAPO2, SAPO3, SAPO4
        Close #NAR
    End If
    If MateriaX <= rus Then
        Label12.Caption = RTrim(SAPO2)
    End If
    If (MateriaX > rus) And (MateriaX <= fis) Then
        Label12.Caption = RTrim(SAPO3)
    End If
    If MateriaX > fis Then
        Label12.Caption = RTrim(SAPO4)
    End If
    Screen.MousePointer = 0
End If
End Sub

Private Sub Command2_Click()
Dim ValiSalto As Boolean, XMax As Single, GuardaY As Single, DsFinal As String
Dim Cont_Sup As Byte, Cont_Alt As Byte, Cont_Bas As Byte, Cont_Baj As Byte

Cont_Sup = 0
Cont_Alt = 0
Cont_Bas = 0
Cont_Baj = 0

If Dir(Ruta & Combo4.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
If RTrim(Text2.Text) = "" Then
    MsgBox "ESCRIBA EL CODIGO DEL ESTUDIANTE", 16, "CONSULTAR OBSERVACIONES"
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
Frame1.Caption = ""
ret = 0
NAR = FreeFile
Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
If Val(Text2.Text) > ret - 1 Or Val(Text2.Text) < 1 Then
    MsgBox "CODIGO NO EXISTE EN ESTE GRUPO", 64, "CONSULTAR"
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
Else
    Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
    Get #NAR, Val(Text2.Text), alugru
    Close #NAR
End If

Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, Val(alugru.num_carnet), alumno
Close #NAR
Frame1.Caption = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " " & "(" & Text2.Text & ")."

Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo4.Text Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
        SP = RTrim(icur.director)
    End If
Wend
Close #NAR
Label9.Caption = RE22
seri = Left(RE22, 3)
If JOJI = "UNICA" Then
    fl = "1"
End If
If JOJI = "MAÑANA" Then
    fl = "2"
End If
If JOJI = "TARDE" Then
    fl = "3"
End If
If JOJI = "NOCHE" Then
    fl = "4"
End If
If Combo3.Text = "PRIMERO" Then
    lwe = 1
End If
If Combo3.Text = "SEGUNDO" Then
    lwe = 2
End If
If Combo3.Text = "TERCERO" Then
    lwe = 3
End If
If Combo3.Text = "CUARTO" Then
    lwe = 4
End If
If Combo3.Text = "FINAL" Then
    lwe = 5
End If
cona = 0
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
Open Ruta & "comentadesemp.edu" For Input As #NAR
Input #NAR, comdpe.bajo, comdpe.basico, comdpe.alto, comdpe.superior
Close #NAR
If (Combo3.Text <> "FINAL") Then
    RESP = MsgBox("DESEA IMPRIMIR EL REPORTE DE " & Frame1.Caption & "?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR REPORTE")
    If RESP = vbYes Then
        Screen.MousePointer = 11
        Call Encabezado
    
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Combo4.Text Then
            OkObs = False
            OkDes = False
            If Dir(Ruta & Combo4.Text & argra.num_area & lwe & ".obs") <> "" Then
                NAR = FreeFile
                Y = 0
                Open Ruta & Combo4.Text & argra.num_area & lwe & ".obs" For Random As #NAR Len = Len(notas)
                While Not EOF(NAR)
                    Y = Y + 1
                    Get #NAR, Y, notas
                    If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                        OkObs = True
                        NAR = FreeFile
                        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                        Get #NAR, argra.num_area, mate
                        Close #NAR
                        NAR = NAR - 1
                        Call SaltaLinea
                        Printer.CurrentX = 0.5
                        Printer.FontSize = 9
                        Printer.FontBold = True
                        Impr_adicional = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:" & notas.FA & ")"
                        Printer.Print Impr_adicional;
                        Printer.FontSize = 8
                        Printer.CurrentX = 14.2
                        If Trim(mate.nom) <> "CONVIVENCIA ESCOLAR" Then
                            Printer.Print "Valoración promedio acumulada:";
                            'Call Definitiva_Imp
                            ValPromAcu = Definitiva_Imp
                            If ValPromAcu <> "" Then
                                ValPromAcu = Split(ValPromAcu, ",")
                                If ValPromAcu(0) <> 0 Then
                                    Printer.CurrentX = 19.2
                                    Printer.Print ValPromAcu(0);
                                    Printer.CurrentX = 20
                                    Printer.Print Format(ValPromAcu(1) / ValPromAcu(0), "#.00")
                                Else
                                    Printer.Print "";
                                    Printer.Print ""
                                End If
                            Else
                                Printer.Print ""
                            End If
                        Else
                            Printer.Print ""
                        End If
                        Printer.FontBold = False
                        GoTo encontrar
                    End If
                Wend
encontrar:
                Close #NAR
                NAR = NAR - 1
            Else
                notas.FA = 0
            End If
            If Dir(Ruta & Combo4.Text & argra.num_area & lwe & ".dsp") <> "" Then
                NAR = FreeFile
                VV = 0
                Open Ruta & Combo4.Text & argra.num_area & lwe & ".dsp" For Random As #NAR Len = Len(notas_desemp)
                While Not EOF(NAR)
                    VV = VV + 1
                    Get #NAR, VV, notas_desemp
                    If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                        OkDes = True
                        If OkObs = False Then
                            NAR = FreeFile
                            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                            Get #NAR, argra.num_area, mate
                            Close #NAR
                            NAR = NAR - 1
                            Call SaltaLinea
                            Printer.CurrentX = 0.5
                            Printer.FontSize = 9
                            Printer.FontBold = True
                            Impr_adicional = RTrim(mate.nom) & "  (I.H:" & argra.ih & " - AUSENCIAS:" & notas.FA & ")"
                            Printer.Print Impr_adicional;
                            Printer.FontSize = 8
                            Printer.CurrentX = 14.2
                            If Trim(mate.nom) <> "CONVIVENCIA ESCOLAR" Then
                                Printer.Print "Valoración promedio acumulada:";
                                'Call Definitiva_Imp
                                ValPromAcu = Definitiva_Imp
                                If ValPromAcu <> "" Then
                                    ValPromAcu = Split(ValPromAcu, ",")
                                    If ValPromAcu(0) <> 0 Then
                                        Printer.CurrentX = 19.2
                                        Printer.Print ValPromAcu(0);
                                        Printer.CurrentX = 20
                                        Printer.Print Format(ValPromAcu(1) / ValPromAcu(0), "#.00")
                                    Else
                                        Printer.Print "";
                                        Printer.Print ""
                                    End If
                                Else
                                    Printer.Print ""
                                End If
                            Else
                                Printer.Print ""
                            End If
                            Printer.FontBold = False
                        End If
                        GoTo encontrar2
                    End If
                Wend
encontrar2:
                Close #NAR
                NAR = NAR - 1
            End If
            If OkDes = True Then
                CROA = 0
                Cont_Lgr = 0
                NAR = FreeFile
                Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                While Not EOF(NAR)
                    CROA = CROA + 1
                    Get #NAR, CROA, logru
                    If Trim(logru.indicador) = "L" Then
                        Cont_Lgr = Cont_Lgr + 1
                    End If
                Wend
                Close #NAR
                For I = 1 To Cont_Lgr
                    NAR = FreeFile
                    Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
                    For h = 1 To 14
                        Get #NAR, h, confdesemp
                        If Trim(argra.grado) = Trim(confdesemp.grado) Then
                            Exit For
                        End If
                    Next h
                    Close #NAR
    '                'VALIDA RANGOS POR DESEMPEÑO
                     If notas_desemp.porcentaje(I) <> 0 Then
                        If notas_desemp.porcentaje(I) <= confdesemp.rango(3) Then
                            Cont_Baj = Cont_Baj + 1
                        End If
                        If (notas_desemp.porcentaje(I) > confdesemp.rango(3)) And (notas_desemp.porcentaje(I) <= confdesemp.rango(2)) Then
                            Cont_Bas = Cont_Bas + 1
                        End If
    
                        If (notas_desemp.porcentaje(I) > confdesemp.rango(2)) And (notas_desemp.porcentaje(I) <= confdesemp.rango(1)) Then
                            Cont_Alt = Cont_Alt + 1
                        End If
                        If notas_desemp.porcentaje(I) > confdesemp.rango(1) Then
                            Cont_Sup = Cont_Sup + 1
                        End If
                    End If
                    '****IMPRIMIR LOGROS ******
                    NAR = FreeFile
                    Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                    Get #NAR, notas_desemp.logro(I), logru
                    If notas_desemp.porcentaje(I) <> 0 Then
                        NAR = FreeFile
                        Open Ruta & fl & seri & argra.num_area & lwe & ".ptj" For Random As #NAR Len = Len(porcent_manual)
                        Get #NAR, I, porcent_manual
                        Close #NAR
                        NAR = NAR - 1
                        Call SaltaLinea
                        Printer.FontSize = 8
                        Printer.CurrentX = 19.2
                        Printer.Print porcent_manual.porcent_logro;
                        Printer.CurrentX = 20.2
                        Printer.Print notas_desemp.porcentaje(I);
                        '*** SE VERIFICA EL TAMAÑO DEL LOGRO PARA EL SALTO DE LINEA ***
                        XMax = Printer.TextWidth(Trim(logru.indicador) & " - " & Trim(logru.observ))
                        If XMax > 16 Then
                            CortaObs (Trim(logru.indicador) & " - " & Trim(logru.observ))
                        Else
                            'Call SaltaLinea
                            Printer.CurrentX = 0.5
                            Printer.FontSize = 8
                            Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                        End If
                    End If
                    Close #NAR
                    NAR = NAR - 1
                Next I
            End If
            '****IMPRIMIR OBSERVACIONES ******
            If OkObs = True Then
                For I = 1 To 10
                    If notas.area(I) <> 0 Then
                        NAR = FreeFile
                        Open Ruta & fl & seri & argra.num_area & lwe & ".lgr" For Random As #NAR Len = Len(logru)
                        Get #NAR, notas.area(I), logru
                        If Trim(logru.indicador) <> "L" Then
                            '*** SE VERIFICA EL TAMAÑO DEL LOGRO PARA EL SALTO DE LINEA ***
                            XMax = Printer.TextWidth(Trim(logru.indicador) & " - " & Trim(logru.observ))
                            If XMax > 16 Then
                                CortaObs (Trim(logru.indicador) & " - " & Trim(logru.observ))
                            Else
                                Call SaltaLinea
                                Printer.CurrentX = 0.5
                                Printer.FontSize = 8
                                Printer.Print Trim(logru.indicador) & " - " & Trim(logru.observ)
                            End If
                            ValiSalto = True
                        End If
                        Close #NAR
                        NAR = NAR - 1
                    End If
                Next I
            End If
            If OkDes = True Or OkObs = True Then
                Printer.Print ""
            End If
        Printer.Line (0.5, Printer.CurrentY)-(21, Printer.CurrentY)
        End If
    Wend
    Close #NAR
    ' ********* IMPRIMIR PIE *********
    'Printer.Line (15.7, 5)-(15.7, Printer.CurrentY)
    Printer.Line (18.9, 4.5)-(18.9, Printer.CurrentY)
    Printer.Line (19.9, 4.5)-(19.9, Printer.CurrentY)
    Printer.Line (21, 4.5)-(21, Printer.CurrentY)
    
    ' VERIFICAR TAMAÑO NECESARIO PARA IMPRIMIR PIE
    'If Printer.CurrentY > 19.5 Then
    If Printer.CurrentY > 24.5 Then
       Printer.NewPage
       Call Encabezado
    End If
    Printer.Print ""
    'Printer.Print ""
    Open Ruta & "leyenda.edu" For Input As #NAR
    Input #NAR, leye.ly1, leye.ly2, leye.ly3, leye.ly4, leye.ly5, leye.ly6, leye.ly7, leye.ly8
    Close #NAR
    Printer.Font.Size = 9
    If Trim(leye.ly1) <> "" Then
        Printer.CurrentX = 0.5
        Printer.Print leye.ly1
    End If
    If Trim(leye.ly2) <> "" Then
        Printer.CurrentX = 0.5
        Printer.Print leye.ly2
    End If
    If Trim(leye.ly3) <> "" Then
        Printer.CurrentX = 0.5
        Printer.Print leye.ly3
    End If
    If Trim(leye.ly4) <> "" Then
        Printer.CurrentX = 0.5
        Printer.Print leye.ly4
    End If
    If Trim(leye.ly5) <> "" Then
        Printer.CurrentX = 0.5
        Printer.Print leye.ly5
    End If
    If Trim(leye.ly6) <> "" Then
        Printer.CurrentX = 0.5
        Printer.Print leye.ly6
    End If
    Printer.Print ""
    GuardaY = Printer.CurrentY
    'IMPRIMIR CUADRO EXPLICATIVO DE DESEMPEÑOS
    Printer.FontSize = 8
    Printer.FontBold = True
    Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
    Printer.CurrentX = 1.7
    Printer.Print "DESEMPEÑOS";
    Printer.CurrentX = 5.3
    Printer.Print "PORCENTAJES";
    Printer.CurrentX = 8
    Printer.Print "TOTAL LOGROS"
    Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
    Printer.FontBold = False
    Printer.CurrentX = 0.7
    Printer.Print confdesemp.desemp(1) & "(Desempeño Superior)";
    Printer.CurrentX = 5.5
    Printer.Print confdesemp.rango(1) + 1 & "% - 100%";
    Printer.CurrentX = 8.8
    'Printer.Print confdesemp.recupera(1)
    'Printer.Print ""
    Printer.Print Cont_Sup
    Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
    Printer.CurrentX = 0.7
    Printer.Print confdesemp.desemp(2) & "(Desempeño Alto)";
    Printer.CurrentX = 5.5
    Printer.Print confdesemp.rango(2) + 1 & "% - " & confdesemp.rango(1) & "%";
    Printer.CurrentX = 8.8
    'Printer.Print confdesemp.recupera(2)
    'Printer.Print ""
    Printer.Print Cont_Alt
    Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
    Printer.CurrentX = 0.7
    Printer.Print confdesemp.desemp(3) & "(Desempeño Básico)";
    Printer.CurrentX = 5.5
    Printer.Print confdesemp.rango(3) + 1 & "% - " & confdesemp.rango(2) & "%";
    Printer.CurrentX = 8.8
    'Printer.Print confdesemp.recupera(3)
    'Printer.Print ""
    Printer.Print Cont_Bas
    Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
    Printer.CurrentX = 0.7
    Printer.Print confdesemp.desemp(4);
    If Trim(confdesemp.desemp(4)) = "*LEP" Then
        Printer.Print " (Logro en Proceso)";
    Else
        Printer.Print " (Desempeño Bajo)";
    End If
    Printer.CurrentX = 5.5
    Printer.Print "  0% - " & confdesemp.rango(3) & "%";
    Printer.CurrentX = 8.8
    'Printer.Print confdesemp.recupera(4)
    'Printer.Print ""
    Printer.Print Cont_Baj
    Printer.Line (0.5, Printer.CurrentY)-(10.6, Printer.CurrentY)
    'Trazar líneas verticales
    Printer.Line (0.5, GuardaY)-(0.5, Printer.CurrentY)
    Printer.Line (5, GuardaY)-(5, Printer.CurrentY)
    Printer.Line (7.8, GuardaY)-(7.8, Printer.CurrentY)
    Printer.Line (10.6, GuardaY)-(10.6, Printer.CurrentY)
    Printer.FontSize = 9
    If Trim(leye.ly7) <> "" Then
        Printer.CurrentX = 0.5
        Printer.Print leye.ly7
    End If
    If Trim(leye.ly8) <> "" Then
        Printer.CurrentX = 0.5
        Printer.Print leye.ly8
    End If
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Line (3, Printer.CurrentY)-(7.7, Printer.CurrentY)
    Printer.Line (13, Printer.CurrentY)-(18, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.CurrentX = 3.5
    Printer.Print ini.Rector;
    Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
    Get #NAR, SP, profe
    Close #NAR
    Printer.CurrentX = 15.6 - ((Len(RTrim(profe.nombres) & " " & RTrim(profe.apellidos)) / 4.8) / 2)
    Printer.Print RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
    Printer.CurrentX = 4.5
    Printer.Print vini.VRector;
    Printer.CurrentX = 14
    Printer.Print vini.VDirector
    Screen.MousePointer = 0
    Printer.EndDoc
    Printer.PaperSize = 1
    End If

Else

' **********  IMPRESIÓN DEL REPORTE FINAL  ****************
'**********************************************************
    lwe = 4
    cona = 0
    RESP = MsgBox("DESEA IMPRIMIR EL REPORTE FINAL DE " & Frame1.Caption & "?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR REPORTE")
    If RESP = vbYes Then
        Screen.MousePointer = 11
        Call Encabezado
        MateriaX = 0
        VeriNivela = False
        NivelaTXT = ""
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            cona = cona + 1
            Get #NAR, cona, argra
            If RTrim(argra.nom_grup) = Combo4.Text Then
                NAR = FreeFile
                Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
                Get #NAR, argra.num_area, mate
                Close #NAR
                Open Ruta & "conf_desemp.edu" For Random As #NAR Len = Len(confdesemp)
                For h = 1 To 14
                    Get #NAR, h, confdesemp
                    If Trim(argra.grado) = Trim(confdesemp.grado) Then
                        Exit For
                    End If
                Next h
                Close #NAR
                NAR = NAR - 1
                
                '****CONVIVENCIA ESCOLAR(27) SE DISCRIMINA EN DISCIPLINA Y CONDUCTA******
                If (mate.num <> 27) Then
                    Printer.CurrentX = 1
                    Printer.FontSize = 9
                    Printer.FontBold = True
                    Printer.Print RTrim(mate.nom);
                    Printer.FontBold = False
                    Printer.CurrentX = 16
                    Printer.Print argra.ih;
                    Printer.CurrentX = 17
                    ValPromAcu = Definitiva_Imp
                    If ValPromAcu <> "" Then
                        ValPromAcu = Split(ValPromAcu, ",")
                        If ValPromAcu(0) <> 0 Then
                            'MATI20.Col = 1
                            'MATI20.Row = MATI20.Rows - 1
                            'MATI20.CellFontBold = True
                            'MATI20.TextMatrix(MATI20.Rows - 1, 1) = ValPromAcu(0)
                            Printer.Print Format(ValPromAcu(1) / ValPromAcu(0), "#.00");
                            'MATI20.Col = 2
                            'VALIDA RANGOS POR DESEMPEÑO
                            Printer.CurrentX = 18.5
                            If (ValPromAcu(1) / ValPromAcu(0)) <> "" Then
                                If (ValPromAcu(1) / ValPromAcu(0)) <= confdesemp.rango(3) Then
                                    'MATI20.CellFontBold = True
                                    MateriaX = MateriaX + 1
                                    
                                    '**********SE VERIFICA SI LA MATERIA FUE NIVELADA*********
                                    If Dir(Ruta & Combo4.Text & argra.num_area & "5.dsp") <> "" Then
                                        'OBTENER ARREGLO CON PORCENTAJES DE DESEMPEÑO
                                        NAR = FreeFile
                                        VV = 0
                                        Open Ruta & Combo4.Text & argra.num_area & "5.dsp" For Random As #NAR Len = Len(notas_desemp)
                                        While Not EOF(NAR)
                                            VV = VV + 1
                                            Get #NAR, VV, notas_desemp
                                            If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                                                If (notas_desemp.porcentaje(1) <> 0) Or (notas_desemp.porcentaje(2) <> 0) Then
                                                    VeriNivela = True
                                                    For I = 1 To 2
                                                        If (notas_desemp.porcentaje(I) <> 0) Then
                                                            NivelaTXT = NivelaTXT + RTrim(mate.nom) & "$" & notas_desemp.porcentaje(I)
                                                            'OBTENER CODIGO DE OBSERVACIÓN
                                                            If Dir(Ruta & Combo4.Text & argra.num_area & "5.obs") <> "" Then
                                                                NAR = FreeFile
                                                                VV2 = 0
                                                                Open Ruta & Combo4.Text & argra.num_area & "5.obs" For Random As #NAR Len = Len(notas)
                                                                While Not EOF(NAR)
                                                                    VV2 = VV2 + 1
                                                                    Get #NAR, VV2, notas
                                                                    If Val(notas.num_carnet) = Val(alugru.num_carnet) Then
                                                                        'For I = 1 To 2
                                                                            If notas.area(I) <> 0 Then
                                                                                NAR = FreeFile
                                                                                Open Ruta & fl & seri & argra.num_area & "5.lgr" For Random As #NAR Len = Len(logru)
                                                                                Get #NAR, notas.area(I), logru
                                                                                    'If (notas_desemp.porcentaje(I) <> 0) Then
                                                                                        NivelaTXT = NivelaTXT + "$" & RTrim(logru.observ)
                                                                                    'End If
                                                                                Close #NAR
                                                                                NAR = NAR - 1
                                                                            Else
                                                                                'NivelaTXT = NivelaTXT + "%"
                                                                            End If
                                                                        'Next I
                                                                        GoTo PorcentEncontrado22
                                                                    'Else
                                                                        'NivelaTXT = NivelaTXT + "%"
                                                                    End If
                                                                Wend
PorcentEncontrado22:
                                                                
                                                                Close #NAR
                                                                NAR = NAR - 1
                                                            Else
                                                                'NivelaTXT = NivelaTXT + "%"
                                                            End If
                                                            NivelaTXT = NivelaTXT + "%"
                                                        End If
                                                    Next I
                                                    If (notas_desemp.porcentaje(1) >= 70) Then
                                                        MateriaX = MateriaX - 1
                                                    Else
                                                        If (notas_desemp.porcentaje(1) <> 0) Then
                                                            MateriaX = MateriaX + 1
                                                        End If
                                                    End If
                                                    If (notas_desemp.porcentaje(2) >= 70) Then
                                                        MateriaX = MateriaX - 2
                                                    Else
                                                        If (notas_desemp.porcentaje(2) <> 0) Then
                                                            MateriaX = MateriaX + 2
                                                        End If
                                                    End If
                                                End If
                                                GoTo PorcentEncontrado11
                                            End If
                                        Wend
PorcentEncontrado11:
                                        Close #NAR
                                        NAR = NAR - 1
                                        
                                        
                                    End If
                                    
                                    Printer.Print "BAJO"
                                    Printer.CurrentX = 1
                                    Printer.FontSize = 8
                                    Printer.Print comdpe.bajo
                                    'Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                End If
                                If ((ValPromAcu(1) / ValPromAcu(0)) > confdesemp.rango(3)) And ((ValPromAcu(1) / ValPromAcu(0)) <= confdesemp.rango(2)) Then
                                    Printer.Print "BÁSICO"
                                    Printer.CurrentX = 1
                                    Printer.FontSize = 8
                                    Printer.Print comdpe.basico
                                    'Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                End If
                        
                                If ((ValPromAcu(1) / ValPromAcu(0)) > confdesemp.rango(2)) And ((ValPromAcu(1) / ValPromAcu(0)) <= confdesemp.rango(1)) Then
                                    Printer.Print "ALTO"
                                    Printer.CurrentX = 1
                                    Printer.FontSize = 8
                                    Printer.Print comdpe.alto
                                    'Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                End If
                                If (ValPromAcu(1) / ValPromAcu(0)) > confdesemp.rango(1) Then
                                    Printer.Print "SUPERIOR"
                                    Printer.CurrentX = 1
                                    Printer.FontSize = 8
                                    Printer.Print comdpe.superior
                                    'Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                End If
                                Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                            End If
                        Else
                            Printer.Print ""
                            'MATI20.TextMatrix(MATI20.Rows - 1, 2) = ""
                        End If
                    End If
                Else
                '****** SE IMPRIME DISCIPLINA Y CONDUCTA******
                    If Dir(Ruta & Combo4.Text & argra.num_area & "4.dsp") <> "" Then
                        NAR = FreeFile
                        VV5 = 0
                        Open Ruta & Combo4.Text & argra.num_area & "4.dsp" For Random As #NAR Len = Len(notas_desemp)
                        While Not EOF(NAR)
                            VV5 = VV5 + 1
                            Get #NAR, VV5, notas_desemp
                            If Val(notas_desemp.num_carnet) = Val(alugru.num_carnet) Then
                                If (Dir(Ruta & fl & seri & argra.num_area & "4.lgr") <> "") Then
                                    NAR = FreeFile
                                    Open Ruta & fl & seri & argra.num_area & "4.lgr" For Random As #NAR Len = Len(logru)
                                    For TT = 1 To 2
                                        Get #NAR, notas_desemp.logro(TT), logru
                                        If (notas_desemp.porcentaje(TT) <> 0) Then
                                            Printer.CurrentX = 1
                                            Printer.FontSize = 9
                                            Printer.FontBold = True
                                            Printer.Print Format(RTrim(logru.observ), ">");
                                            Printer.FontBold = False
                                            Printer.CurrentX = 17
                                            Printer.Print Format(notas_desemp.porcentaje(TT), "#.00");
                                            Printer.CurrentX = 18.5
                                            'RANGO DE DESEMPEÑOS
                                            If notas_desemp.porcentaje(TT) <= confdesemp.rango(3) Then
                                                Printer.Print "BAJO"
                                            End If
                                            If (notas_desemp.porcentaje(TT) > confdesemp.rango(3)) And (notas_desemp.porcentaje(TT) <= confdesemp.rango(2)) Then
                                                Printer.Print "BÁSICO"
                                            End If
                                    
                                            If (notas_desemp.porcentaje(TT) > confdesemp.rango(2)) And (notas_desemp.porcentaje(TT) <= confdesemp.rango(1)) Then
                                                Printer.Print "ALTO"
                                            End If
                                            If notas_desemp.porcentaje(TT) > confdesemp.rango(1) Then
                                                Printer.Print "SUPERIOR"
                                            End If
                                            Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
                                        End If
                                    Next TT
                                    Close #NAR
                                    NAR = NAR - 1
                                End If
                            End If
                        Wend
                        Close #NAR
                        NAR = NAR - 1
                    End If
                End If
            End If
        Wend
        Close #NAR
        Printer.Line (15.7, 5)-(15.7, Printer.CurrentY)
        Printer.Line (16.8, 5)-(16.8, Printer.CurrentY)
        Printer.Line (18.1, 5)-(18.1, Printer.CurrentY)
        Printer.Print ""
        Printer.FontSize = 9
        Printer.CurrentX = 1
        Printer.FontBold = True
        ' Abrir parámetros de promoción
        If (Dir(Ruta & "rangpro.txt") <> "") And (Dir(Ruta & "promovido.txt") <> "") Then
            Open Ruta & "rangpro.txt" For Input As #NAR
            Input #NAR, rus, fis
            Close #NAR
            Open Ruta & "promovido.txt" For Input As #NAR
            Input #NAR, SAPO2, SAPO3, SAPO4
            Close #NAR
        End If
        If MateriaX <= rus Then
            Printer.Print RTrim(SAPO2)
        End If
        If (MateriaX > rus) And (MateriaX <= fis) Then
            Printer.Print RTrim(SAPO3)
        End If
        If MateriaX > fis Then
            Printer.Print RTrim(SAPO4)
        End If
        Printer.FontBold = False
        Printer.Print ""
        If VeriNivela = True Then
            Printer.CurrentX = 1
            Printer.FontBold = True
            Printer.Print "NIVELACIÓN"
            Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
            Printer.CurrentX = 1
            Printer.Print "ASIGNATURA";
            Printer.CurrentX = 11
            Printer.Print "VAL";
            Printer.CurrentX = 12.5
            Printer.Print "DESEMPEÑO";
            Printer.CurrentX = 15
            Printer.Print "FECHA / ACTA"
            Printer.Line (1, Printer.CurrentY)-(21, Printer.CurrentY)
            'Printer.Print NivelaTXT
            Printer.FontBold = False
            Cortar = Split(NivelaTXT, "%")
            For Recorrer = 0 To UBound(Cortar)
                'Printer.CurrentX = 1
                Cortar2 = Split(Cortar(Recorrer), "$")
                For Recorrer2 = 0 To UBound(Cortar2)
                    If Recorrer2 = 0 Then
                        Printer.CurrentX = 1
                    End If
                    If Recorrer2 = 1 Then
                        Printer.CurrentX = 12.5
                        If (Cortar2(1)) <= confdesemp.rango(3) Then
                            Printer.Print "BAJO";
                        End If
                        If ((Cortar2(1)) > confdesemp.rango(3)) And ((Cortar2(1)) <= confdesemp.rango(2)) Then
                            Printer.Print "BÁSICO";
                        End If
                        If ((Cortar2(1)) > confdesemp.rango(2)) And ((Cortar2(1)) <= confdesemp.rango(1)) Then
                            Printer.Print "ALTO";
                        End If
                        If (Cortar2(1)) > confdesemp.rango(1) Then
                            Printer.Print "SUPERIOR";
                        End If
                        Printer.CurrentX = 11
                    End If
                    If Recorrer2 = 2 Then
                        Printer.CurrentX = 15
                    End If
                    If Recorrer2 <> 1 Then
                        Printer.Print Cortar2(Recorrer2);
                    Else
                        Printer.Print Format(Cortar2(Recorrer2), "#.00");
                    End If
                Next Recorrer2
                Printer.Print ""
            Next Recorrer
            Printer.Line (1, Printer.CurrentY - 0.2)-(21, Printer.CurrentY - 0.2)
            Printer.Print ""
        End If
        Printer.CurrentX = 1
        Printer.FontBold = True
        Printer.Print "OBSERVACIONES:"
        Printer.FontBold = False
        Printer.Print ""
        J = 0
        If (Dir(Ruta & "lrf" & Combo4.Text & ".lrf") <> "") And (Dir(Ruta & "orf" & Combo4.Text & ".orf") <> "") Then
            cona = 0
            Open Ruta & "lrf" & Combo4.Text & ".lrf" For Random As #NAR Len = Len(leyfin)
            While Not EOF(NAR)
                cona = cona + 1
                Get #NAR, cona, leyfin
                If Val(leyfin.num_carnet) = Val(alugru.num_carnet) Then
                    NAR = FreeFile
                    Open Ruta & "orf" & Combo4.Text & ".orf" For Random As #NAR Len = Len(obsfin)
                    For I = 1 To 5
                        If leyfin.fnob(I) <> 0 Then
                            Get #NAR, leyfin.fnob(I), obsfin
                            Printer.CurrentX = 1
                            
                            XMax = Printer.TextWidth(Trim(obsfin))
                            If XMax > 17 Then
                                CortaObs (Trim(obsfin))
                            Else
                                Printer.Print Trim(obsfin)
                            End If
                        Else
                            J = J + 1
                        End If
                    Next I
                    Close #NAR
                    Close #(NAR - 1)
                    NAR = NAR - 1
                    GoTo saobfi
                End If
            Wend
            Close #NAR
        Else
            J = 5
        End If
saobfi:
        If J <> 0 Then
            For I = 1 To J
                Printer.Print ""
            Next I
        End If
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        'Printer.Print ""
        'Firma del rector y director de grupo
        Printer.Line (3, Printer.CurrentY)-(7.7, Printer.CurrentY)
        Printer.Line (13, Printer.CurrentY)-(18, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 0.1
        Printer.CurrentX = 3.5
        Printer.Print ini.Rector;
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, SP, profe
        Close #NAR
        Printer.CurrentX = 15.6 - ((Len(RTrim(profe.nombres) & " " & RTrim(profe.apellidos)) / 4.8) / 2)
        Printer.Print RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        Printer.CurrentX = 4.5
        Printer.Print vini.VRector;
        Printer.CurrentX = 14
        Printer.Print vini.VDirector
        
        Screen.MousePointer = 0
        Printer.EndDoc
        Printer.PaperSize = 1
    End If
End If

End Sub

Private Sub Command3_Click()
If Dir(Ruta & Combo4.Text & ".gru") = "" Then
    MsgBox "NO EXISTE ESTE GRUPO", 48, "Imprimir Reportes"
    Combo4.SetFocus
    Exit Sub
End If
NAR = FreeFile
CONT = 0
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = Combo4.Text Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
        SP = RTrim(icur.director)
        NumAlias = CONT
    End If
    CONT = CONT + 1
Wend
Close #NAR
CONT = 0
VerAlias = ""
If Dir(Ruta & "aliasgrupos.edu") <> "" Then
    Open Ruta & "aliasgrupos.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, aliasg
        If CONT = NumAlias Then
            VerAlias = aliasg
        End If
        CONT = CONT + 1
    Wend
    Close #NAR
End If

Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, SP, profe
Close #NAR
PERI = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
seri = Left(RE22, 3)
If JOJI = "UNICA" Then
    fl = "1"
End If
If JOJI = "MAÑANA" Then
    fl = "2"
End If
If JOJI = "TARDE" Then
    fl = "3"
End If
If JOJI = "NOCHE" Then
    fl = "4"
End If
If Combo3.Text = "PRIMERO" Then
    lwe = 1
End If
If Combo3.Text = "SEGUNDO" Then
    lwe = 2
End If
If Combo3.Text = "TERCERO" Then
    lwe = 3
End If
If Combo3.Text = "CUARTO" Then
    lwe = 4
End If
If Combo3.Text = "FINAL" Then
    lwe = 5
End If
ret = 0
Open Ruta & Combo4.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
Open Ruta & "inicial.edu" For Input As #NAR
Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
Close #NAR
Open Ruta & "comentadesemp.edu" For Input As #NAR
Input #NAR, comdpe.bajo, comdpe.basico, comdpe.alto, comdpe.superior
Close #NAR
' *****LLAMAR FORMULARIO DE IMPRESION POR INTERVALO *****
Impr_ReportAcadem.Frame2.Caption = Combo4.Text
Impr_ReportAcadem.Label3.Caption = VerAlias
Impr_ReportAcadem.Show 1
End Sub

Private Sub Command4_Click()
Dim LongNom As Single
If MATI20.Rows = 1 Then
    MsgBox "NO HAY INFORMACION PARA COPIAR", 48, "COPIAR"
    Exit Sub
End If
Clipboard.Clear
cop = ""
Printer.ScaleMode = 7
'Copia el nombre del estudiante
cop = Left(Frame1.Caption, Len(Frame1.Caption) - 5) & vbCrLf & vbCrLf
cop = cop + "ÁREAS" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "I.H." & Chr(9) & "VAL" & Chr(9) & "VALORACIÓN" & Chr(9) & vbCrLf
For I = 1 To (MATI20.Rows - 1)
        MATI20.Row = I
        MATI20.Col = 0
        If (Right(MATI20.Text, 1) = ")") Then
            nom = Left(MATI20.Text, Len(MATI20.Text) - 9)
            ihcpy = Left(Right(MATI20.Text, 2), 1)
        Else
            nom = MATI20.Text
            ihcpy = ""
        End If
        LongNom = Printer.TextWidth(nom)
        While LongNom < 8
            nom = nom & Chr(9)
            LongNom = Printer.TextWidth(nom)
        Wend
        'VALIDA RANGOS POR DESEMPEÑO
        If MATI20.TextMatrix(I, 1) <> "" Then
            If MATI20.TextMatrix(I, 1) <= confdesemp.rango(3) Then
                valcpy = Trim(confdesemp.desemp(4))
                valcpy2 = "BAJO"
            End If
            If (MATI20.TextMatrix(I, 1) > confdesemp.rango(3)) And (MATI20.TextMatrix(I, 1) <= confdesemp.rango(2)) Then
                valcpy = Trim(confdesemp.desemp(3))
                valcpy2 = "BÁSICO"
            End If
        
            If (MATI20.TextMatrix(I, 1) > confdesemp.rango(2)) And (MATI20.TextMatrix(I, 1) <= confdesemp.rango(1)) Then
                valcpy = Trim(confdesemp.desemp(2))
                valcpy2 = "ALTO"
            End If
            If MATI20.TextMatrix(I, 1) > confdesemp.rango(1) Then
                valcpy = Trim(confdesemp.desemp(1))
                valcpy2 = "SUPERIOR"
            End If
        Else
            valcpy = ""
            valcpy2 = ""
        End If
        cop = cop + nom & ihcpy & Chr(9) & valcpy & Chr(9) & valcpy2 & vbCrLf
Next I
'Copia la promoción
cop = cop & vbCrLf & vbCrLf & Label12.Caption
Clipboard.SetText cop
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta por pantalla o impresión de boletines."
End Sub

Private Sub MATI20_Click()
If MATI20.Row > 0 Then
   MATI20.Col = 0
   MATI20.ToolTipText = Left(RTrim(MATI20.Text), 200)
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
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
Private Sub Form_Load()
'Dim icur As inforcur
MATI20.Row = 0
MATI20.Col = 0
MATI20.ColWidth(0) = 8000
MATI20.CellFontBold = True
MATI20.CellForeColor = RGB(255, 255, 255)
MATI20.CellBackColor = RGB(0, 0, 150)
MATI20.Text = "                M A T E R I A"
MATI20.Col = 1
MATI20.ColWidth(1) = 600
MATI20.CellFontBold = True
MATI20.CellForeColor = RGB(255, 255, 255)
MATI20.CellBackColor = RGB(0, 0, 150)
MATI20.Text = "  %"
MATI20.Col = 2
MATI20.ColWidth(2) = 650
MATI20.CellFontBold = True
MATI20.CellForeColor = RGB(255, 255, 255)
MATI20.CellBackColor = RGB(0, 0, 150)
MATI20.Text = " CAL"

If Dir(Ruta & "head1.jpg") <> "" Then
    Image1.Picture = LoadPicture(Ruta & "head1.jpg")
    Label1.Caption = "Imagen izquierda"
Else
    Label1.Caption = "No hay imagen"
End If

If Dir(Ruta & "head2.jpg") <> "" Then
    Image2.Picture = LoadPicture(Ruta & "head2.jpg")
    Label2.Caption = "Imagen derecha"
Else
    Label2.Caption = "No hay imagen"
End If


If Dir(Ruta & "infcur.edu") <> "" Then
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo4.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Combo4.Text = Combo4.List(0)
Else
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
End If
Text2.MaxLength = 3
'Option3.Value = True
End Sub
