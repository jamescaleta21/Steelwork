VERSION 5.00
Begin VB.Form frmFacComandaSunat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta RUC a Sunat"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRuc 
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8415
      Begin VB.TextBox txtRazSoc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   6495
      End
      Begin VB.TextBox txtEst 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1755
         TabIndex        =   10
         Top             =   720
         Width           =   6495
      End
      Begin VB.TextBox txtCon 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1755
         TabIndex        =   6
         Top             =   1200
         Width           =   6495
      End
      Begin VB.TextBox txtDir 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1755
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   6495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RAZ�N SOCIAL:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   307
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   795
         TabIndex        =   9
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONDICI�N:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   8
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCI�N:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   525
         TabIndex        =   7
         Top             =   1680
         Width           =   1170
      End
   End
   Begin VB.Frame FraDni 
      Height          =   3255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   8415
      Begin VB.TextBox txtNombres 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   6495
      End
      Begin VB.TextBox txtPaterno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1440
         Width           =   6495
      End
      Begin VB.TextBox txtMaterno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1920
         Width           =   6495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRES:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   19
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP. PATERNO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AP. MATERNO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.TextBox txtNroDocumento 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7080
      TabIndex        =   2
      Top             =   4200
      Width           =   1350
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5640
      TabIndex        =   1
      Top             =   4200
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NRO DOCUMENTO:"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1635
   End
End
Attribute VB_Name = "frmFacComandaSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gNroDocumento As String
Public gEsRUC As Boolean
'Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function SetPixelV Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'
'Dim Rred
'Dim Ggreen
'Dim Bblue
'Dim Pixel
'Dim xDat As String
'Dim xRazSoc As String, xEst As String, xCon As String, xDir As String
'Dim xRazSocX As Long, xEstX As Long, xConX As Long, xDirX As Long
'Dim xRazSocY As Long, xEstY As Long, xConY As Long, xDirY As Long
Public gAcepta As Boolean
Public gRS, gDIR  As String

Private Sub cmdAceptar_Click()
If gEsRUC And Len(Trim(Me.txtRazSoc.Text)) = 0 Then
    MsgBox "No hay nada que actualizar.", vbInformation, Pub_Titulo
    Exit Sub
End If
If gEsRUC = False And Len(Trim(Me.txtNombres.Text)) = 0 And Len(Trim(Me.txtPaterno.Text)) = 0 And Len(Trim(Me.txtMaterno.Text)) = 0 Then
    MsgBox "No hay nada que actualizar.", vbInformation, Pub_Titulo
    Exit Sub
End If

If gEsRUC Then
gDIR = Trim(Me.txtDir.Text)
gRS = Trim(Me.txtRazSoc.Text)
Else
gDIR = ""
gRS = Trim(Me.txtNombres.Text) & " " & Trim(Me.txtPaterno.Text) & " " & Trim(Me.txtMaterno.Text)
End If
gAcepta = True
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub ConsultaDocumento(pNroDocto As String)
   On Error GoTo cCruc

    Dim p          As Object

    Dim Texto      As String, xTOk As String

    Dim cadena     As String, xvRUC As String

    Dim sInputJson As String

    gEsRUC = True

    MousePointer = vbHourglass
    Set httpURL = New WinHttp.WinHttpRequest
    
    If IsNumeric(pNroDocto) Then
        If Len(Trim(pNroDocto)) = 8 Then
            gEsRUC = False
        End If

        xvRUC = pNroDocto
    Else

        If Len(Trim(pNroDocto)) = 8 Then
            gEsRUC = False
        End If

        xvRUC = pNroDocto
    End If

    xTOk = Leer_Ini(App.Path & "\config.ini", "TOKEN", "")
    
    If gEsRUC Then
        cadena = "http://dniruc.apisperu.com/api/v1/ruc/" & xvRUC & "?token=" & xTOk
    Else
        cadena = "http://dniruc.apisperu.com/api/v1/dni/" & xvRUC & "?token=" & xTOk
    End If
    
    httpURL.Open "GET", cadena
    httpURL.Send
    
    Texto = httpURL.ResponseText

    'sInputJson = "{items:" & Texto & "}"

    Set p = JSON.parse(Texto)


    
    If Len(Trim(pNroDocto)) = 0 Then
        If IsNumeric(pNroDocto) Then
            If Len(Trim(pNroDocto)) = 11 Or Len(Trim(pNroDocto)) = 8 Then
                If Texto = "[]" Then
                    MousePointer = vbDefault
                    MsgBox ("No se obtuvo resultados")
                    Me.txtCon.Text = ""
                    Me.txtDir.Text = ""
                    Me.txtEst.Text = ""
                    Me.txtRazSoc.Text = ""
                    Me.txtPaterno.Text = ""
                    Me.txtNombres.Text = ""
                    Me.txtMaterno.Text = ""

                    Exit Sub

                End If

                If Len(Trim(Texto)) = 0 Then
                    MousePointer = vbDefault
                    MsgBox ("No se obtuvo resultados")
                    Me.txtCon.Text = ""
                    Me.txtDir.Text = ""
                    Me.txtEst.Text = ""
                    Me.txtRazSoc.Text = ""
                    Me.txtPaterno.Text = ""
                    Me.txtNombres.Text = ""
                    Me.txtMaterno.Text = ""

                    Exit Sub

                End If

                If gEsRUC Then
                    Me.txtDir.Text = p.Item("direccion")
                    Me.txtRazSoc.Text = p.Item("razonSocial")
                   Me.txtEst.Text = p.Item("estado")
            Me.txtCon.Text = p.Item("condicion")
                Else
                    
                    Me.txtDir.Text = ""
                    Me.txtEst.Text = ""
                    Me.txtCon.Text = ""
                   ' Me.txtDni.Text = p.Item("dni")
                   ' Me.txtRS.Text = " " & p.Item("apellidoPaterno") & " " & p.Item("apellidoMaterno")
                    Me.txtNombres.Text = p.Item("nombres")
                    Me.txtPaterno.Text = p.Item("apellidoPaterno")
                    Me.txtMaterno.Text = p.Item("apellidoMaterno")
                End If
    
               
            
            Else
                MsgBox "El ruc debe tener 11 caracteres", vbCritical, Pub_Titulo
            End If

        Else
            MsgBox "El ruc debe ser Numeros", vbCritical, Pub_Titulo
        End If

    Else

        If Texto = "[]" Then
            MousePointer = vbDefault
            MsgBox ("No se obtuvo resultados")
             Me.txtCon.Text = ""
                    Me.txtDir.Text = ""
                    Me.txtEst.Text = ""
                    Me.txtRazSoc.Text = ""
                    Me.txtPaterno.Text = ""
                    Me.txtNombres.Text = ""
                    Me.txtMaterno.Text = ""

            Exit Sub

        End If

        If Len(Trim(Texto)) = 0 Then
            MousePointer = vbDefault
            MsgBox ("No se obtuvo resultados")
            Me.txtCon.Text = ""
                    Me.txtDir.Text = ""
                    Me.txtEst.Text = ""
                    Me.txtRazSoc.Text = ""
                    Me.txtPaterno.Text = ""
                    Me.txtNombres.Text = ""
                    Me.txtMaterno.Text = ""

            Exit Sub

        End If
        
        If gEsRUC Then
            'Me.txtDni.Text = ""
            Me.txtDir.Text = p.Item("direccion")
            Me.txtRazSoc.Text = p.Item("razonSocial")
            Me.txtEst.Text = p.Item("estado")
            Me.txtCon.Text = p.Item("condicion")
            Me.txtPaterno.Text = ""
            Me.txtMaterno.Text = ""
            Me.txtNombres.Text = ""
            'Me.txtruc.Text = p.Item("ruc")
        Else
              Me.txtDir.Text = ""
                    Me.txtEst.Text = ""
                    Me.txtCon.Text = ""
                   ' Me.txtDni.Text = p.Item("dni")
                   ' Me.txtRS.Text = " " & p.Item("apellidoPaterno") & " " & p.Item("apellidoMaterno")
                    Me.txtNombres.Text = p.Item("nombres")
                    Me.txtPaterno.Text = p.Item("apellidoPaterno")
                    Me.txtMaterno.Text = p.Item("apellidoMaterno")
        End If
    

      
                
    End If
       
    MousePointer = vbDefault

    Exit Sub

cCruc:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

'Private Sub Command1_Click()
' PicCaptcha.AutoRedraw = True
'    PicCaptcha.AutoSize = True
'    PicCaptcha.ScaleMode = 3
'
'    'CONSULTA EL RUC
'    Dim hWeb As String
'
'    'Generando el Captcha desde la pagina de la Sunat http://www.sunat.gob.pe/cl-ti-itmrconsruc/captcha?accion=image
'    'Jalando Datos de la Pagina http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorRuc&nroRuc=?&search1=?&codigo=NTPY&tipdoc=1
'    'If CboDocumento.ListIndex = 0 Then
'    If Len(gRUC) <> 11 Then MsgBox "Por favor ingrese RUC de 11 Digitos", vbExclamation, "Atencion": Exit Sub
'    hWeb = "http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorRuc&nroRuc="
'    '  Else
'    '    If Len(txtnrodocumento.Text) <> 8 Then MsgBox "Por favor ingrese DNI de 8 Digitos", vbExclamation, "Atencion": txtnrodocumento.SetFocus: Exit Sub
'    '    hWeb = "http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias?accion=consPorTipdoc&nrodoc="
'    '  End If
'
'    Call Limpiar
'    Call DescargaCaptcha
'
'    If ConvertirImagenTexto = True Then
'        If Len(Trim(TxtCaptcha.Text)) = 0 Then
'            btnCon = True
'        Else
'            Call Descargar(hWeb & gRUC & "&codigo=" & TxtCaptcha.Text & "&tipdoc=1")
'
'            If OTROsunat(gRUC) = False Then
'                btnCon = True
'            End If
'        End If
'    End If
'End Sub

Private Sub Form_Load()
    gAcepta = False
    Me.txtNroDocumento.Text = gNroDocumento
    ConsultaDocumento Trim(gNroDocumento)
'Command1_Click
End Sub

Private Sub Limpiar()

    txtRazSoc.Text = ""
    txtEst.Text = ""
    txtCon.Text = ""
    txtDir.Text = ""
End Sub
'Function DescargaCaptcha()
'  On Error Resume Next
'  Call Descargar("http://www.sunat.gob.pe/cl-ti-itmrconsruc/captcha?accion=image")
'  PicCaptcha.Picture = LoadPicture(GetDirTemp & "\sunat.tmp")
'End Function

'Function ConvertirImagenTexto() As Boolean
'    On Error Resume Next
'    Dim ShellPath As String
'    Dim TEXTO As String
'    Dim x As Integer, y As Integer
'
'    TxtCaptcha.Text = ""
'
''    If chkCambiarResolucion.Value = 1 Then    ' Si deseamos convertir la imagen
'        For y = 0 To PicCaptcha.ScaleHeight - 1
'            For x = 0 To PicCaptcha.ScaleWidth - 1
'            Pixel = GetPixel(PicCaptcha.HDC, x, y)
'            GetRGB Pixel
'            Rred = 250 - Rred
'            Ggreen = 250 - Ggreen
'            Bblue = 250 - Bblue
'            SetPixelV PicCaptcha.HDC, x, y, RGB(Rred, Ggreen, Bblue)
'            Next
'            PicCaptcha.Refresh
'        Next
'        PicCaptcha.Refresh
''    End If
'
'    Call SavePicture(PicCaptcha.Image, GetDirTemp & "\sunat.tmp")   'Guardando la Imagen para convertir a Texto
'
'    ShellPath = GetShortDir(App.Path) & "\modulo.dll " & GetDirTemp & "\sunat.tmp " & GetDirTemp & "\output" & " -psm"
'
'    If ShellAndWait(ShellPath, vbMinimizedNoFocus) = True Then  ' Esperando a que el OCR Convierta el Texto
'
'        Open GetDirTemp & "\output.txt" For Input As #1   'Mostrando el texto Convertido
'          While Not EOF(1)
'              Line Input #1, TEXTO
'              TxtCaptcha.Text = UCase(Replace(TEXTO, Chr(13), ""))
'          Wend
'        Close #1
'
'        Kill GetDirTemp & "\output.txt"   'Borrando el texto generado
'        ConvertirImagenTexto = True
'    Else
''        MsgBox "La imagen no se pudo Convertir a Texto"
'      ConvertirImagenTexto = False
'    End If
'
'End Function

'Function OTROsunat(ByVal xNum As String) As Boolean
''On Error Resume Next
'  On Error GoTo EsteErr
'    Dim tmpVal As String
'    Dim xTabla() As String
'    Dim PosisionScript As Integer, PosisionScript1 As Integer
'
'
'        Call Limpiar
'
'        xDat = OpenTxt 'xWml.responseText
''        If Len(xDat) <= 635 Then
''            Call Habilitar(False)
''            MsgBox "El numero Ruc ingresado no existe en la Base de datos de la SUNAT", vbCritical, "Error"
''            txtnrodocumento.SetFocus
''            Exit Function
''        End If
'        xDat = Replace(xDat, vbCrLf, "")
'        xDat = Replace(xDat, "     ", " ")
'        xDat = Replace(xDat, "    ", " ")
'        xDat = Replace(xDat, "   ", " ")
'        xDat = Replace(xDat, "  ", " ")
'        xDat = Replace(xDat, "( ", "(")
'        xDat = Replace(xDat, " )", ")")
'        xDat = Replace(xDat, Chr(34), "'")
'        xDat = Replace(xDat, "<tr>", "")
'        xDat = Replace(xDat, "</td>", "")
'        xDat = Replace(xDat, "</tr>", "")
'
'        'If CboDocumento.ListIndex = 0 Then
'          If InStr(xDat, "La aplicaci&oacute;n ha retornado el siguiente problema") > 0 Then OTROsunat = True: MsgBox "A ocurrido un error" & vbCrLf & "- Verifique su conexion a Internet" & vbCrLf & "- Pagina de sunat en Mantenimiento", vbCritical, "Error": Exit Function
'          If InStr(xDat, "consultado no es v�lido") > 0 Then MsgBox "El n�mero de RUC " & gRUC & " consultado no es v�lido. Debe verificar el n�mero y volver a ingresar.", vbCritical, "Error": OTROsunat = True: Exit Function
''          If InStr(xDat, "El codigo ingresado es incorrecto") > 0 Then MsgBox "El n�mero de RUC " & txtnrodocumento.Text & " consultado no es v�lido. Debe verificar el n�mero y volver a ingresar.", vbCritical, "Error": txtnrodocumento.SetFocus: OTROsunat = True: Exit Function
''        Else
''          If InStr(xDat, "El Sistema RUC NO REGISTRA") > 0 Then MsgBox "El Sistema RUC NO REGISTRA un n�mero de RUC para el DNI n�mero " & txtnrodocumento.Text & " consultado.", vbCritical, "Error": txtnrodocumento.SetFocus: OTROsunat = True: Exit Function
''          PosisionScript = InStr(xDat, "<a href='javascript:sendNroRuc(")
''          If PosisionScript > 0 Then CboDocumento.ListIndex = 0: txtnrodocumento.Text = Mid$(xDat, PosisionScript + 31, 11): OTROsunat = True: btnCon = True: Exit Function Else MsgBox "El n�mero de DNI " & txtnrodocumento.Text & " consultado no tiene nigun RUC asociado", vbCritical, "Error": txtnrodocumento.SetFocus: OTROsunat = True: Exit Function
''        End If
'
'        'Call Habilitar(True)
'
'        PosisionScript = InStr(xDat, "}</script>")
'        If PosisionScript > 0 Then xDat = Mid(xDat, PosisionScript + 10)
'
'        PosisionScript = InStr(xDat, "<td class='bgn' colspan=1>Sistema de Emisi&oacute;n de Comprobante:")
'        If PosisionScript > 0 Then xDat = Mid(xDat, 1, PosisionScript)
'
'        PosisionScript = InStr(xDat, "<td class='bgn' colspan=1>Tipo Contribuyente:")
'        PosisionScript1 = InStr(xDat, "<td class='bgn' colspan=1>Estado del Contribuyente:")
''        MsgBox PosisionScript & "-" & PosisionScript1
'
'        If PosisionScript > 0 And PosisionScript1 > 0 Then tmpVal = Mid(xDat, PosisionScript, PosisionScript1 - PosisionScript)
'        xDat = Replace$(xDat, tmpVal, "")
''        If PosisionScript1 > 0 Then xDat = Mid(xDat, 1, PosisionScript1)      'xTabla(2)
'
'        xTabla = Split(xDat, "<td ")
'
'
'
'
'          xTabla(2) = Replace(xTabla(2), "class='bg' colspan=3>" & xNum & " - ", "")
'          xTabla(4) = Replace(xTabla(4), "class='bg' colspan=1>", "")
'          xTabla(7) = Replace(xTabla(7), "class='bg' colspan=3>", "")
'          xTabla(9) = Replace(Mid(xTabla(9), 1, InStr(xTabla(9), "<!--") - 1), "class='bg' colspan=3>", "")
'
'        xRazSoc = Trim(CStr(xTabla(2)))
'        xEst = Trim(CStr(xTabla(4)))
'        xCon = Trim(CStr(xTabla(7)))
'        xDir = Trim(CStr(xTabla(9)))
'
'        txtRazSoc.Text = xRazSoc
'        txtEst.Text = xEst
'        txtCon.Text = xCon
'        txtDir.Text = xDir
'
'        OTROsunat = True
'    Exit Function
'EsteErr:
'    OTROsunat = False
''    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Error"
'End Function

'Private Sub GetRGB(ByVal COL As String)
'    On Error Resume Next
'    Bblue = COL \ (256 ^ 2)
'    Ggreen = (COL - Bblue * 256 ^ 2) \ 256
'    Rred = (COL - Bblue * 256 ^ 2 - Ggreen * 256)
'End Sub
'Function OpenTxt() As String
'On Error Resume Next
'
''Open "d:\sunat.txt" For Input As #1
'Open GetDirTemp & "\sunat.tmp" For Input As #1
'
'Dim Linea As String, Total As String
'    Do Until EOF(1)
'    Line Input #1, Linea
'        Total = Total + Linea + vbCrLf
'    Loop
'Close #1
'    OpenTxt = Total
'
'    If Len(dir(GetDirTemp & "\sunat.tmp")) Then
'        Kill GetDirTemp & "\sunat.tmp"
'    End If
'
'End Function

Private Sub txtnrodocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        gRUC = Me.txtNroDocumento.Text
        
    End If

End Sub

