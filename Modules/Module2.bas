Attribute VB_Name = "Module2"
Public NUMERO As Integer
Public WS_NUMDOC As Long
Public WS_NUMSER As Integer
Public exito As Boolean
Public ACV_TEXTO As String
' CAMTEX
Public LOC_BRUTO As Currency
Public LK_CAJERO As String * 1

Public CAMPOS1 As Integer
Dim indice1, INDICE2, INDICE3, INDICE4, CONTADOR As Integer
Public nn   As Integer
Public m_ind As Integer
Public tab_avanza(150) As Integer
Public LK_QUEDA As String * 1

Public Sub cancela_todo()
Dim h As Integer
nn = 2
m_ind = 0
Do Until val(tra_llave(nn)) = 0 Or nn = 62 '

m_ind = m_ind + 1

indice = TABLA_TAG(tra_llave(nn))


If TypeOf FORMGEN.Controls(indice) Is MSFlexGrid Then
   If FORMGEN.Controls(indice).Visible = True Then
      FORMGEN.Controls(indice).Clear
      FORMGEN.Controls(indice).rows = 3
   End If
End If
'If TypeOf FORMGEN.Controls(indice) Is TextBox Then
'   If FORMGEN.Controls(indice).Visible = True Then
'      FORMGEN.Controls(indice).text = ""
'   End If
'End If

'If TypeOf FORMGEN.Controls(indice) Is ComboBox Then
'  FORMGEN.Controls(indice).ListIndex = 0
'   FORMGEN.Controls(indice).Name
'End If


nn = nn + 4
Loop

'FORMGEN.fragas.Visible = False
FORMGEN.i_importe.Text = ""
If FORMGEN.Frame4.Visible = True Then
   FORMGEN.i_subtotal.Text = ""
   FORMGEN.i_gastos.Text = ""
   FORMGEN.i_descto.Text = ""
   FORMGEN.i_impto.Text = ""
   
   FORMGEN.i_neto.Text = ""
   FORMGEN.i_cant.Text = ""
   FORMGEN.i_flete.Text = ""
   FORMGEN.i_dias.Text = ""
   FORMGEN.i_TEXTONCRE.Text = ""

   If LK_CODTRA = 2401 And LK_QUEDA = "A" Then
   On Error GoTo TT
   fila = FORMGEN.grid_fac.rows - 1
    For h = 2 To fila
     If Trim(FORMGEN.grid_fac.TextMatrix(h, 1)) = "" Then GoTo PAS
     If val(FORMGEN.grid_fac.TextMatrix(h, 38)) = 0 Then
        FORMGEN.grid_fac.Row = h
        FORMGEN.grid_fac.RemoveItem h
        h = h - 1
       fila = fila - 1
      End If
PAS:
    Next h
   Else
     FORMGEN.grid_fac.Clear
     FORMGEN.grid_fac.rows = 3
   End If
TT:
   fila = 0
   FORMGEN.i_cant.Locked = False
   FORMGEN.TEXTOVAR.Visible = False
End If
FORMGEN.textovarl.Visible = False
If LK_CODTRA = 2401 And LK_QUEDA = "A" Then
Else
 FORMGEN.i_nomCLI.Caption = ""
 FORMGEN.i_nomven.Caption = ""
 FORMGEN.i_nomart.Caption = ""
 FORMGEN.i_codcli.Text = ""
 FORMGEN.i_codven.Text = ""
 FORMGEN.i_numguia.Text = ""
 FORMGEN.i_serguia.Text = ""
End If

FORMGEN.i_nomban.Caption = ""
FORMGEN.i_nomban2.Caption = ""
FORMGEN.i_limcre.Text = ""
FORMGEN.i_turno.Text = ""
If LK_CODTRA = 2748 Then
 FORMGEN.i_numfac_c.Text = ""
 FORMGEN.i_numser_c.Text = ""
 FORMGEN.imp_gas1.Text = ""
 FORMGEN.imp_gas2.Text = ""
 FORMGEN.cta_gas1.Text = ""
 FORMGEN.cta_gas2.Text = ""
End If
If LK_CODTRA = 2401 Then
 FORMGEN.i_codban.Text = ""
 FORMGEN.i_chenum.Text = ""
 FORMGEN.i_dias.Text = ""
End If
If LK_CODTRA = 1401 Then
  FORMGEN.i_situacion.ListIndex = -1
  FORMGEN.t_doc.Text = ""
  FORMGEN.i_num_ini.Text = ""
  FORMGEN.i_num_lote.Text = ""
End If




'If FORMGEN.i_codcli.Visible = True Then FORMGEN.gridl.Visible = False
FORMGEN.textovar_canje.Visible = False
'FORMGEN.grid_che.Visible = False

'FORMGEN.i_fbg.Enabled = True

FORMGEN.i_dias.Enabled = True
FORMGEN.i_fecha_vcto.Enabled = True
FORMGEN.i_dias.Enabled = True
FORMGEN.i_cant.Locked = False
FORMGEN.textovarl.Visible = False
FORMGEN.textovar_canje.Visible = False
FORMGEN.gridl.Visible = False
FORMGEN.grid_canje.Visible = False
FORMGEN.fracli.Visible = False
FORMGEN.t_nombre.Text = ""
FORMGEN.t_doc.Text = ""
FORMGEN.t_direc.Text = ""
FORMGEN.i_condi.ListIndex = -1


'agregado para chofer mic
FORMGEN.i_codcho.Text = ""
FORMGEN.i_nomcho.Caption = ""


Inicio_de_Todo
If SUT_LLAVE.EOF = False Then
   pasa_def
End If
FILAX = 0
fin:
End Sub
Public Sub Inicio_de_Todo()
'CAMBIARLO A OTRO SITIO...
'LK_RELACION_STOCK = par_llave!PAR_RELACION_STOCK
LK_RELACION_STOCK = 2.5

FORMGEN.i_fecha_vcto.Text = LK_FECHA_DIA
FORMGEN.i_camal.Visible = False
WS_LETRA_ACTIVA = False
WS_BRUTO = 0
SUB_CANT = 0
SUB_JABAS = 0
SUB_UNIDAD = 0
WS_DESCTO2 = 0
fila = 0


End Sub
Public Sub avanza_campo()
On Error GoTo SALE
If tab_avanza(NUMERO) = 0 Then
   FORMGEN.grabar.SetFocus
   GoTo fin
End If
If tab_avanza(NUMERO) = 100 Then
      FORMGEN.grid_fac.Row = 2
      FORMGEN.grid_fac.COL = 1
      FORMGEN.grid_fac.SetFocus
      GoTo fin
End If

indice = tab_avanza(NUMERO)
NUMERO = TABLA_TAG(indice)

If FORMGEN.Controls(TABLA_TAG(indice)).Enabled = False Then
   GoTo fin
End If
'If TypeOf FORMGEN.Controls(TABLA_TAG(indice)) Is Frame Then
'Else

'MsgBox FORMGEN.Controls(TABLA_TAG(indice)).Name
If indice = 36 Then
   If LK_EMP = "3AA" Then
     FORMGEN.i_fecha_compra.SetFocus
     FORMGEN.i_fecha_compra.SelStart = 0
     FORMGEN.i_fecha_compra.SelLength = 2
     Exit Sub
   Else
     Azul2 FORMGEN.i_fecha_compra, FORMGEN.i_fecha_compra
   End If
ElseIf indice = 200 Then
   'If LK_EMP = "3AA" Then
     FORMGEN.i_fecha_can.SetFocus
     FORMGEN.i_fecha_can.SelStart = 0
     FORMGEN.i_fecha_can.SelLength = 2
     Exit Sub
   'Else
   '  Azul2 FORMGEN.i_fecha_can, FORMGEN.i_fecha_can
   'End If
ElseIf indice = 109 Then
   'If LK_EMP = "3AA" Then
     FORMGEN.i_fecha_oper.SetFocus
     FORMGEN.i_fecha_oper.SelStart = 0
     FORMGEN.i_fecha_oper.SelLength = 2
  '   Exit Sub
  ' Else
  '   Azul2 FORMGEN.i_fecha_can, FORMGEN.i_fecha_can
ElseIf indice = 43 Then
   Azul FORMGEN.i_dias, FORMGEN.i_dias
Else
  FORMGEN.Controls(TABLA_TAG(indice)).SetFocus
End If
'End If
If TypeOf FORMGEN.Controls(TABLA_TAG(indice)) Is ComboBox Then
   SendKeys "%{DOWN}"
End If
fin:
SALE:
End Sub
Public Sub Azul(VART As Variant, varc As TextBox)
If varc.Enabled = True And varc.Visible = True Then
   varc.SetFocus
   varc.SelStart = 0
   varc.SelLength = Len(VART)
End If
End Sub
Public Sub Azul2(VART As Variant, varc As MaskEdBox)
If varc.Enabled = True And varc.Visible = True Then
   varc.SetFocus
   varc.SelStart = 0
   varc.SelLength = Len(VART)
End If
End Sub
Public Sub Azul3(VART As Variant, varc As RichTextBox)
If varc.Enabled And varc.Visible Then
   varc.SetFocus
   varc.SelStart = 0
   varc.SelLength = Len(VART)
End If
End Sub

Public Function ALINEA(VAR As String) As String
Dim TEMP As String * 15
Dim N1 As Integer
Dim N2 As Integer
N1 = InStr(1, VAR, " ") - 1
N2 = Len(VAR) - N1
VAR = String(N2, "    ") + Left(VAR, N1)
ALINEA = VAR
End Function

Public Function MIRA_DERECHOS(NUMERO As Integer) As String
Dim i As Integer
If NUMERO = 0 Then
   GoTo SALIR
End If

For i = 1 To 10
    If NUMERO = lk_GRUPOS(i) Then
       MIRA_DERECHOS = "S"
       Exit For
    End If
Next i
    
    
SALIR:
End Function

Public Function MIRA_DERECHOS2(NUMERO As Integer) As String
Dim i As Integer
If NUMERO = 0 Then
   GoTo SALIR
End If

Do Until lk_CODTRAS(i) = 0

 For i = 1 To 10
    If NUMERO = lk_TRANSA(i) Then
       MIRA_DERECHOS2 = "S"
       Exit For
    End If
 Next i
Loop
SALIR:
End Function


Public Sub Repo_Grid()
Dim count As Integer
Dim cant As String * 10
Dim nom As String * 25
Dim cod As String * 10
Dim SP As String * 10
Dim Prec As String * 10
Dim wBruto As Double
Dim wGastos As Double
Dim wDescto As Double
Dim wImpto As Double

Print #1, String(70, "-")
Print #1, "Codigo"; Spc(10); "Descripci�n"; Spc(15); "Cantidad"; Spc(4); "Precio"; Spc(4); "Total"
Print #1, String(70, "-")
For count = 1 To WS_ULT_FILA
    FORMGEN.grid_fac.Row = count
    FORMGEN.grid_fac.COL = 2
    nom = FORMGEN.grid_fac.Text
    FORMGEN.grid_fac.COL = 8
    cant = val(FORMGEN.grid_fac.Text)
    FORMGEN.grid_fac.COL = 1
    cod = Trim(FORMGEN.grid_fac.Text)
    cod = ALINEA(cod)
    FORMGEN.grid_fac.COL = 6
    SP = Format(FORMGEN.grid_fac.Text, "###,###,##0.00")
    FORMGEN.grid_fac.COL = 5
    Prec = Format(FORMGEN.grid_fac.Text, "###,###,##0.00")
    SP = ALINEA(SP)
    cant = ALINEA(cant)
    Prec = ALINEA(Prec)
    'Escribe al Archivo
    Print #1, cod; Spc(2); nom; Spc(1); cant; Spc(1); Prec; Spc(1); SP
 Next count
wBruto = Format(FORMGEN.i_subtotal, "###,###,##0.00")
wGastos = Format(val(FORMGEN.i_gastos), "###,###,##0.00")
wDescto = Format(val(FORMGEN.i_descto), "###,###,##0.00")
wImpto = Format(val(FORMGEN.i_impto), "###,###,##0.00")
Print #1,
Print #1,
Print #1, String(70, "-")
Print #1, "Bruto: "; wBruto; Spc(4); "Gasto.: "; wGastos; Spc(3); "Descto.: "; wDescto; Spc(3); "Impto.  : "; wImpto
Print #1,
'Print #1, Spc(48); "TOTAL GEN. :  "; Format(PUB_SUBTOTAL, "###,###,##0.00")
'Print #1, String(70, "=") FALTA LOS INTERESES


End Sub




Public Sub REPO_GRIDK(Ncolumnas As Integer, GridR As MSFlexGrid)
Dim SW As Integer
Dim i As Integer
Dim count As Integer
Dim Cc(12) As String * 15
Dim Nro As Integer
SW = 0
If Ncolumnas < 3 And Ncolumnas > 14 Then
   MsgBox "Columna debe ser: 3 > < 14 ", 48
   Exit Sub
End If
Nro = 0
'Open "Reporte.RTF" For Output As #1
For count = 0 To GridR.rows - 1 'Inicio del for Principal
  GridR.Row = count
  For i = 1 To Ncolumnas
   GridR.COL = i
   Cc(i) = GridR.Text
 '  If CC(I) < "0" Or CC(I) > "9" Then
   If IsNumeric(Cc(i)) Or count = 0 Or i = 3 Then
      Cc(i) = ALINEA(Cc(i))
    End If
  ' End If
  Next i
  If Cc(0) = Space(15) And Cc(1) = Space(15) And Cc(2) = Space(15) Then
   If count <> 0 Then
        Nro = 1
   Else
        Nro = 0
   End If
  Else
     Nro = 0
  End If
  If Nro = 1 Then ' Fin de liena llena
     Exit For
  End If
If SW = 0 Then
  Print #1, Spc(6); String(90, "-")
  SW = 1
End If
Select Case Ncolumnas + 1
Case 4
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3)
Case 5
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4)
Case 6
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4); Spc(1); Cc(5)
Case 7
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4); Spc(1); Cc(5); Spc(1); Cc(6)
Case 8
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4); Spc(1); Cc(5); Spc(1); Cc(6); Spc(1); Cc(7)
Case 9
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4); Spc(1); Cc(5); Spc(1); Cc(6); Spc(1); Cc(7); Spc(1); Cc(8)
Case 10
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4); Spc(1); Cc(5); Spc(1); Cc(6); Spc(1); Cc(7); Spc(1); Cc(8); Spc(1); Cc(9)
Case 11
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4); Spc(1); Cc(5); Spc(1); Cc(6); Spc(1); Cc(7); Spc(1); Cc(8); Spc(1); Cc(9); Spc(1); Cc(10)
Case 12
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4); Spc(1); Cc(5); Spc(1); Cc(6); Spc(1); Cc(7); Spc(1); Cc(8); Spc(1); Cc(9); Spc(1); Cc(10); Spc(1); Cc(11)
Case 13
  Print #1, Cc(0); Spc(1); Cc(1); Spc(1); Cc(2); Spc(1); Cc(3); Spc(1); Cc(4); Spc(1); Cc(5); Spc(1); Cc(6); Spc(1); Cc(7); Spc(1); Cc(8); Spc(1); Cc(9); Spc(1); Cc(10); Spc(1); Cc(11); Spc(1); Cc(12)
End Select
SALIRr:
If SW = 1 Then
  Print #1, Spc(6); String(90, "-")
  SW = 2
End If

Next count ' Fin del for principal

Print #1,
Print #1,
Print #1, Spc(6); String(90, "-")
'Print #1, Spc(48); "TOTAL GEN. :  "; Format(I, "###,###,##0.00")
'Close #1
End Sub


Public Function Nulo_Valor0(Optional valor) As Variant
If IsNull(valor) = True Or valor = "" Then
   Nulo_Valor0 = 0
Else
   Nulo_Valor0 = valor
End If

End Function
Public Function Nulo_Valors(Optional valor) As Variant
If IsNull(valor) = True Then
   Nulo_Valors = ""
Else
   Nulo_Valors = valor
End If

End Function

Public Function NULO_DATE(Optional FEC) As Date
If IsDate(FEC) Then
   NULO_DATE = CDate(FEC)
End If

End Function
Public Function redondea(valor As Variant) As Variant
redondea = Format(valor, "########0.00")
End Function
Public Sub SOLO_DECIMAL(wsTexto As TextBox, Optional wsKeyAscii)
Dim car
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
End Sub

Public Sub SOLO_DECIMAL_no_NEGATIVO(wsTexto As TextBox, Optional wsKeyAscii)
    Dim car
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    
    ' Validaci�n para el punto decimal
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    ' Validaci�n para otros caracteres (ahora sin permitir el "-")
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
      End If
    End If
End Sub

Public Sub SOLO_ENTERO(Optional tecla)

Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If car < "0" Or car > "9" Then
    If tecla <> 8 And tecla <> 13 Then
        tecla = 0
        Beep
    End If
End If
End Sub

Public Sub CAPTURA_DATOS()

Static UNICO As String
Dim Control As Object
Dim enlace As Integer


nn = 2
m_ind = 0
Do Until val(tra_llave(nn)) = 0 Or nn = 62
m_ind = m_ind + 1
ETIQUETAX(m_ind) = FORMGEN.LABELGEN(m_ind).Caption
NUMERO = TABLA_TAG(tra_llave(nn))
If TypeOf FORMGEN.Controls(NUMERO) Is label Then
    enlace = 1
 Else
    enlace = 0
 End If
 If TypeOf FORMGEN.Controls(NUMERO) Is TextBox Then
    enlace = 0
 Else
    enlace = -1
 End If
 If TypeOf FORMGEN.Controls(NUMERO) Is MSFlexGrid Then
    enlace = -1
 End If
 If TypeOf FORMGEN.Controls(NUMERO) Is label Then
    enlace = -1
 End If

If enlace > -1 Then
     If enlace = 0 Then
       TEXTOX(m_ind) = FORMGEN.Controls(NUMERO).Text
    Else
       NOMBREX(m_ind) = FORMGEN.Controls(NUMERO).Caption
    End If
End If

nn = nn + 4
Loop



End Sub
Public Function JALA_CTA(CODCTA As String) As String
SQ_OPER = 1
PUB_CUENTA = CODCTA
LEER_COM_LLAVE
If com_llave.EOF Then
 JALA_CTA = "Cta. No Definida en Contabilidad"
Else
 JALA_CTA = Trim(com_llave!com_descripcion)
End If
End Function

Public Function dDouble(ByVal valor As Variant) As Double
    If IsNumeric(valor) Then
        dDouble = CDbl(valor)
    Else
        dDouble = 0
    End If
End Function
