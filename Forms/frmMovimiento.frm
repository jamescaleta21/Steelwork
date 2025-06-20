VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmMovimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Movimientos"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11340
   Begin TabDlg.SSTab SSTMovimiento 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Movimiento"
      TabPicture(0)   =   "frmMovimiento.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "lblMovimientoID"
      Tab(0).Control(8)=   "DatUbicacionDestino"
      Tab(0).Control(9)=   "DatResponsableDestino"
      Tab(0).Control(10)=   "DatResponsableOrigen"
      Tab(0).Control(11)=   "DatActivo"
      Tab(0).Control(12)=   "dtpFechaMovimiento"
      Tab(0).Control(13)=   "txtObservacion"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMovimiento.frx":0EE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblUbicacion"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DatActivoSearch"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lvDetalle"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin MSComctlLib.ListView lvDetalle 
         Height          =   4575
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtObservacion 
         Height          =   1335
         Left            =   -71400
         TabIndex        =   7
         Tag             =   "X"
         Top             =   4800
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtpFechaMovimiento 
         Height          =   300
         Left            =   -71400
         TabIndex        =   3
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   189136897
         CurrentDate     =   45825
      End
      Begin MSDataListLib.DataCombo DatActivo 
         Height          =   315
         Left            =   -71400
         TabIndex        =   2
         Top             =   1680
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DatResponsableOrigen 
         Height          =   315
         Left            =   -71400
         TabIndex        =   4
         Top             =   2880
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DatResponsableDestino 
         Height          =   315
         Left            =   -71400
         TabIndex        =   5
         Top             =   3480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DatUbicacionDestino 
         Height          =   315
         Left            =   -71400
         TabIndex        =   6
         Top             =   4200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DatActivoSearch 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HISTORIAL DE MOVIMIENTOS"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   10815
      End
      Begin VB.Label lblUbicacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         TabIndex        =   20
         Tag             =   "X"
         Top             =   1080
         Width           =   8895
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación Actual:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1133
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   1110
         TabIndex        =   18
         Top             =   660
         Width           =   600
      End
      Begin VB.Label lblMovimientoID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -71400
         TabIndex        =   17
         Tag             =   "X"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación:"
         Height          =   195
         Left            =   -72615
         TabIndex        =   16
         Top             =   4800
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación Destino:"
         Height          =   195
         Left            =   -73065
         TabIndex        =   15
         Top             =   4260
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable Destino:"
         Height          =   195
         Left            =   -73320
         TabIndex        =   14
         Top             =   3540
         Width           =   1860
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable Origen:"
         Height          =   195
         Left            =   -73260
         TabIndex        =   13
         Top             =   2940
         Width           =   1785
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Movimiento:"
         Height          =   195
         Left            =   -73080
         TabIndex        =   12
         Top             =   2333
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   -72075
         TabIndex        =   11
         Top             =   1740
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimiento ID:"
         Height          =   195
         Left            =   -72795
         TabIndex        =   10
         Top             =   1133
         Width           =   1320
      End
   End
   Begin MSComctlLib.ImageList iMovimiento 
      Left            =   12840
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimiento.frx":0F02
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimiento.frx":129C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimiento.frx":1636
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimiento.frx":19D0
            Key             =   "transfer"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMovimiento 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   1164
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Guardar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DatActivoSearch_Change()
Me.lvDetalle.ListItems.Clear
    On Error GoTo xSearch

    If Me.DatActivoSearch.BoundText = -1 Then Exit Sub
    LimpiaParametros oCmdEjec
    oCmdEjec.Prepared = True
    oCmdEjec.CommandText = "[sw].[USP_MOVIMIENTO_SEARCH]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVOID", adInteger, adParamInput, , Me.DatActivoSearch.BoundText)
    
    Set oRSmain = oCmdEjec.Execute
    
    If Not oRSmain.EOF Then
        Me.lblUbicacion.Caption = oRSmain!ubicacion
        
        Dim orsT As ADODB.Recordset

        Set orsT = oRSmain.NextRecordset
        
        Dim itemX As Object
        
        Do While Not orsT.EOF
            Set itemX = Me.lvDetalle.ListItems.Add(, , IIf(IsNull(orsT!movimientoId), "", orsT!movimientoId), Me.iMovimiento.ListImages(4).key, Me.iMovimiento.ListImages(4).key)
            itemX.SubItems(1) = orsT!fechaMovimiento
            itemX.SubItems(2) = IIf(IsNull(orsT!responsableOrigen), "", orsT!responsableOrigen)
            itemX.SubItems(3) = IIf(IsNull(orsT!responsableDestino), "", orsT!responsableDestino)
            itemX.SubItems(4) = orsT!ubicacion
            itemX.SubItems(5) = orsT!tipoMovimiento
            itemX.SubItems(6) = orsT!cuRegistro
            orsT.MoveNext
        Loop

    End If
    
    Exit Sub
xSearch:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
Estado_Botones InicializarFormulario
DesactivarControles Me
Me.DatActivoSearch.Enabled = True
configuraLV
llenaComboActivo
Me.tbMovimiento.ImageList = Me.iMovimiento
Me.tbMovimiento.Buttons(1).Image = Me.iMovimiento.ListImages(1).index
Me.tbMovimiento.Buttons(2).Image = Me.iMovimiento.ListImages(2).index
Me.tbMovimiento.Buttons(3).Image = Me.iMovimiento.ListImages(3).index
End Sub

Private Sub Estado_Botones(val As Valores)

    Select Case val

        Case InicializarFormulario, grabar, cancelar
            Me.tbMovimiento.Buttons(1).Enabled = True
            Me.tbMovimiento.Buttons(2).Enabled = False
            Me.tbMovimiento.Buttons(3).Enabled = False
          
            Me.SSTMovimiento.tab = 1

        Case Nuevo
            Me.tbMovimiento.Buttons(1).Enabled = False
            Me.tbMovimiento.Buttons(2).Enabled = True
            Me.tbMovimiento.Buttons(3).Enabled = True
          
            Me.lvDetalle.Enabled = False
            Me.SSTMovimiento.tab = 0


    End Select

End Sub

Private Sub tbMovimiento_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            Me.dtpFechaMovimiento.Value = LK_FECHA_DIA
            
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[sw].[USP_MOVIMIENTO_DATOSCOMBOS]"
            oCmdEjec.Prepared = True
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            
            Set oRSmain = oCmdEjec.Execute
            
            Set Me.DatActivo.RowSource = oRSmain
            Me.DatActivo.BoundColumn = oRSmain.Fields(0).Name
            Me.DatActivo.ListField = oRSmain.Fields(1).Name
            Me.DatActivo.BoundText = -1
            
            Dim orsT As ADODB.Recordset

            Set orsT = oRSmain.NextRecordset
            Set Me.DatResponsableOrigen.RowSource = orsT
            Me.DatResponsableOrigen.BoundColumn = orsT.Fields(0).Name
            Me.DatResponsableOrigen.ListField = orsT.Fields(1).Name
            Me.DatResponsableOrigen.BoundText = -1
                
            Set Me.DatResponsableDestino.RowSource = orsT
            Me.DatResponsableDestino.BoundColumn = orsT.Fields(0).Name
            Me.DatResponsableDestino.ListField = orsT.Fields(1).Name
            Me.DatResponsableDestino.BoundText = -1
                
            Set orsT = oRSmain.NextRecordset
            Set Me.DatUbicacionDestino.RowSource = orsT
            Me.DatUbicacionDestino.BoundColumn = orsT.Fields(0).Name
            Me.DatUbicacionDestino.ListField = orsT.Fields(1).Name
            Me.DatUbicacionDestino.BoundText = -1
            
            Me.DatActivo.SetFocus
            
        Case 2 'Grabar
        
            If Me.DatActivo.BoundText = -1 Then
                MsgBox "Debe elegir el Activo.", vbInformation, Pub_Titulo
                Me.DatActivo.SetFocus
            ElseIf Me.DatResponsableOrigen.BoundText = -1 Then
                MsgBox "Debe elegir el Responsable Origen.", vbInformation, Pub_Titulo
                Me.DatResponsableOrigen.SetFocus
            ElseIf Me.DatResponsableDestino.BoundText = -1 Then
                MsgBox "Debe elegir el Responsable Destino.", vbInformation, Pub_Titulo
                Me.DatResponsableDestino.SetFocus
            ElseIf Me.DatUbicacionDestino.BoundText = -1 Then
                MsgBox "Debe elegir la Ubicación Destino.", vbInformation, Pub_Titulo
                Me.DatUbicacionDestino.SetFocus
            ElseIf Me.DatResponsableOrigen.BoundText = Me.DatResponsableDestino.BoundText Then
                MsgBox "El Responsable Origen no puede ser igual al Responsable Destino", vbInformation, Pub_Titulo
                Me.DatResponsableOrigen.SetFocus
            Else
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[sw].[USP_MOVIMIENTO_REGISTER]"
        
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVOID", adInteger, adParamInput, , Me.DatActivo.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAMOVIMIENTO", adDBTimeStamp, adParamInput, , Me.dtpFechaMovimiento.Value)
                 
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RESPONSABLEORIGENID", adInteger, adParamInput, , Me.DatResponsableOrigen.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RESPONSABLEDESTINOID", adInteger, adParamInput, , Me.DatResponsableDestino.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@UBICACIONDESTINOID", adInteger, adParamInput, , Me.DatUbicacionDestino.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                If Len(Trim(Me.txtObservacion.Text)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@OBSERVACION", adVarChar, adParamInput, 300, Trim(Me.txtObservacion.Text))
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        MsgBox oRSmain!Message, vbInformation, Pub_Titulo
                        Estado_Botones grabar
                    Else
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If
               
            End If

        Case 3 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvDetalle.Enabled = True
            Me.DatActivoSearch.Enabled = True
            Me.DatActivoSearch.SetFocus

    End Select

End Sub

Private Sub configuraLV()
Me.lvDetalle.Icons = Me.iMovimiento
Me.lvDetalle.SmallIcons = Me.iMovimiento
With Me.lvDetalle
.HideColumnHeaders = False
    .FullRowSelect = True
    .Gridlines = True
    .ColumnHeaders.Add , , "movimientoID"
    .ColumnHeaders.Add , , "Fecha Movimiento"
    .ColumnHeaders.Add , , "Responsable Origen"
    .ColumnHeaders.Add , , "Responsable Destino"
    .ColumnHeaders.Add , , "Ubicacion"
    .ColumnHeaders.Add , , "Movimiento"
    .ColumnHeaders.Add , , "Realizado Por"
End With
End Sub

Private Sub llenaComboActivo()

    On Error GoTo xLlena

    LimpiaParametros oCmdEjec
    oCmdEjec.Prepared = True
    oCmdEjec.CommandText = "[sw].[USP_MOVIMIENTO_LOAD_ACTIVO]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    Set oRSmain = oCmdEjec.Execute

    If Not oRSmain.EOF Then
        Set Me.DatActivoSearch.RowSource = oRSmain
        Me.DatActivoSearch.BoundColumn = oRSmain.Fields(0).Name
        Me.DatActivoSearch.ListField = oRSmain.Fields(1).Name
        Me.DatActivoSearch.BoundText = -1

    End If

    Exit Sub
xLlena:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub
