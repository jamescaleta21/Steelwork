VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmMainActivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Activos"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainActivo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10110
   Begin MSComctlLib.ImageList iActivo 
      Left            =   11280
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainActivo.frx":0ECA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainActivo.frx":1264
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainActivo.frx":15FE
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainActivo.frx":1998
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainActivo.frx":1D32
            Key             =   "desactive"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainActivo.frx":22CC
            Key             =   "active"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainActivo.frx":2866
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainActivo.frx":2C00
            Key             =   "responsable"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbActivo 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1164
      ButtonWidth     =   1879
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Guardar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Desactivar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Activar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Eliminar"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTActivo 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Activo"
      TabPicture(0)   =   "frmMainActivo.frx":319A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblActivoID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DatUbicacion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "DatResponsable"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "DatProveedor"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ComActivo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCodigo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDescripcion"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtNroSerie"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DatCategoria"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCostoInicial"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dtpFechaIngreso"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMainActivo.frx":31B6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "txtSearch"
      Tab(1).Control(2)=   "lvListado"
      Tab(1).ControlCount=   3
      Begin MSComCtl2.DTPicker dtpFechaIngreso 
         Height          =   300
         Left            =   3360
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   346619905
         CurrentDate     =   45825
      End
      Begin VB.TextBox txtCostoInicial 
         Height          =   300
         Left            =   6960
         TabIndex        =   7
         Tag             =   "X"
         Top             =   2280
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DatCategoria 
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Top             =   2760
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtNroSerie 
         Height          =   300
         Left            =   6240
         TabIndex        =   4
         Tag             =   "X"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   3360
         TabIndex        =   5
         Tag             =   "X"
         Top             =   1800
         Width           =   4815
      End
      Begin VB.TextBox txtCodigo 
         Height          =   300
         Left            =   3360
         TabIndex        =   3
         Tag             =   "X"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox ComActivo 
         Height          =   315
         ItemData        =   "frmMainActivo.frx":31D2
         Left            =   3360
         List            =   "frmMainActivo.frx":31DC
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvListado 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   2
         Top             =   960
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7858
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
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   -74040
         TabIndex        =   1
         Top             =   600
         Width           =   8775
      End
      Begin MSDataListLib.DataCombo DatProveedor 
         Height          =   315
         Left            =   3360
         TabIndex        =   9
         Top             =   3240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DatResponsable 
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         Top             =   3720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DatUbicacion 
         Height          =   315
         Left            =   3360
         TabIndex        =   11
         Top             =   4200
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ingreso:"
         Height          =   195
         Left            =   1905
         TabIndex        =   26
         Top             =   2333
         Width           =   1290
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
         Height          =   195
         Left            =   2310
         TabIndex        =   25
         Top             =   4260
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable:"
         Height          =   195
         Left            =   2040
         TabIndex        =   24
         Top             =   3780
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         Height          =   195
         Left            =   2235
         TabIndex        =   23
         Top             =   3300
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Costo Inicial:"
         Height          =   195
         Left            =   5760
         TabIndex        =   22
         Top             =   2333
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría:"
         Height          =   195
         Left            =   2280
         TabIndex        =   21
         Top             =   2820
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Serie:"
         Height          =   195
         Left            =   5280
         TabIndex        =   20
         Top             =   1373
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   2130
         TabIndex        =   19
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo ID:"
         Height          =   195
         Left            =   2325
         TabIndex        =   18
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label lblActivoID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   16
         Tag             =   "X"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   2595
         TabIndex        =   15
         Top             =   4740
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   13
         Top             =   600
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmMainActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Private Sub DatCategoria_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.DatProveedor.SetFocus
End Sub

Private Sub DatProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.DatResponsable.SetFocus
End Sub

Private Sub DatResponsable_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.DatUbicacion.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CentrarFormulario MDIForm1, Me
    DesactivarControles Me
    configurarLv
    activoSearch Me.txtSearch.Text
    Estado_Botones InicializarFormulario
    Me.dtpFechaIngreso.Value = LK_FECHA_DIA
    Me.tbActivo.ImageList = Me.iActivo
    Me.tbActivo.Buttons(1).Image = Me.iActivo.ListImages(1).index
    Me.tbActivo.Buttons(2).Image = Me.iActivo.ListImages(2).index
    Me.tbActivo.Buttons(3).Image = Me.iActivo.ListImages(3).index
    Me.tbActivo.Buttons(4).Image = Me.iActivo.ListImages(4).index
    Me.tbActivo.Buttons(5).Image = Me.iActivo.ListImages(5).index
    Me.tbActivo.Buttons(6).Image = Me.iActivo.ListImages(6).index
    Me.tbActivo.Buttons(7).Image = Me.iActivo.ListImages(7).index

End Sub

Private Sub configurarLv()
    Me.lvListado.Icons = Me.iActivo
    Me.lvListado.SmallIcons = Me.iActivo

    With Me.lvListado
        .HideColumnHeaders = False
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , , "Id", 1000
        .ColumnHeaders.Add , , "Código", 1500
        .ColumnHeaders.Add , , "Descripción", 3500
        .ColumnHeaders.Add , , "Costo Inicial", 1500
        .ColumnHeaders.Add , , "Activo", 1000

    End With

End Sub

Private Sub activoSearch(xdato As String)

    On Error GoTo xSearch

    Me.lvListado.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[sw].[USP_ACTIVO_SEARCH]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    If Len(Trim(xdato)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 100, xdato)
    
    Set oRSmain = oCmdEjec.Execute
    
    If Not oRSmain.EOF Then

        Dim itemX As Object

        Do While Not oRSmain.EOF
            Set itemX = Me.lvListado.ListItems.Add(, , oRSmain!activoId, Me.iActivo.ListImages(1).key, Me.iActivo.ListImages(8).key)
            itemX.SubItems(1) = oRSmain!codigoActivo
            itemX.SubItems(2) = oRSmain!descripcion
            itemX.SubItems(3) = oRSmain!costoinicial
            itemX.SubItems(4) = oRSmain!activo

            If oRSmain!activo = "NO" Then
                Me.lvListado.ListItems(itemX.index).ForeColor = vbRed
                Me.lvListado.ListItems(itemX.index).ListSubItems(1).ForeColor = vbRed
                Me.lvListado.ListItems(itemX.index).ListSubItems(2).ForeColor = vbRed
                Me.lvListado.ListItems(itemX.index).ListSubItems(3).ForeColor = vbRed
                Me.lvListado.ListItems(itemX.index).ListSubItems(4).ForeColor = vbRed

            End If

            oRSmain.MoveNext
        Loop

    End If

    Exit Sub
xSearch:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Estado_Botones(val As Valores)

    Select Case val

        Case InicializarFormulario, grabar, cancelar, Eliminar, Desactivar, Activar
            Me.tbActivo.Buttons(1).Enabled = True
            Me.tbActivo.Buttons(2).Enabled = False
            Me.tbActivo.Buttons(3).Enabled = False
            Me.tbActivo.Buttons(4).Enabled = False
            Me.tbActivo.Buttons(5).Enabled = False
            Me.tbActivo.Buttons(6).Enabled = False
            Me.tbActivo.Buttons(7).Enabled = False
            Me.SSTActivo.tab = 1

        Case Nuevo, Editar
            Me.tbActivo.Buttons(1).Enabled = False
            Me.tbActivo.Buttons(2).Enabled = True
            Me.tbActivo.Buttons(3).Enabled = False
            Me.tbActivo.Buttons(4).Enabled = True
            Me.tbActivo.Buttons(5).Enabled = False
            Me.tbActivo.Buttons(6).Enabled = False
            Me.tbActivo.Buttons(7).Enabled = False
            Me.lvListado.Enabled = False
            Me.txtSearch.Enabled = False
            Me.SSTActivo.tab = 0

        Case buscar
            Me.tbActivo.Buttons(1).Enabled = True
            Me.tbActivo.Buttons(2).Enabled = False
            Me.tbActivo.Buttons(3).Enabled = False
            Me.tbActivo.Buttons(4).Enabled = False
            Me.SSTActivo.tab = 1

        Case AntesDeActualizar
            Me.tbActivo.Buttons(1).Enabled = False
            Me.tbActivo.Buttons(2).Enabled = False
            Me.tbActivo.Buttons(3).Enabled = True
            Me.tbActivo.Buttons(4).Enabled = True

            If Me.ComActivo.ListIndex = 0 Then
                Me.tbActivo.Buttons(5).Enabled = True
                Me.tbActivo.Buttons(6).Enabled = False
            Else
                Me.tbActivo.Buttons(5).Enabled = False
                Me.tbActivo.Buttons(6).Enabled = True

            End If

            Me.tbActivo.Buttons(7).Enabled = True
            Me.SSTActivo.tab = 0

    End Select

End Sub

Sub Mandar_Datos()

    With Me.lvListado
        Me.lblActivoID.Caption = .SelectedItem.Text
        Me.txtCodigo.Text = .SelectedItem.SubItems(1)
        Me.txtDescripcion.Text = .SelectedItem.SubItems(2)
        Me.txtCostoInicial.Text = .SelectedItem.SubItems(3)

        If .SelectedItem.SubItems(3) = "SI" Then
            Me.ComActivo.ListIndex = 0
        Else
            Me.ComActivo.ListIndex = 1

        End If
    
        Estado_Botones AntesDeActualizar
        
        llenaCombos
        traerDatosActivo Me.lblActivoID.Caption

    End With

End Sub


Private Sub lvListado_DblClick()
Mandar_Datos
End Sub

Private Sub tbActivo_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            llenaCombos
            VNuevo = True
            Me.dtpFechaIngreso.Value = LK_FECHA_DIA
            Me.ComActivo.Enabled = False
            Me.ComActivo.ListIndex = 0
            Me.txtCodigo.SetFocus

        Case 2 'Guardar

            If Len(Trim(Me.txtCodigo.Text)) = 0 Then
                MsgBox "Debe ingresar el código", vbInformation, Pub_Titulo
                Me.txtCodigo.SetFocus
            ElseIf Len(Trim(Me.txtDescripcion.Text)) = 0 Then
                MsgBox "Debe ingresar la Descripción", vbInformation, Pub_Titulo
                Me.txtDescripcion.SetFocus
            ElseIf Len(Trim(Me.txtNroSerie.Text)) = 0 Then
                MsgBox "Debe ingresar el Número de Serie", vbCritical, Pub_Titulo
                Me.txtNroSerie.SetFocus
            ElseIf Len(Trim(Me.txtCostoInicial.Text)) = 0 Then
                MsgBox "Debe ingresar el Costo Inicial", vbCritical, Pub_Titulo
                Me.txtCostoInicial.SetFocus
            ElseIf val(Me.txtCostoInicial.Text) <= 0 Then
                MsgBox "Costo Inicial incorrecto", vbCritical, Pub_Titulo
                Me.txtCostoInicial.SelStart = 0
                Me.txtCostoInicial.SelLength = Len(Me.txtCostoInicial.Text)
                Me.txtCostoInicial.SetFocus
            ElseIf Me.DatCategoria.BoundText = -1 Then
                MsgBox "Debe elegir la Categoria.", vbInformation, Pub_Titulo
                Me.DatCategoria.SetFocus
            ElseIf Me.DatProveedor.BoundText = -1 Then
                MsgBox "Debe elegir el Proveedor.", vbInformation, Pub_Titulo
                Me.DatProveedor.SetFocus
            ElseIf Me.DatResponsable.BoundText = -1 Then
                MsgBox "Debe elegir el Responsable.", vbInformation, Pub_Titulo
                Me.DatResponsable.SetFocus
            ElseIf Me.DatUbicacion.BoundText = -1 Then
                MsgBox "Debe elegir la Ubicación.", vbInformation, Pub_Titulo
                Me.DatUbicacion.SetFocus
                
            Else
                LimpiaParametros oCmdEjec

                If VNuevo Then
                    oCmdEjec.CommandText = "[sw].[USP_ACTIVO_REGISTER]"
                Else
                    oCmdEjec.CommandText = "[sw].[USP_ACTIVO_UPDATE]"

                End If

                On Error GoTo grabar

                Dim Smensaje As String

                Dim vIDz     As Integer

                Smensaje = ""
                vIDz = 0

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                
                If Not VNuevo Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVOID", adInteger, adParamInput, , Me.lblActivoID.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGOACTIVO", adVarChar, adParamInput, 20, Trim(Me.txtCodigo.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCRIPCION", adVarChar, adParamInput, 100, Trim(Me.txtDescripcion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMEROSERIE", adVarChar, adParamInput, 50, Trim(Me.txtNroSerie.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CATEGORIAID", adInteger, adParamInput, , Me.DatCategoria.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PROVEEDORID", adInteger, adParamInput, , Me.DatProveedor.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COSTOINICIAL", adDouble, adParamInput, , Me.txtCostoInicial.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHAINGRESO", adDBTimeStamp, adParamInput, , Me.dtpFechaIngreso.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RESPONSABLEID", adInteger, adParamInput, , Me.DatResponsable.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@UBICACIONID", adInteger, adParamInput, , Me.DatUbicacion.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        DesactivarControles Me
                        Estado_Botones grabar
                        Me.lvListado.Enabled = True
                        Me.txtSearch.Enabled = True
                
                        activoSearch Me.txtSearch.Text
                    Else
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If

                Exit Sub

grabar:
                MsgBox Err.Description, vbInformation, Pub_Titulo

            End If

        Case 3 'Modificar
            VNuevo = False
            Estado_Botones Editar
            ActivarControles Me
            Me.ComActivo.Enabled = False
            Me.txtCodigo.SelStart = 0
            Me.txtCodigo.SelLength = Len(Me.txtCodigo.Text)
            Me.txtCodigo.SetFocus
            Me.txtSearch.Enabled = False

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvListado.Enabled = True
            Me.txtSearch.Enabled = True
            Me.txtSearch.SetFocus
            
        Case 5 'Desactivar
            
            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            
                On Error GoTo Desactiva
            
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[sw].[USP_ACTIVO_STATUS]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PROVEEDORID", adInteger, adParamInput, , Me.lblActivoID.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVO", adBoolean, adParamInput, , False)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        LimpiarControles Me
                        Estado_Botones Desactivar
                        Me.lvListado.Enabled = True
                        activoSearch Me.txtSearch.Text
                    Else
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If
            
                Exit Sub
            
Desactiva:
                MsgBox Err.Description, vbInformation, Pub_Titulo
            
            End If
            
        Case 6 'ACTIVAR
            
            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then

                On Error GoTo Activa
            
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[sw].[USP_ACTIVO_STATUS]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PROVEEDORID", adInteger, adParamInput, , Me.lblActivoID.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVO", adBoolean, adParamInput, , True)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        LimpiarControles Me
                        Estado_Botones Activar
                        Me.lvListado.Enabled = True
                        activoSearch Me.txtSearch.Text
                    Else
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If
            
                Exit Sub
            
Activa:
                MsgBox Err.Description, vbInformation, Pub_Titulo
            
            End If

        Case 7 'ELIMINAR

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            
                Dim vACT As String
            
                vACT = ""
            
                On Error GoTo Elimina
            
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[sw].[USP_ACTIVO_DELETE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PROVEEDORID", adInteger, adParamInput, , Me.lblActivoID.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
              
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        DesactivarControles Me
                        Estado_Botones Eliminar
                        Me.lvListado.Enabled = True
                        Me.txtSearch.Enabled = True
                
                        activoSearch Me.txtSearch.Text
                    Else
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If
            
                Exit Sub
            
Elimina:
                MsgBox Err.Description, vbInformation, Pub_Titulo
            
            End If

    End Select

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then
    Me.txtNroSerie.SelLength = 0
    Me.txtNroSerie.SelLength = Len(Me.txtNroSerie.Text)
    Me.txtNroSerie.SetFocus
End If
End Sub

Private Sub txtCostoInicial_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL_no_NEGATIVO Me.txtCostoInicial, KeyAscii
If KeyAscii = vbKeyReturn Then Me.DatCategoria.SetFocus

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then
    Me.dtpFechaIngreso.SetFocus
End If
End Sub

Private Sub txtNroSerie_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then
    Me.txtDescripcion.SelLength = 0
    Me.txtDescripcion.SelLength = Len(Me.txtDescripcion.Text)
    Me.txtDescripcion.SetFocus
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then activoSearch Me.txtSearch.Text
End Sub

Private Sub llenaCombos()

    On Error GoTo llenaCombos

    LimpiaParametros oCmdEjec
    oCmdEjec.Prepared = True
    oCmdEjec.CommandText = "[sw].[USP_ACTIVO_DATOSCOMBOS]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    
    Set oRSmain = oCmdEjec.Execute
    
    Set Me.DatCategoria.RowSource = oRSmain
    Me.DatCategoria.BoundColumn = oRSmain.Fields(0).Name
    Me.DatCategoria.ListField = oRSmain.Fields(1).Name
    Me.DatCategoria.BoundText = -1
    
    Dim orsTMP As ADODB.Recordset

    Set orsTMP = oRSmain.NextRecordset
    
    'Proveedor
    Set Me.DatProveedor.RowSource = orsTMP
    Me.DatProveedor.BoundColumn = orsTMP.Fields(0).Name
    Me.DatProveedor.ListField = orsTMP.Fields(1).Name
    Me.DatProveedor.BoundText = -1
    
    'Responsable
    Set orsTMP = oRSmain.NextRecordset
    Set Me.DatResponsable.RowSource = orsTMP
    Me.DatResponsable.BoundColumn = orsTMP.Fields(0).Name
    Me.DatResponsable.ListField = orsTMP.Fields(1).Name
    Me.DatResponsable.BoundText = -1
    
    'Ubicacion
    Set orsTMP = oRSmain.NextRecordset
    Set Me.DatUbicacion.RowSource = orsTMP
    Me.DatUbicacion.BoundColumn = orsTMP.Fields(0).Name
    Me.DatUbicacion.ListField = orsTMP.Fields(1).Name
    Me.DatUbicacion.BoundText = -1
    
    Exit Sub
llenaCombos:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub traerDatosActivo(cIDactivo As Integer)
On Error GoTo xDatos
    LimpiaParametros oCmdEjec
    oCmdEjec.Prepared = True
    oCmdEjec.CommandText = "[sw].[USP_ACTIVO_FILL]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@activoId", adInteger, adParamInput, , cIDactivo)
    
    Set oRSmain = oCmdEjec.Execute
    
    If Not oRSmain.EOF Then
        Me.txtNroSerie.Text = oRSmain!numeroSerie
        Me.DatCategoria.BoundText = oRSmain!categoriaId
        Me.DatProveedor.BoundText = oRSmain!proveedorId
        Me.DatUbicacion.BoundText = oRSmain!ubicacionId
        Me.DatResponsable.BoundText = oRSmain!responsableId
        Me.dtpFechaIngreso.Value = oRSmain!FECHAINGRESO
    End If
    
Exit Sub
xDatos:
MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

