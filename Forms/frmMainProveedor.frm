VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Proveedores"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainProveedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9480
   Begin MSComctlLib.Toolbar tbProveedor 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
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
   Begin TabDlg.SSTab SSTProveedor 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Proveedor"
      TabPicture(0)   =   "frmMainProveedor.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "lblProveedorId"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(4)=   "ComActivo"
      Tab(0).Control(5)=   "txtRazonSocial"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMainProveedor.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lvListado"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtRazonSocial 
         Height          =   300
         Left            =   -71175
         TabIndex        =   6
         Tag             =   "X"
         Top             =   2280
         Width           =   3975
      End
      Begin VB.ComboBox ComActivo 
         Height          =   315
         ItemData        =   "frmMainProveedor.frx":05C2
         Left            =   -71175
         List            =   "frmMainProveedor.frx":05CC
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3120
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvListado 
         Height          =   3855
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6800
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
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   8175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor ID:"
         Height          =   195
         Left            =   -72525
         TabIndex        =   10
         Top             =   1613
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   -72465
         TabIndex        =   9
         Top             =   2333
         Width           =   1170
      End
      Begin VB.Label lblProveedorId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -71175
         TabIndex        =   8
         Tag             =   "X"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   -71895
         TabIndex        =   7
         Top             =   3180
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList iProveedor 
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
            Picture         =   "frmMainProveedor.frx":05D8
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainProveedor.frx":0972
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainProveedor.frx":0D0C
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainProveedor.frx":10A6
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainProveedor.frx":1440
            Key             =   "desactive"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainProveedor.frx":19DA
            Key             =   "active"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainProveedor.frx":1F74
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainProveedor.frx":230E
            Key             =   "proveedor"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMainProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
DesactivarControles Me
configurarLv
proveedorSearch Me.txtSearch.Text
Estado_Botones InicializarFormulario
Me.tbProveedor.ImageList = Me.iProveedor
Me.tbProveedor.Buttons(1).Image = Me.iProveedor.ListImages(1).index
Me.tbProveedor.Buttons(2).Image = Me.iProveedor.ListImages(2).index
Me.tbProveedor.Buttons(3).Image = Me.iProveedor.ListImages(3).index
Me.tbProveedor.Buttons(4).Image = Me.iProveedor.ListImages(4).index
Me.tbProveedor.Buttons(5).Image = Me.iProveedor.ListImages(5).index
Me.tbProveedor.Buttons(6).Image = Me.iProveedor.ListImages(6).index
Me.tbProveedor.Buttons(7).Image = Me.iProveedor.ListImages(7).index
End Sub

Private Sub configurarLv()
Me.lvListado.Icons = Me.iProveedor
Me.lvListado.SmallIcons = Me.iProveedor

    With Me.lvListado
        .HideColumnHeaders = False
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , , "Id"
        .ColumnHeaders.Add , , "Razón Social", 5000
        .ColumnHeaders.Add , , "Activo"

    End With

End Sub

Private Sub proveedorSearch(xdato As String)

    On Error GoTo xSearch

    Me.lvListado.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[sw].[USP_PROVEEDOR_SEARCH]"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    If Len(Trim(xdato)) <> 0 Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 100, xdato)
    
    Set oRSmain = oCmdEjec.Execute
    
    If Not oRSmain.EOF Then

        Dim itemX As Object

        Do While Not oRSmain.EOF
            Set itemX = Me.lvListado.ListItems.Add(, , oRSmain!proveedorid, Me.iProveedor.ListImages(1).key, Me.iProveedor.ListImages(8).key)
            itemX.SubItems(1) = oRSmain!razonsocial
            itemX.SubItems(2) = oRSmain!activo

            If oRSmain!activo = "NO" Then
                Me.lvListado.ListItems(itemX.index).ForeColor = vbRed
                Me.lvListado.ListItems(itemX.index).ListSubItems(1).ForeColor = vbRed
                Me.lvListado.ListItems(itemX.index).ListSubItems(2).ForeColor = vbRed

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
            Me.tbProveedor.Buttons(1).Enabled = True
            Me.tbProveedor.Buttons(2).Enabled = False
            Me.tbProveedor.Buttons(3).Enabled = False
            Me.tbProveedor.Buttons(4).Enabled = False
            Me.tbProveedor.Buttons(5).Enabled = False
            Me.tbProveedor.Buttons(6).Enabled = False
            Me.tbProveedor.Buttons(7).Enabled = False
            Me.SSTProveedor.tab = 1

        Case Nuevo, Editar
            Me.tbProveedor.Buttons(1).Enabled = False
            Me.tbProveedor.Buttons(2).Enabled = True
            Me.tbProveedor.Buttons(3).Enabled = False
            Me.tbProveedor.Buttons(4).Enabled = True
            Me.tbProveedor.Buttons(5).Enabled = False
            Me.tbProveedor.Buttons(6).Enabled = False
            Me.tbProveedor.Buttons(7).Enabled = False
            Me.lvListado.Enabled = False
            Me.txtSearch.Enabled = False
            Me.SSTProveedor.tab = 0

        Case buscar
            Me.tbProveedor.Buttons(1).Enabled = True
            Me.tbProveedor.Buttons(2).Enabled = False
            Me.tbProveedor.Buttons(3).Enabled = False
            Me.tbProveedor.Buttons(4).Enabled = False
            Me.SSTProveedor.tab = 1

        Case AntesDeActualizar
            Me.tbProveedor.Buttons(1).Enabled = False
            Me.tbProveedor.Buttons(2).Enabled = False
            Me.tbProveedor.Buttons(3).Enabled = True
            Me.tbProveedor.Buttons(4).Enabled = True

            If Me.ComActivo.ListIndex = 0 Then
                Me.tbProveedor.Buttons(5).Enabled = True
                Me.tbProveedor.Buttons(6).Enabled = False
            Else
                Me.tbProveedor.Buttons(5).Enabled = False
                Me.tbProveedor.Buttons(6).Enabled = True

            End If
Me.tbProveedor.Buttons(7).Enabled = True
            Me.SSTProveedor.tab = 0

    End Select

End Sub

Sub Mandar_Datos()

    With Me.lvListado
        Me.lblProveedorId.Caption = .SelectedItem.Text
        Me.txtRazonSocial.Text = .SelectedItem.SubItems(1)

        If .SelectedItem.SubItems(2) = "SI" Then
            Me.ComActivo.ListIndex = 0
            Else
            Me.ComActivo.ListIndex = 1
            End If
    
        Estado_Botones AntesDeActualizar

    End With

End Sub

Private Sub lvListado_DblClick()
Mandar_Datos
End Sub

Private Sub tbProveedor_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            VNuevo = True
            Me.ComActivo.Enabled = False
            Me.ComActivo.ListIndex = 0
            Me.txtRazonSocial.SetFocus

        Case 2 'Guardar

            If Len(Trim(Me.txtRazonSocial.Text)) = 0 Then
                MsgBox "Debe ingresar la Razón Social", vbCritical, Pub_Titulo
                Me.txtRazonSocial.SetFocus
          
            Else
                LimpiaParametros oCmdEjec

                If VNuevo Then
                    oCmdEjec.CommandText = "[sw].[USP_PROVEEDOR_REGISTER]"
                Else
                    oCmdEjec.CommandText = "[sw].[USP_PROVEEDOR_UPDATE]"

                End If

                On Error GoTo grabar

                Dim Smensaje As String

                Dim vIDz     As Integer

                Smensaje = ""
                vIDz = 0

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

                If Not VNuevo Then oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PROVEEDORID", adInteger, adParamInput, , Me.lblProveedorId.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 100, Trim(Me.txtRazonSocial.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        DesactivarControles Me
                        Estado_Botones grabar
                        Me.lvListado.Enabled = True
                        Me.txtSearch.Enabled = True
                
                        proveedorSearch Me.txtSearch.Text
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
            Me.txtRazonSocial.SetFocus
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
                oCmdEjec.CommandText = "[sw].[USP_PROVEEDOR_STATUS]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PROVEEDORID", adInteger, adParamInput, , Me.lblProveedorId.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVO", adBoolean, adParamInput, , False)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        LimpiarControles Me
                        Estado_Botones Desactivar
                        Me.lvListado.Enabled = True
                        proveedorSearch Me.txtSearch.Text
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
                oCmdEjec.CommandText = "[sw].[USP_PROVEEDOR_STATUS]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PROVEEDORID", adInteger, adParamInput, , Me.lblProveedorId.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVO", adBoolean, adParamInput, , True)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)

                Set oRSmain = oCmdEjec.Execute
            
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        LimpiarControles Me
                        Estado_Botones Activar
                        Me.lvListado.Enabled = True
                        proveedorSearch Me.txtSearch.Text
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
                oCmdEjec.CommandText = "[sw].[USP_PROVEEDOR_DELETE]"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PROVEEDORID", adInteger, adParamInput, , Me.lblProveedorId.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
              
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        DesactivarControles Me
                        Estado_Botones Eliminar
                        Me.lvListado.Enabled = True
                        Me.txtSearch.Enabled = True
                
                        proveedorSearch Me.txtSearch.Text
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

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then proveedorSearch Me.txtSearch.Text
End Sub

