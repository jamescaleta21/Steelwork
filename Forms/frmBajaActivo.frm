VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmBajaActivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Baja de Activos"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6375
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6135
      Begin VB.TextBox txtMotivoBaja 
         Height          =   300
         Left            =   360
         TabIndex        =   7
         Tag             =   "X"
         Top             =   3120
         Width           =   5415
      End
      Begin MSComCtl2.DTPicker dtpFechaBaja 
         Height          =   300
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   367919105
         CurrentDate     =   45827
      End
      Begin MSDataListLib.DataCombo DatActivo 
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblUbicacionActual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   360
         TabIndex        =   9
         Tag             =   "X"
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación Actual:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1200
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo Baja:"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Baja:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList iBajaActivo 
      Left            =   5640
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBajaActivo.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBajaActivo.frx":039A
            Key             =   "save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBajaActivo.frx":0734
            Key             =   "undo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbBajaActivo 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
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
Attribute VB_Name = "frmBajaActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DatActivo_Change()
    On Error GoTo xSearch

    If Me.DatActivo.BoundText = -1 Then Exit Sub
    LimpiaParametros oCmdEjec
    oCmdEjec.Prepared = True
    oCmdEjec.CommandText = "[sw].[USP_BAJA_UBICACIONACTUAL_ACTIVO]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVOID", adInteger, adParamInput, , Me.DatActivo.BoundText)
    
    Set oRSmain = oCmdEjec.Execute
    
    If Not oRSmain.EOF Then
        Me.lblUbicacionActual.Caption = oRSmain!ubicacion
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
    Me.tbBajaActivo.ImageList = Me.iBajaActivo
    Me.tbBajaActivo.Buttons(1).Image = Me.iBajaActivo.ListImages(1).index
    Me.tbBajaActivo.Buttons(2).Image = Me.iBajaActivo.ListImages(2).index
    Me.tbBajaActivo.Buttons(3).Image = Me.iBajaActivo.ListImages(3).index

End Sub

Private Sub Estado_Botones(val As Valores)

    Select Case val

        Case InicializarFormulario, grabar, cancelar
            Me.tbBajaActivo.Buttons(1).Enabled = True
            Me.tbBajaActivo.Buttons(2).Enabled = False
            Me.tbBajaActivo.Buttons(3).Enabled = False

        Case Nuevo
            Me.tbBajaActivo.Buttons(1).Enabled = False
            Me.tbBajaActivo.Buttons(2).Enabled = True
            Me.tbBajaActivo.Buttons(3).Enabled = True

    End Select

End Sub

Private Sub tbBajaActivo_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            Me.dtpFechaBaja.Value = LK_FECHA_DIA
            
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "[sw].[USP_MOVIMIENTO_LOAD_ACTIVO]"
            oCmdEjec.Prepared = True
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            
            Set oRSmain = oCmdEjec.Execute
            
            Set Me.DatActivo.RowSource = oRSmain
            Me.DatActivo.BoundColumn = oRSmain.Fields(0).Name
            Me.DatActivo.ListField = oRSmain.Fields(1).Name
            Me.DatActivo.BoundText = -1
            
            Me.DatActivo.SetFocus
            
        Case 2 'Grabar
        
            If Me.DatActivo.BoundText = -1 Then
                MsgBox "Debe elegir el Activo.", vbInformation, Pub_Titulo
                Me.DatActivo.SetFocus
            ElseIf Len(Trim(Me.txtMotivoBaja.Text)) = 0 Then
                MsgBox "Debe ingresar el motivo de Baja.", vbInformation, Pub_Titulo
                Me.txtMotivoBaja.SetFocus
            Else
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "[sw].[USP_BAJA_REGISTER]"
        
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVOID", adInteger, adParamInput, , Me.DatActivo.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHABAJA", adDBTimeStamp, adParamInput, , Me.dtpFechaBaja.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MOTIVOBAJA", adVarChar, adParamInput, 200, Me.txtMotivoBaja.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CUREGISTRO", adVarChar, adParamInput, 20, LK_CODUSU)
                
                Set oRSmain = oCmdEjec.Execute
                
                If Not oRSmain.EOF Then
                    If oRSmain!Code = 0 Then
                        MsgBox oRSmain!Message, vbInformation, Pub_Titulo
                        DesactivarControles Me
                        Estado_Botones grabar
                    Else
                        MsgBox oRSmain!Message, vbCritical, Pub_Titulo

                    End If

                End If
               
            End If

        Case 3 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me

    End Select

End Sub

Private Sub txtMotivoBaja_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
End Sub
