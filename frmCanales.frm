VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCanales 
   Caption         =   "Canales de marketing"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   12150
   WindowState     =   2  'Maximized
   Begin VB.Frame frcCanales 
      Height          =   8595
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   19545
      Begin VB.Frame Frame1 
         Height          =   8295
         Left            =   150
         TabIndex        =   15
         Top             =   150
         Width           =   7035
         Begin VB.ComboBox cmbED_Activo 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2250
            Width           =   3255
         End
         Begin VB.TextBox txtDB_Activo 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Activo"
            DataSource      =   "adoCanales"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   2250
            Width           =   3225
         End
         Begin VB.TextBox txtED_Canal 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   90
            TabIndex        =   2
            Top             =   1410
            Width           =   6825
         End
         Begin VB.TextBox txtED_IdCanal 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   90
            TabIndex        =   1
            Top             =   570
            Width           =   2115
         End
         Begin VB.TextBox txtDB_IdCanal 
            DataField       =   "IdCanal"
            DataSource      =   "adoCanales"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   570
            Width           =   2115
         End
         Begin VB.TextBox txtDB_Canal 
            DataField       =   "Canal"
            DataSource      =   "adoCanales"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1410
            Width           =   6825
         End
         Begin VB.Frame Frame3 
            Height          =   1125
            Left            =   150
            TabIndex        =   16
            Top             =   7020
            Width           =   6735
            Begin VB.CommandButton cmdSalir 
               Caption         =   "Salir"
               Height          =   795
               Left            =   5430
               Picture         =   "frmCanales.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdCancelar 
               Caption         =   "Cancelar"
               Height          =   795
               Left            =   4101
               Picture         =   "frmCanales.frx":0442
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdGuardar 
               Caption         =   "Guardar"
               Height          =   795
               Left            =   2774
               Picture         =   "frmCanales.frx":0884
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdModificar 
               Caption         =   "Modificar"
               Height          =   795
               Left            =   1447
               Picture         =   "frmCanales.frx":0CC6
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdNuevo 
               Caption         =   "Nuevo"
               Height          =   795
               Left            =   120
               Picture         =   "frmCanales.frx":1108
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   210
               Width           =   1155
            End
         End
         Begin VB.Label lbl_IdCanal 
            Caption         =   "Código"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   330
            Width           =   2055
         End
         Begin VB.Label lbl_Canal 
            Caption         =   "Canal de marketing"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1170
            Width           =   2055
         End
         Begin VB.Label label1 
            Caption         =   "Activo"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2010
            Width           =   2055
         End
         Begin VB.Label lblRequeridos 
            Alignment       =   2  'Center
            Caption         =   "Algunos campos requeridos aún están vacíos."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   150
            TabIndex        =   24
            Top             =   6750
            Visible         =   0   'False
            Width           =   6735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   8295
         Left            =   7290
         TabIndex        =   4
         Top             =   150
         Width           =   12105
         Begin VB.Frame Frame4 
            Height          =   795
            Left            =   90
            TabIndex        =   5
            Top             =   150
            Width           =   11895
            Begin VB.CommandButton cmdASCDES 
               Caption         =   "ASC"
               Height          =   285
               Left            =   6240
               TabIndex        =   30
               Top             =   270
               Width           =   735
            End
            Begin VB.TextBox txtFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   345
               Left            =   9690
               TabIndex        =   11
               Top             =   270
               Width           =   2085
            End
            Begin VB.ComboBox cmbCamposFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   7890
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   270
               Width           =   1725
            End
            Begin VB.ComboBox cmbCamposOrden 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   4440
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   270
               Width           =   1725
            End
            Begin VB.OptionButton optInactivos 
               Caption         =   "Inactivos"
               Height          =   285
               Left            =   2310
               TabIndex        =   8
               Top             =   270
               Width           =   1125
            End
            Begin VB.OptionButton optActivos 
               Caption         =   "Activos"
               Height          =   285
               Left            =   1200
               TabIndex        =   7
               Top             =   270
               Width           =   1125
            End
            Begin VB.OptionButton optTodos 
               Caption         =   "Todos"
               Height          =   285
               Left            =   90
               TabIndex        =   6
               Top             =   270
               Width           =   1125
            End
            Begin VB.Label Label7 
               Caption         =   "Filtro"
               Height          =   255
               Left            =   7440
               TabIndex        =   13
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Label6 
               Caption         =   "Orden"
               Height          =   255
               Left            =   3930
               TabIndex        =   12
               Top             =   300
               Width           =   765
            End
         End
         Begin MSDataGridLib.DataGrid dtgClientes 
            Bindings        =   "frmCanales.frx":154A
            Height          =   7185
            Left            =   90
            TabIndex        =   14
            Top             =   990
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   12674
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Listado de canales"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "IdCanal"
               Caption         =   "Id Canal"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Canal"
               Caption         =   "Canal"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Activo"
               Caption         =   "Activo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   9195,024
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1140,095
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSAdodcLib.Adodc adoCanales 
      Height          =   330
      Left            =   8700
      Top             =   0
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\DEGPLUS\DB\dbDEG.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\DEGPLUS\DB\dbDEG.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Canales"
      Caption         =   "adoCanales"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoClientes 
      Height          =   330
      Left            =   8700
      Top             =   360
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\DEGPLUS\DB\dbDEG.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\DEGPLUS\DB\dbDEG.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Clientes"
      Caption         =   "adoClientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "CANALES DE MARKETING"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   60
      TabIndex        =   28
      Top             =   0
      Width           =   8625
   End
End
Attribute VB_Name = "frmCanales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vsPasoNM As String ' Variable de paso por Nuevo o Modificar
Dim vsRegistros As String ' Cuanto registros tiene la tabla de Canales
Dim vsTAI As String ' Lista Todos, Activos o Inactivos
Dim vsOrden As String ' Establace el orden de los registros
Dim vsASCDES As String ' Establece si el orden es ASCendente o DEScendente
Dim vsCampo As String ' Establece que campo se usara para el filtro
Dim vsFiltro As String ' Filtra los canales
Dim vsIdCanal_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsCanal_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsCanalAnterior As String ' Almacena el valor anterior del canal luego de una modificacion
Dim vsI As Integer ' Contador simple

'//// FALTA EL CODIGO QUE ACTUALIZA LOS CANALES DE LA TABLA CLIENTES CUANDO SE MODIFICA EL NOMBRE
'//// DE UN CANAL EN EL MODULO DE CANALES


Private Sub Form_Load()
  
  lblTitulo.Top = 0
  lblTitulo.Left = 0
  lblTitulo.Width = Screen.Width
  frcCanales.Left = (Screen.Width - frcCanales.Width) / 2
  
  proLimpiaCampos
  
  cmbED_Activo.AddItem "Activo"
  cmbED_Activo.AddItem "Inactivo"
  cmbED_Activo.Text = cmbED_Activo.List(0)
  
  cmbCamposOrden.AddItem "Código"
  cmbCamposOrden.AddItem "Canal"
  cmbCamposOrden.AddItem "Activo"
  cmbCamposOrden.Text = cmbCamposOrden.List(0)
  
  cmbCamposFiltro.AddItem "Código"
  cmbCamposFiltro.AddItem "Canal"
  cmbCamposFiltro.AddItem "Activo"
  cmbCamposFiltro.Text = cmbCamposFiltro.List(0)
  
  optTodos.Value = True
  
  Call proListadoFull
  
  If vsRegistros = 0 Then
    Call proActivarBotones(True, False, False, False, True)
  Else
    Call proActivarBotones(True, True, False, False, True)
  End If
  
  vsTAI = "Todos"
  vsOrden = "IdCanal"
  vsASCDES = "ASC"
  vsCampo = "IdCanal"
  vsFiltro = ""
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub txtED_IdCanal_Change()

  If txtED_IdCanal.Visible = True Then
    If txtED_IdCanal.Text <> "" Then
      lbl_IdCanal.ForeColor = vbBlack
      lbl_IdCanal.FontBold = False
      vsIdCanal_Completo = True
    Else
      lbl_IdCanal.ForeColor = vbRed
      lbl_IdCanal.FontBold = True
      vsIdCanal_Completo = False
    End If
  
    If (vsIdCanal_Completo And vsCanal_Completo) = True Then
      lblRequeridos.Visible = False
    Else
      lblRequeridos.Visible = True
    End If
  End If
  
End Sub

Private Sub txtED_Canal_Change()

  If txtED_Canal.Visible = True Then
    If txtED_Canal.Text <> "" Then
      lbl_Canal.ForeColor = vbBlack
      lbl_Canal.FontBold = False
      vsCanal_Completo = True
    Else
      lbl_Canal.ForeColor = vbRed
      lbl_Canal.FontBold = True
      vsCanal_Completo = False
    End If
    If vsPasoNM = "NUEVO" Then
      If (vsIdCanal_Completo And vsCanal_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If vsCanal_Completo Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If
  
End Sub

Private Sub cmdNuevo_Click()

  vsPasoNM = "NUEVO"
  
  txtED_IdCanal.Text = ""
  txtED_Canal.Text = ""
  
  txtED_IdCanal.Visible = True
  txtED_Canal.Visible = True
  cmbED_Activo.Visible = True
   
  txtED_IdCanal.SetFocus
  
  lblRequeridos.Visible = True
  lbl_IdCanal.ForeColor = vbRed
  lbl_IdCanal.FontBold = True
  lbl_Canal.ForeColor = vbRed
  lbl_Canal.FontBold = True
  
  vsIdCanal_Completo = False
  vsCanal_Completo = False
   
  Call proADControlGrid
   
  Call proActivarBotones(False, False, True, True, False)
  
End Sub

Private Sub cmdModificar_Click()

  vsPasoNM = "MODIFICAR"
  
  Call proADControlGrid
  
  txtED_Canal.Visible = True
  cmbED_Activo.Visible = True
  
  txtED_Canal.Text = txtDB_Canal.Text
  cmbED_Activo.Text = txtDB_Activo.Text
  
  txtED_Canal.SelStart = 0
  txtED_Canal.SelLength = Len(txtED_Canal.Text)
  
  txtED_Canal.SetFocus
  
  Call proActivarBotones(False, False, True, True, False)
   
End Sub

Private Sub cmdGuardar_Click()

  If txtED_Canal.Text = "" Then
    MsgBox "El dato de 'Canal' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
    
  If vsPasoNM = "NUEVO" Then
    If txtED_IdCanal.Text = "" Then
      MsgBox "El dato de 'Código' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
      Exit Sub
    End If
    If Not IsNumeric(txtED_IdCanal) Then
      MsgBox "El dato de 'Código' debe ser un número. Debe corregir este dato.", vbExclamation, "Dato erroneo"
      Exit Sub
    End If
  
    adoCanales.Recordset.AddNew
    adoCanales.Recordset![IdCanal] = txtED_IdCanal.Text
  End If
    
  vsCanalAnterior = adoCanales.Recordset![Canal]
  adoCanales.Recordset![Canal] = txtED_Canal.Text
  adoCanales.Recordset![Activo] = cmbED_Activo.Text
  adoCanales.Recordset.Update
  adoCanales.Refresh
  
  If vsPasoNM = "MODIFICAR" Then
    adoClientes.RecordSource = "SELECT * FROM [Clientes] WHERE [Canal]='" & vsCanalAnterior & "' ORDER BY [Canal]"
    adoClientes.CommandType = adCmdText
    adoClientes.Refresh
    If adoClientes.Recordset.RecordCount <> 0 Then
      adoClientes.Recordset.MoveFirst
      For vsI = 1 To adoClientes.Recordset.RecordCount
        adoClientes.Recordset![Canal] = txtED_Canal.Text
        adoClientes.Recordset.Update
        adoClientes.Refresh
        adoClientes.Recordset.MoveNext
      Next vsI
    End If
  End If
  
  Call proListadoFull
  
  Call cmdCancelar_Click

End Sub

Private Sub cmdCancelar_Click()

  proLimpiaCampos
  
  lbl_IdCanal.ForeColor = vbBlack
  lbl_IdCanal.FontBold = False
  lbl_Canal.ForeColor = vbBlack
  lbl_Canal.FontBold = False
  
  lblRequeridos.Visible = False
  
  If vsRegistros = 0 Then
    Call proActivarBotones(True, False, False, False, True)
  Else
    Call proActivarBotones(True, True, False, False, True)
  End If

  txtDB_IdCanal.SetFocus
  
  Call proADControlGrid

End Sub

Private Sub cmdSalir_Click()

  Unload Me

End Sub

Private Sub optTodos_Click()

  If optTodos.Value = True Then
    vsTAI = "Todos"
  End If

  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)

End Sub

Private Sub optActivos_Click()
  
  If optActivos.Value = True Then
    vsTAI = "Activo"
  End If
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub optInactivos_Click()

  If optInactivos.Value = True Then
    vsTAI = "Inactivo"
  End If

  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub cmbCamposOrden_Click()
 
  If cmbCamposOrden.Text = "Código" Then
    vsOrden = "IdCanal"
  Else
    vsOrden = cmbCamposOrden.Text
  End If
  
  vsTAI = "Todos"
  vsASCDES = "ASC"
  vsCampo = "IdCanal"
  vsFiltro = ""
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)

End Sub

Private Sub cmdASCDES_Click()

  If vsASCDES = "ASC" Then
    vsASCDES = "DESC"
    cmdASCDES.Caption = "DES"
  Else
    vsASCDES = "ASC"
    cmdASCDES.Caption = "ASC"
  End If

  vsTAI = "Todos"
  'vsFiltro = ""
  
  dtgClientes.SetFocus

  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)

End Sub

Private Sub cmbCamposFiltro_Click()

  If cmbCamposFiltro.Text = "Código" Then
    vsCampo = "IdCanal"
  Else
    vsCampo = cmbCamposFiltro.Text
  End If

  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub txtFiltro_Change()
  
  vsFiltro = txtFiltro.Text
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub proLimpiaCampos()

  txtED_IdCanal.Text = ""
  txtED_Canal.Text = ""
  
  txtED_IdCanal.Visible = False
  txtED_Canal.Visible = False
  cmbED_Activo.Visible = False

End Sub

Private Sub proActivarBotones(ByVal N As Boolean, ByVal M As Boolean, ByVal G As Boolean, ByVal C As Boolean, ByVal S As Boolean)

  cmdNuevo.Enabled = N
  cmdModificar.Enabled = M
  cmdGuardar.Enabled = G
  cmdCancelar.Enabled = C
  cmdSalir.Enabled = S

End Sub

Private Sub proListadoFull()

  adoCanales.RecordSource = "SELECT * FROM [Canales] ORDER BY [IdCanal]"
  adoCanales.CommandType = adCmdText
  adoCanales.Refresh
  vsRegistros = adoCanales.Recordset.RecordCount
End Sub

Private Sub proListarRegistros(ByVal Estado As String, ByVal Orden As String, ByVal ASCDES As String, ByVal Campo As String, ByVal Filtro As String)

  If Estado = "Todos" Then
    adoCanales.RecordSource = "SELECT * FROM [Canales] WHERE [" & Campo & "] LIKE '%" & Filtro & "%' ORDER BY " & Orden & " " & ASCDES
  Else
    adoCanales.RecordSource = "SELECT * FROM [Canales] WHERE [Activo]= '" & Estado & "' ORDER BY [" & Orden & "]" & ASCDES
  End If
  adoCanales.CommandType = adCmdText
  adoCanales.Refresh

End Sub

Private Sub proADControlGrid() ' Procedimiento que Activa/Desactiva el Control Grid
  
  If optTodos.Enabled = False Then
    optTodos.Enabled = True
    optActivos.Enabled = True
    optInactivos.Enabled = True
    cmbCamposOrden.Enabled = True
    cmdASCDES.Enabled = True
    cmbCamposFiltro.Enabled = True
    txtFiltro.Enabled = True
  Else
    optTodos.Enabled = False
    optActivos.Enabled = False
    optInactivos.Enabled = False
    cmbCamposOrden.Enabled = False
    cmdASCDES.Enabled = False
    cmbCamposFiltro.Enabled = False
    txtFiltro.Enabled = False
  End If

End Sub
