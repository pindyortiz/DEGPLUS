VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmClientes 
   Caption         =   "Clientes"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   14580
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adoCanales 
      Height          =   330
      Left            =   9000
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
   Begin VB.Frame frcClientes 
      Height          =   8595
      Left            =   300
      TabIndex        =   11
      Top             =   630
      Width           =   19545
      Begin VB.Frame Frame2 
         Height          =   8295
         Left            =   7290
         TabIndex        =   7
         Top             =   150
         Width           =   12105
         Begin VB.Frame Frame4 
            Height          =   795
            Left            =   90
            TabIndex        =   8
            Top             =   150
            Width           =   11895
            Begin VB.CommandButton cmdASCDES 
               Caption         =   "ASC"
               Height          =   285
               Left            =   6240
               TabIndex        =   32
               Top             =   270
               Width           =   735
            End
            Begin VB.OptionButton optTodos 
               Caption         =   "Todos"
               Height          =   285
               Left            =   90
               TabIndex        =   9
               Top             =   270
               Width           =   1125
            End
            Begin VB.OptionButton optActivos 
               Caption         =   "Activos"
               Height          =   285
               Left            =   1200
               TabIndex        =   10
               Top             =   270
               Width           =   1125
            End
            Begin VB.OptionButton optInactivos 
               Caption         =   "Inactivos"
               Height          =   285
               Left            =   2310
               TabIndex        =   29
               Top             =   270
               Width           =   1125
            End
            Begin VB.ComboBox cmbCamposOrden 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   4440
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   270
               Width           =   1725
            End
            Begin VB.ComboBox cmbCamposFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   7890
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   270
               Width           =   1725
            End
            Begin VB.TextBox txtFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   345
               Left            =   9690
               TabIndex        =   26
               Top             =   270
               Width           =   2085
            End
            Begin VB.Label Label6 
               Caption         =   "Orden"
               Height          =   255
               Left            =   3930
               TabIndex        =   31
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Label7 
               Caption         =   "Filtro"
               Height          =   255
               Left            =   7440
               TabIndex        =   30
               Top             =   300
               Width           =   765
            End
         End
         Begin MSDataGridLib.DataGrid dtgClientes 
            Bindings        =   "frmClientes.frx":0000
            Height          =   7185
            Left            =   90
            TabIndex        =   25
            Top             =   990
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   12674
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            Caption         =   "Listado de clientes"
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "IdCliente"
               Caption         =   "Cliente"
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
               DataField       =   "RazonSocial"
               Caption         =   "Razón social"
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
            BeginProperty Column03 
               DataField       =   "CodigoVendedor"
               Caption         =   "Vendedor"
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
            BeginProperty Column04 
               DataField       =   "HabilitaDistribucion"
               Caption         =   "Habilita distribución"
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
            BeginProperty Column05 
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
                  ColumnWidth     =   1200,189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3644,788
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2505,26
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   1200,189
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnWidth     =   1500,095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1200,189
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8295
         Left            =   150
         TabIndex        =   12
         Top             =   150
         Width           =   7035
         Begin VB.ComboBox cmbED_HabilitaDistribucion 
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
            Left            =   5490
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2250
            Width           =   1425
         End
         Begin VB.TextBox txtDB_HabilitaDistribucion 
            DataField       =   "HabilitaDistribucion"
            DataSource      =   "adoClientes"
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
            Left            =   5490
            TabIndex        =   38
            Top             =   2250
            Width           =   1425
         End
         Begin VB.TextBox txtED_CodigoVendedor 
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
            Left            =   3660
            TabIndex        =   4
            Top             =   2250
            Width           =   1725
         End
         Begin VB.TextBox txtDB_CodigoVendedor 
            DataField       =   "CodigoVendedor"
            DataSource      =   "adoClientes"
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
            Left            =   3660
            TabIndex        =   36
            Top             =   2250
            Width           =   1725
         End
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
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   3090
            Width           =   1755
         End
         Begin VB.TextBox txtDB_Activo 
            DataField       =   "Activo"
            DataSource      =   "adoClientes"
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
            Left            =   150
            TabIndex        =   34
            Top             =   3090
            Width           =   1755
         End
         Begin VB.ComboBox cmbED_Canal 
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
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2250
            Width           =   3405
         End
         Begin VB.TextBox txtDB_Canal 
            DataField       =   "Canal"
            DataSource      =   "adoClientes"
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
            Left            =   150
            TabIndex        =   33
            Top             =   2250
            Width           =   3405
         End
         Begin VB.TextBox txtED_IdCliente 
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
            Height          =   450
            Left            =   150
            TabIndex        =   1
            Top             =   600
            Width           =   2115
         End
         Begin VB.TextBox txtED_RazonSocial 
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
            Height          =   510
            Left            =   120
            TabIndex        =   2
            Top             =   1440
            Width           =   6765
         End
         Begin VB.Frame Frame3 
            Height          =   1125
            Left            =   150
            TabIndex        =   19
            Top             =   7020
            Width           =   6735
            Begin VB.CommandButton cmdNuevo 
               Caption         =   "Nuevo"
               Height          =   795
               Left            =   120
               Picture         =   "frmClientes.frx":001A
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdModificar 
               Caption         =   "Modificar"
               Height          =   795
               Left            =   1447
               Picture         =   "frmClientes.frx":045C
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdGuardar 
               Caption         =   "Guardar"
               Height          =   795
               Left            =   2774
               Picture         =   "frmClientes.frx":089E
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdCancelar 
               Caption         =   "Cancelar"
               Height          =   795
               Left            =   4101
               Picture         =   "frmClientes.frx":0CE0
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdSalir 
               Caption         =   "Salir"
               Height          =   795
               Left            =   5430
               Picture         =   "frmClientes.frx":1122
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   210
               Width           =   1155
            End
         End
         Begin VB.TextBox txtDB_RazonSocial 
            DataField       =   "RazonSocial"
            DataSource      =   "adoClientes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   6765
         End
         Begin VB.TextBox txtDB_IdCliente 
            DataField       =   "IdCliente"
            DataSource      =   "adoClientes"
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
            Left            =   150
            TabIndex        =   14
            Top             =   600
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Habilita distribución"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5490
            TabIndex        =   39
            Top             =   2010
            Width           =   1515
         End
         Begin VB.Label lbl_CodigoVendedor 
            Caption         =   "Código de vendedor"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3660
            TabIndex        =   37
            Top             =   2010
            Width           =   1725
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
            TabIndex        =   35
            Top             =   6750
            Visible         =   0   'False
            Width           =   6705
         End
         Begin VB.Label Label4 
            Caption         =   "Activo"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   18
            Top             =   2850
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Canal"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2010
            Width           =   2055
         End
         Begin VB.Label lbl_RazonSocial 
            Caption         =   "Razón social"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1170
            Width           =   2055
         End
         Begin VB.Label lbl_IdCliente 
            Caption         =   "Código"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   330
            Width           =   2055
         End
      End
   End
   Begin MSAdodcLib.Adodc adoClientes 
      Height          =   330
      Left            =   9000
      Top             =   330
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
      Caption         =   "CLIENTES"
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
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   8625
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vsPasoNM As String ' Variable de paso por Nuevo o Modificar
Dim vsRegistrosClientes As Integer ' Cuantos registros tiene la tabla de Clientes
Dim vsRegistrosCanales As Integer ' Cuantos registros tiene la tabla de Canales
Dim vsI As Integer ' Contador simple
Dim vsTAI As String ' Lista Todos, Activos o Inactivos
Dim vsOrden As String ' Establace el orden de los registros
Dim vsASCDES As String ' Establece si el orden es ASCendente o DEScendente
Dim vsCampo As String ' Establece que campo se usara para el filtro
Dim vsFiltro As String ' Filtra los canales
Dim vsIdCliente_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsRazonSocial_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsCodigoVendedor_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio

Private Sub Form_Load()
  
  lblTitulo.Top = 0
  lblTitulo.Left = 0
  lblTitulo.Width = Screen.Width
  frcClientes.Left = (Screen.Width - frcClientes.Width) / 2
  
  proLimpiaCampos
  
  adoCanales.RecordSource = "SELECT * FROM [Canales] ORDER BY [IdCanal]"
  adoCanales.CommandType = adCmdText
  adoCanales.Refresh
  vsRegistrosCanales = adoCanales.Recordset.RecordCount
  
  If vsRegistrosCanales <> 0 Then
    adoCanales.Recordset.MoveFirst
    For vsI = 1 To vsRegistrosCanales
      cmbED_Canal.AddItem adoCanales.Recordset![Canal]
      adoCanales.Recordset.MoveNext
    Next vsI
    cmbED_Canal.Text = cmbED_Canal.List(0)
  Else
    MsgBox "La tabla de canales de marketing no puede estar vacía para inresar al modulo de clientes. Por favor ingrese en el modulo de Canales de marketing para agregarlos", vbInformation, "Datos incompletos"
    Exit Sub
  End If
  
  cmbED_Activo.AddItem "Activo"
  cmbED_Activo.AddItem "Inactivo"
  cmbED_Activo.Text = cmbED_Activo.List(0)
  
  cmbED_HabilitaDistribucion.AddItem "Si"
  cmbED_HabilitaDistribucion.AddItem "No"
  cmbED_HabilitaDistribucion.Text = cmbED_HabilitaDistribucion.List(0)
  
  cmbCamposOrden.AddItem "Código"
  cmbCamposOrden.AddItem "Razón social"
  cmbCamposOrden.AddItem "Canal"
  cmbCamposOrden.AddItem "Código de vendedor"
  cmbCamposOrden.AddItem "Habilita distribución"
  cmbCamposOrden.AddItem "Activo"
  cmbCamposOrden.Text = cmbCamposOrden.List(0)
  
  cmbCamposFiltro.AddItem "Código"
  cmbCamposFiltro.AddItem "Razón social"
  cmbCamposFiltro.AddItem "Canal"
  cmbCamposFiltro.AddItem "Código de vendedor"
  cmbCamposFiltro.AddItem "Habilita distribución"
  cmbCamposFiltro.AddItem "Activo"
  cmbCamposFiltro.Text = cmbCamposFiltro.List(0)
  
  optTodos.Value = True
  
  Call proListadoFull
  
  If vsRegistrosClientes = 0 Then
    Call proActivarBotones(True, False, False, False, True)
  Else
    Call proActivarBotones(True, True, False, False, True)
  End If
  
  vsTAI = "Todos"
  vsOrden = "IdCliente"
  vsASCDES = "ASC"
  vsCampo = "IdCliente"
  vsFiltro = ""
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub txtED_IdCliente_Change()

  If txtED_IdCliente.Visible = True Then
    If txtED_IdCliente.Text <> "" Then
      lbl_IdCliente.ForeColor = vbBlack
      lbl_IdCliente.FontBold = False
      vsIdCliente_Completo = True
    Else
      lbl_IdCliente.ForeColor = vbRed
      lbl_IdCliente.FontBold = True
      vsIdCliente_Completo = False
    End If
  
    If (vsIdCliente_Completo And vsRazonSocial_Completo And vsCodigoVendedor_Completo) = True Then
      lblRequeridos.Visible = False
    Else
      lblRequeridos.Visible = True
    End If
  End If
  
End Sub

Private Sub txtED_RazonSocial_Change()

  If txtED_RazonSocial.Visible = True Then
    If txtED_RazonSocial.Text <> "" Then
      lbl_RazonSocial.ForeColor = vbBlack
      lbl_RazonSocial.FontBold = False
      vsRazonSocial_Completo = True
    Else
      lbl_RazonSocial.ForeColor = vbRed
      lbl_RazonSocial.FontBold = True
      vsRazonSocial_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
      If (vsIdCliente_Completo And vsRazonSocial_Completo And vsCodigoVendedor_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If vsRazonSocial_Completo Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If
  
End Sub

Private Sub txtED_CodigoVendedor_Change()

  If txtED_CodigoVendedor.Visible = True Then
    If txtED_CodigoVendedor.Text <> "" Then
      lbl_CodigoVendedor.ForeColor = vbBlack
      lbl_CodigoVendedor.FontBold = False
      vsCodigoVendedor_Completo = True
    Else
      lbl_CodigoVendedor.ForeColor = vbRed
      lbl_CodigoVendedor.FontBold = True
      vsCodigoVendedor_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
      If (vsIdCliente_Completo And vsRazonSocial_Completo And vsCodigoVendedor_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If vsRazonSocial_Completo Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If
  
End Sub


Private Sub cmdNuevo_Click()

  vsPasoNM = "NUEVO"
  
  txtED_IdCliente.Text = ""
  txtED_RazonSocial.Text = ""
  txtED_CodigoVendedor.Text = ""

  txtED_IdCliente.Visible = True
  txtED_RazonSocial.Visible = True
  cmbED_Canal.Visible = True
  txtED_CodigoVendedor.Visible = True
  cmbED_HabilitaDistribucion.Visible = True
  cmbED_Activo.Visible = True
   
  txtED_IdCliente.SetFocus
  
  lblRequeridos.Visible = True
  lbl_IdCliente.ForeColor = vbRed
  lbl_IdCliente.FontBold = True
  lbl_RazonSocial.ForeColor = vbRed
  lbl_RazonSocial.FontBold = True
  lbl_CodigoVendedor.ForeColor = vbRed
  lbl_CodigoVendedor.FontBold = True
  
  vsIdCliente_Completo = False
  vsRazonSocial_Completo = False
  vsCodigoVendedor_Completo = False
  
  Call proADControlGrid
  
  Call proActivarBotones(False, False, True, True, False)
  
End Sub

Private Sub cmdModificar_Click()

  vsPasoNM = "MODIFICAR"
  
  Call proADControlGrid
  
  txtED_RazonSocial.Visible = True
  cmbED_Canal.Visible = True
  txtED_CodigoVendedor.Visible = True
  cmbED_HabilitaDistribucion.Visible = True
  cmbED_Activo.Visible = True
  
  txtED_RazonSocial.Text = txtDB_RazonSocial.Text
  cmbED_Canal.Text = txtDB_Canal.Text
  txtED_CodigoVendedor.Text = txtDB_CodigoVendedor.Text
  cmbED_HabilitaDistribucion.Text = txtDB_HabilitaDistribucion.Text
  cmbED_Activo.Text = txtDB_Activo.Text
  
  txtED_RazonSocial.SelStart = 0
  txtED_RazonSocial.SelLength = Len(txtED_RazonSocial.Text)
  
  txtED_RazonSocial.SetFocus
  
  Call proActivarBotones(False, False, True, True, False)
   
End Sub

Private Sub cmdGuardar_Click()

  If txtED_RazonSocial.Text = "" Then
    MsgBox "El dato de 'Razón social' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
  If txtED_CodigoVendedor.Text = "" Then
    MsgBox "El dato de 'Código de vendedor' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
  If Not IsNumeric(txtED_CodigoVendedor.Text) Then
    MsgBox "El dato de 'Código de vendedor' debe ser un número. Debe coregir ese dato.", vbExclamation, "Dato erroneo"
    Exit Sub
  End If
       
  If vsPasoNM = "NUEVO" Then
    If txtED_IdCliente.Text = "" Then
      MsgBox "El dato de 'Código' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
      Exit Sub
    End If
    If Not IsNumeric(txtED_IdCliente) Then
      MsgBox "El dato de 'Código' debe ser un número. Debe corregir este dato.", vbExclamation, "Dato erroneo"
      Exit Sub
    End If
  
    adoClientes.Recordset.AddNew
    adoClientes.Recordset![IdCliente] = txtED_IdCliente.Text
  End If
    adoClientes.Recordset![RazonSocial] = txtED_RazonSocial.Text
    adoClientes.Recordset![Canal] = cmbED_Canal.Text
    adoClientes.Recordset![CodigoVendedor] = txtED_CodigoVendedor.Text
    adoClientes.Recordset![HabilitaDistribucion] = cmbED_HabilitaDistribucion.Text
    adoClientes.Recordset![Activo] = cmbED_Activo.Text
    adoClientes.Recordset.Update
    adoClientes.Refresh
  
  Call proListadoFull
  
  Call cmdCancelar_Click

End Sub

Private Sub cmdCancelar_Click()

  proLimpiaCampos
  
  lbl_IdCliente.ForeColor = vbBlack
  lbl_IdCliente.FontBold = False
  lbl_RazonSocial.ForeColor = vbBlack
  lbl_RazonSocial.FontBold = False
  lbl_CodigoVendedor.ForeColor = vbBlack
  lbl_CodigoVendedor.FontBold = False
  
  lblRequeridos.Visible = False
  
  If vsRegistrosClientes = 0 Then
    Call proActivarBotones(True, False, False, False, True)
  Else
    Call proActivarBotones(True, True, False, False, True)
  End If

  txtDB_IdCliente.SetFocus
  
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
 
  Select Case cmbCamposOrden
    Case "Código": vsOrden = "IdCliente"
    Case "Razón social": vsOrden = "Razonsocial"
    Case "Canal": vsOrden = "Canal"
    Case "Código de vendedor": vsOrden = "CodigoVendedor"
    Case "Habilita distribución": vsOrden = "HabilitaDistribucion"
    Case "Activo": vsOrden = "Activo"
  End Select
  
  vsTAI = "Todos"
  vsASCDES = "ASC"
  vsCampo = "IdCliente"
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

  Select Case cmbCamposFiltro
    Case "Código": vsCampo = "IdCliente"
    Case "Razón social": vsCampo = "Razonsocial"
    Case "Canal": vsCampo = "Canal"
    Case "Código de vendedor": vsOrden = "CodigoVendedor"
    Case "Habilita distribución": vsOrden = "HabilitaDistribucion"
    Case "Activo": vsCampo = "Activo"
  End Select
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub txtFiltro_Change()
  
  vsFiltro = txtFiltro.Text
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub proLimpiaCampos()

  txtED_IdCliente.Text = ""
  txtED_RazonSocial.Text = ""
  txtED_CodigoVendedor.Text = ""
  
  txtED_IdCliente.Visible = False
  txtED_RazonSocial.Visible = False
  cmbED_Canal.Visible = False
  txtED_CodigoVendedor.Visible = False
  cmbED_HabilitaDistribucion.Visible = False
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

  adoClientes.RecordSource = "SELECT * FROM [Clientes] ORDER BY [IdCliente]"
  adoClientes.CommandType = adCmdText
  adoClientes.Refresh
  vsRegistrosClientes = adoClientes.Recordset.RecordCount
  
End Sub

Private Sub proListarRegistros(ByVal Estado As String, ByVal Orden As String, ByVal ASCDES As String, ByVal Campo As String, ByVal Filtro As String)

  If Estado = "Todos" Then
    adoClientes.RecordSource = "SELECT * FROM [Clientes] WHERE [" & Campo & "] LIKE '%" & Filtro & "%' ORDER BY " & Orden & " " & ASCDES
  Else
    adoClientes.RecordSource = "SELECT * FROM [Clientes] WHERE [Activo]= '" & Estado & "' ORDER BY [" & Orden & "]" & ASCDES
  End If
  adoClientes.CommandType = adCmdText
  adoClientes.Refresh

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

