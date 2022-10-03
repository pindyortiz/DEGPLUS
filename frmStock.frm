VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStock 
   Caption         =   "Stock"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame frcArticulos 
      Height          =   8595
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   19545
      Begin VB.Frame Frame2 
         Height          =   8295
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   12105
         Begin VB.Frame Frame4 
            Height          =   795
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   11895
            Begin VB.OptionButton optTodos 
               Caption         =   "Todos"
               Height          =   285
               Left            =   90
               TabIndex        =   21
               Top             =   270
               Width           =   1125
            End
            Begin VB.OptionButton optActivos 
               Caption         =   "Activos"
               Height          =   285
               Left            =   1200
               TabIndex        =   20
               Top             =   270
               Width           =   1125
            End
            Begin VB.OptionButton optInactivos 
               Caption         =   "Inactivos"
               Height          =   285
               Left            =   2310
               TabIndex        =   19
               Top             =   270
               Width           =   1125
            End
            Begin VB.ComboBox cmbCamposOrden 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   4440
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   270
               Width           =   1725
            End
            Begin VB.ComboBox cmbCamposFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   7890
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   270
               Width           =   1725
            End
            Begin VB.TextBox txtFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   345
               Left            =   9690
               TabIndex        =   16
               Top             =   270
               Width           =   2085
            End
            Begin VB.CommandButton cmdASCDES 
               Caption         =   "ASC"
               Height          =   285
               Left            =   6240
               TabIndex        =   15
               Top             =   270
               Width           =   735
            End
            Begin VB.Label Label6 
               Caption         =   "Orden"
               Height          =   255
               Left            =   3930
               TabIndex        =   23
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Label7 
               Caption         =   "Filtro"
               Height          =   255
               Left            =   7440
               TabIndex        =   22
               Top             =   300
               Width           =   765
            End
         End
         Begin MSDataGridLib.DataGrid dtgArticulos 
            Bindings        =   "frmStock.frx":0000
            Height          =   7185
            Left            =   120
            TabIndex        =   24
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
            Caption         =   "Listado de artículos"
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "IdArticulo"
               Caption         =   "Artículo"
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
               DataField       =   "IdSegunProveedor"
               Caption         =   "Id x Proveedor"
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
               DataField       =   "Descripcion"
               Caption         =   "Descripción"
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
               DataField       =   "UxB"
               Caption         =   "UxB"
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
               DataField       =   "PrecioVenta"
               Caption         =   "Precio"
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
               DataField       =   "StockBultos"
               Caption         =   "Bultos"
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
            BeginProperty Column06 
               DataField       =   "StockUnidades"
               Caption         =   "Unidades"
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
            BeginProperty Column07 
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
                  Alignment       =   1
                  ColumnWidth     =   1200,189
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1200,189
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   3945,26
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   494,929
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1094,74
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnWidth     =   1094,74
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  ColumnWidth     =   1094,74
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1140,095
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8295
         Left            =   12390
         TabIndex        =   1
         Top             =   120
         Width           =   7035
         Begin VB.TextBox txtDiferenciaUnidades 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   5640
            TabIndex        =   34
            Top             =   1410
            Width           =   1035
         End
         Begin VB.TextBox txtDB_StockUnidades 
            DataField       =   "StockUnidades"
            DataSource      =   "adoArticulos"
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
            Left            =   1200
            TabIndex        =   32
            Top             =   1410
            Width           =   1035
         End
         Begin VB.TextBox txtNuevoStockBultos 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
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
            Left            =   2310
            TabIndex        =   31
            Top             =   1410
            Width           =   1035
         End
         Begin VB.TextBox txtDiferenciaBultos 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   4530
            TabIndex        =   29
            Top             =   1410
            Width           =   1035
         End
         Begin VB.TextBox txtNuevoStockUnidades 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
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
            Left            =   3420
            TabIndex        =   28
            Top             =   1410
            Width           =   1035
         End
         Begin VB.TextBox txtDB_StockBultos 
            DataField       =   "StockBultos"
            DataSource      =   "adoArticulos"
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
            TabIndex        =   26
            Top             =   1410
            Width           =   1035
         End
         Begin VB.Frame Frame3 
            Height          =   1125
            Left            =   150
            TabIndex        =   5
            Top             =   7020
            Width           =   6735
            Begin VB.CommandButton cmdModificar 
               Caption         =   "Modificar"
               Height          =   795
               Left            =   1447
               Picture         =   "frmStock.frx":001B
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdGuardar 
               Caption         =   "Guardar"
               Height          =   795
               Left            =   2774
               Picture         =   "frmStock.frx":045D
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdCancelar 
               Caption         =   "Cancelar"
               Height          =   795
               Left            =   4101
               Picture         =   "frmStock.frx":089F
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdSalir 
               Caption         =   "Salir"
               Height          =   795
               Left            =   5430
               Picture         =   "frmStock.frx":0CE1
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   210
               Width           =   1155
            End
         End
         Begin VB.TextBox txtDB_IdArticulo 
            DataField       =   "IdArticulo"
            DataSource      =   "adoArticulos"
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
            TabIndex        =   4
            Top             =   570
            Width           =   2115
         End
         Begin VB.TextBox txtDB_Descripcion 
            DataField       =   "Descripcion"
            DataSource      =   "adoArticulos"
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
            Left            =   2280
            TabIndex        =   3
            Top             =   570
            Width           =   4605
         End
         Begin VB.ComboBox cmbED_Motivo 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
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
            TabIndex        =   2
            Top             =   2250
            Width           =   3555
         End
         Begin VB.Label Label4 
            Caption         =   "Bultos/Unidades nuevas"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2340
            TabIndex        =   33
            Top             =   1170
            Width           =   2085
         End
         Begin VB.Label Label3 
            Caption         =   "Bultos/Unidades diferencia"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4560
            TabIndex        =   30
            Top             =   1170
            Width           =   2025
         End
         Begin VB.Label Label1 
            Caption         =   "Bultos/Unidades actuales"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1170
            Width           =   2085
         End
         Begin VB.Label lbl_Activo 
            Caption         =   "Motivo"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   90
            TabIndex        =   12
            Top             =   2010
            Width           =   2055
         End
         Begin VB.Label lbl_IdArticulo 
            Caption         =   "Código"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   330
            Width           =   2055
         End
         Begin VB.Label lbl_IdSegunProveedor 
            Caption         =   "Descripción"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2340
            TabIndex        =   10
            Top             =   330
            Width           =   2115
         End
      End
   End
   Begin MSAdodcLib.Adodc adoArticulos 
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
      RecordSource    =   "Articulos"
      Caption         =   "adoArticulos"
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
   Begin MSAdodcLib.Adodc adoMovimientoStock 
      Height          =   330
      Left            =   8700
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
      RecordSource    =   "MovimientoStock"
      Caption         =   "adoMovimientoStock"
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
      Caption         =   "CONTROL DE STOCK"
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
      TabIndex        =   25
      Top             =   0
      Width           =   8625
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vsOrden As String ' Establace el orden de los registros
Dim vsTAI As String ' Lista Todos, Activos o Inactivos
Dim vsASCDES As String ' Establece si el orden es ASCendente o DEScendente
Dim vsCampo As String ' Establece que campo se usara para el filtro
Dim vsFiltro As String ' Filtra los canales
Dim vsRegistrosArticulos As Integer ' Cuantos registros tiene la tabla de Articulos
Dim vsMovStock As Double ' Número de movimiento de stock

Private Sub Form_Load()
  
  lblTitulo.Top = 0
  lblTitulo.Left = 0
  lblTitulo.Width = Screen.Width
  frcArticulos.Left = (Screen.Width - frcArticulos.Width) / 2

  cmbED_Motivo.AddItem "Compra de mercadería"
  cmbED_Motivo.AddItem "Movimiento InterDeposito"
  cmbED_Motivo.AddItem "Control de stock"
  cmbED_Motivo.AddItem "Rotura/Perdida/Cambio"
  cmbED_Motivo.Text = cmbED_Motivo.List(0)
  
  cmbCamposOrden.AddItem "Código"
  cmbCamposOrden.AddItem "Código según poveedor"
  cmbCamposOrden.AddItem "Descripción"
  cmbCamposOrden.AddItem "Precio de costo"
  cmbCamposOrden.AddItem "Precio de venta"
  cmbCamposOrden.AddItem "Activo"
  cmbCamposOrden.Text = cmbCamposOrden.List(0)
  
  cmbCamposFiltro.AddItem "Código"
  cmbCamposFiltro.AddItem "Código según proveedor"
  cmbCamposFiltro.AddItem "Descripción"
  cmbCamposFiltro.AddItem "Precio de costo"
  cmbCamposFiltro.AddItem "Precio de venta"
  cmbCamposFiltro.AddItem "Activo"
  cmbCamposFiltro.Text = cmbCamposFiltro.List(0)
  
  optTodos.Value = True
  
  Call proListadoFull
  
  If vsRegistrosArticulos = 0 Then
    Call proActivarBotones(False, False, False, True)
  Else
    Call proActivarBotones(True, False, False, True)
  End If

End Sub

Private Sub Form_Activate()

  dtgArticulos.SetFocus

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
    Case "Código": vsOrden = "IdArticulo"
    Case "Código según proveedor": vsOrden = "IdSegunProveedor"
    Case "Descripción": vsOrden = "Descripcion"
    Case "Precio de costo": vsOrden = "PrecioCosto"
    Case "Precio de venta": vsOrden = "PrecioVenta"
    Case "Activo": vsOrden = "Activo"
  End Select

  vsTAI = "Todos"
  vsASCDES = "ASC"
  vsCampo = "IdArticulo"
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
  
  dtgArticulos.SetFocus
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)

End Sub

Private Sub cmbCamposFiltro_Click()

  Select Case cmbCamposFiltro
    Case "Código": vsCampo = "IdArticulo"
    Case "Código según proveedor": vsCampo = "IdSegunProveedor"
    Case "Descripción": vsCampo = "Descripcion"
    Case "Precio de costo": vsCampo = "PrecioCosto"
    Case "Precio de venta": vsCampo = "PrecioVenta"
    Case "Activo": vsCampo = "Activo"
  End Select

  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub txtFiltro_Change()
  
  vsFiltro = txtFiltro.Text
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub txtNuevoStockBultos_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    
    If Not IsNumeric(txtNuevoStockBultos.Text) Then
      MsgBox "Debe introducir un número", vbExclamation, "Dato erroneo"
      txtNuevoStockBultos.Text = 0
      txtNuevoStockBultos.SelStart = 0
      txtNuevoStockBultos.SelLength = Len(txtNuevoStockBultos.Text)
      Exit Sub
    End If
    
    If txtNuevoStockBultos.Text < 0 Then
      MsgBox "Debe introducir un número mayor que cero", vbExclamation, "Dato erroneo"
      txtNuevoStockBultos.Text = 0
      txtNuevoStockBultos.SelStart = 0
      txtNuevoStockBultos.SelLength = Len(txtNuevoStockBultos.Text)
      Exit Sub
    End If
    
    txtDiferenciaBultos.Text = txtNuevoStockBultos - txtDB_StockBultos.Text
    
    txtNuevoStockUnidades.Text = 0
    txtNuevoStockUnidades.SelStart = 0
    txtNuevoStockUnidades.SelLength = Len(txtNuevoStockUnidades.Text)
    txtNuevoStockUnidades.SetFocus
  
  End If

End Sub

Private Sub txtNuevoStockUnidades_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    
    If Not IsNumeric(txtNuevoStockUnidades.Text) Then
      MsgBox "Debe introducir un número", vbExclamation, "Dato erroneo"
      txtNuevoStockUnidades.Text = 0
      txtNuevoStockUnidades.SelStart = 0
      txtNuevoStockUnidades.SelLength = Len(txtNuevoStockUnidades.Text)
      Exit Sub
    End If
    
    If txtNuevoStockUnidades.Text < 0 Then
      MsgBox "Debe introducir un número mayor que cero", vbExclamation, "Dato erroneo"
      txtNuevoStockUnidades.Text = 0
      txtNuevoStockUnidades.SelStart = 0
      txtNuevoStockUnidades.SelLength = Len(txtNuevoStockUnidades.Text)
      Exit Sub
    End If
    
    txtDiferenciaUnidades.Text = txtNuevoStockUnidades - txtDB_StockUnidades.Text
    
    cmbED_Motivo.SetFocus
    Call proActivarBotones(False, True, True, False)
  
  End If

End Sub

Private Sub cmdModificar_Click()
    
  Call proADControlGrid
  
  txtNuevoStockBultos.Enabled = True
  txtNuevoStockUnidades.Enabled = True
  txtDiferenciaBultos.Enabled = True
  txtDiferenciaUnidades.Enabled = True
  cmbED_Motivo.Enabled = True
  
  txtNuevoStockBultos.Text = 0
  txtNuevoStockBultos.SelStart = 0
  txtNuevoStockBultos.SelLength = Len(txtNuevoStockBultos.Text)
  txtNuevoStockBultos.SetFocus
    
  Call proActivarBotones(False, False, True, False)
   
End Sub

Private Sub cmdGuardar_Click()

  adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [IdArticulo]=" & txtDB_IdArticulo.Text & " ORDER BY [IdArticulo]"
  adoArticulos.CommandType = adCmdText
  adoArticulos.Refresh
  
  If adoArticulos.Recordset.RecordCount <> 0 Then
    adoArticulos.Recordset![StockBultos] = txtNuevoStockBultos.Text
    adoArticulos.Recordset![StockUnidades] = txtNuevoStockUnidades.Text
    adoArticulos.Recordset.Update
    adoArticulos.Refresh
  Else
    MsgBox "Existe un error con el stock de este artículo. Consulte a su administrador", vbCritical, "Error importante"
    Exit Sub
  End If
  
  adoMovimientoStock.RecordSource = "SELECT * FROM [MovimientoStock] ORDER BY [IdMovStock]"
  adoMovimientoStock.CommandType = adCmdText
  adoMovimientoStock.Refresh
  vsMovStock = adoMovimientoStock.Recordset.RecordCount
  adoMovimientoStock.Recordset.AddNew
  adoMovimientoStock.Recordset![IdMovStock] = vsMovStock + 1
  adoMovimientoStock.Recordset![Fecha] = Date
  adoMovimientoStock.Recordset![Hora] = Time
  adoMovimientoStock.Recordset![Motivo] = UCase(cmbED_Motivo.Text)
  adoMovimientoStock.Recordset![IdArticulo] = txtDB_IdArticulo.Text
  adoMovimientoStock.Recordset![Descripcion] = txtDB_Descripcion.Text
  adoMovimientoStock.Recordset![StockBultos] = txtNuevoStockBultos.Text
  adoMovimientoStock.Recordset![StockUnidades] = txtNuevoStockUnidades.Text
  adoMovimientoStock.Recordset.Update
  adoMovimientoStock.Refresh
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
  Call proActivarBotones(True, False, False, True)
  
  txtNuevoStockBultos.Text = ""
  txtNuevoStockUnidades.Text = ""
  txtNuevoStockBultos.Enabled = False
  txtNuevoStockUnidades.Enabled = False
  txtDiferenciaBultos.Text = ""
  txtDiferenciaUnidades.Text = ""
  txtDiferenciaBultos.Enabled = False
  txtDiferenciaUnidades.Enabled = False
  cmbED_Motivo.Enabled = False
  
End Sub

Private Sub cmdCancelar_Click()
  
  optTodos.Value = True
  cmbCamposOrden.Text = cmbCamposOrden.List(0)
  vsASCDES = "ASC"
  cmdASCDES.Caption = "ASC"
  cmbCamposFiltro.Text = cmbCamposFiltro.List(0)
  txtFiltro.Text = ""
  
  Call proListadoFull
  
  txtNuevoStockBultos.Text = ""
  txtNuevoStockUnidades.Text = ""
  txtNuevoStockBultos.Enabled = False
  txtNuevoStockUnidades.Enabled = False
  txtDiferenciaBultos.Text = ""
  txtDiferenciaUnidades.Text = ""
  txtDiferenciaBultos.Enabled = False
  txtDiferenciaUnidades.Enabled = False
  cmbED_Motivo.Enabled = False
  
  Call proActivarBotones(True, False, False, True)

End Sub

Private Sub cmdSalir_Click()

  Unload Me
  
End Sub

Private Sub proListadoFull()

  adoArticulos.RecordSource = "SELECT * FROM [Articulos] ORDER BY [IdArticulo]"
  adoArticulos.CommandType = adCmdText
  adoArticulos.Refresh
  vsRegistrosArticulos = adoArticulos.Recordset.RecordCount
  
End Sub

Private Sub proActivarBotones(ByVal M As Boolean, ByVal G As Boolean, ByVal C As Boolean, ByVal S As Boolean)

  cmdModificar.Enabled = M
  cmdGuardar.Enabled = G
  cmdCancelar.Enabled = C
  cmdSalir.Enabled = S

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

Private Sub proListarRegistros(ByVal Estado As String, ByVal Orden As String, ByVal ASCDES As String, ByVal Campo As String, ByVal Filtro As String)

  If Estado = "Todos" Then
    adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [" & Campo & "] LIKE '%" & Filtro & "%' ORDER BY " & Orden & " " & ASCDES
  Else
    adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [Activo]= '" & Estado & "' ORDER BY [" & Orden & "]" & ASCDES
  End If
  
  adoArticulos.CommandType = adCmdText
  adoArticulos.Refresh
  
End Sub
