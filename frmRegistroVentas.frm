VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmRegistroVentas 
   Caption         =   "Registro de ventas"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crcPresupuesto 
      Left            =   19110
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Program Files\DEGPLUS\Reports\presupuestox.rpt"
   End
   Begin VB.Frame frcVentas 
      Height          =   8595
      Left            =   30
      TabIndex        =   5
      Top             =   630
      Width           =   20265
      Begin VB.Frame Frame4 
         Height          =   8295
         Left            =   8760
         TabIndex        =   17
         Top             =   150
         Width           =   11385
         Begin VB.Frame Frame5 
            Height          =   795
            Left            =   90
            TabIndex        =   18
            Top             =   120
            Width           =   11175
            Begin VB.OptionButton optTodos 
               Caption         =   "Todos"
               Height          =   285
               Left            =   240
               TabIndex        =   25
               Top             =   270
               Width           =   855
            End
            Begin VB.OptionButton optActivos 
               Caption         =   "Activos"
               Height          =   285
               Left            =   1140
               TabIndex        =   24
               Top             =   270
               Width           =   915
            End
            Begin VB.OptionButton optInactivos 
               Caption         =   "Inactivos"
               Height          =   285
               Left            =   2070
               TabIndex        =   23
               Top             =   270
               Width           =   1125
            End
            Begin VB.ComboBox cmbCamposOrden 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   4290
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   270
               Width           =   1725
            End
            Begin VB.ComboBox cmbCamposFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   7650
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   270
               Width           =   1725
            End
            Begin VB.TextBox txtFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   345
               Left            =   9420
               TabIndex        =   20
               Top             =   270
               Width           =   1575
            End
            Begin VB.CommandButton cmdASCDES 
               Caption         =   "ASC"
               Height          =   285
               Left            =   6060
               TabIndex        =   19
               Top             =   270
               Width           =   735
            End
            Begin VB.Label Label6 
               Caption         =   "Orden"
               Height          =   255
               Left            =   3780
               TabIndex        =   27
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Label7 
               Caption         =   "Filtro"
               Height          =   255
               Left            =   7200
               TabIndex        =   26
               Top             =   300
               Width           =   765
            End
         End
         Begin MSDataGridLib.DataGrid dtgComprobantes 
            Bindings        =   "frmRegistroVentas.frx":0000
            Height          =   3345
            Left            =   90
            TabIndex        =   28
            Top             =   990
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   5900
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
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
            Caption         =   "Listado de comprobantes"
            ColumnCount     =   18
            BeginProperty Column00 
               DataField       =   "Fecha"
               Caption         =   "Fecha"
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
               DataField       =   "TipoComprobante"
               Caption         =   "Tipo Comprobante"
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
               DataField       =   "IdVenta"
               Caption         =   "Nro Comprobante"
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
            BeginProperty Column04 
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
            BeginProperty Column05 
               DataField       =   "ImporteComprobante"
               Caption         =   "Importe total"
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
               DataField       =   "BultosComprobante"
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
            BeginProperty Column07 
               DataField       =   "UnidadesComprobante"
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
            BeginProperty Column08 
               DataField       =   "ImporteComprobante"
               Caption         =   "ImporteComprobante"
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
            BeginProperty Column09 
               DataField       =   "BultosComprobante"
               Caption         =   "BultosComprobante"
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
            BeginProperty Column10 
               DataField       =   "UnidadesComprobante"
               Caption         =   "UnidadesComprobante"
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
            BeginProperty Column11 
               DataField       =   "TipoRegistro"
               Caption         =   "TipoRegistro"
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
            BeginProperty Column12 
               DataField       =   "IdArticulo"
               Caption         =   "IdArticulo"
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
            BeginProperty Column13 
               DataField       =   "Descripcion"
               Caption         =   "Descripcion"
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
            BeginProperty Column14 
               DataField       =   "BultosCantidad"
               Caption         =   "BultosCantidad"
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
            BeginProperty Column15 
               DataField       =   "UnidadesCantidad"
               Caption         =   "UnidadesCantidad"
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
            BeginProperty Column16 
               DataField       =   "PrecioVenta"
               Caption         =   "PrecioVenta"
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
            BeginProperty Column17 
               DataField       =   "Subtotal"
               Caption         =   "Subtotal"
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
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1454,74
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1395,213
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   705,26
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   3344,882
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column07 
                  Alignment       =   2
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1484,787
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1395,213
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1635,024
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column14 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column15 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1319,811
               EndProperty
               BeginProperty Column16 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1065,26
               EndProperty
               BeginProperty Column17 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1065,26
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dtgItems 
            Bindings        =   "frmRegistroVentas.frx":001D
            Height          =   3795
            Left            =   90
            TabIndex        =   51
            Top             =   4380
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   6694
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
            Caption         =   "Items de comprobante"
            ColumnCount     =   18
            BeginProperty Column00 
               DataField       =   "IdVenta"
               Caption         =   "Nro Comprobante"
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
               DataField       =   "BultosCantidad"
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
            BeginProperty Column04 
               DataField       =   "UnidadesCantidad"
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
            BeginProperty Column05 
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
            BeginProperty Column06 
               DataField       =   "Subtotal"
               Caption         =   "SubTotal"
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
               DataField       =   "Sucursal"
               Caption         =   "Sucursal"
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
            BeginProperty Column08 
               DataField       =   "ImporteComprobante"
               Caption         =   "ImporteComprobante"
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
            BeginProperty Column09 
               DataField       =   "BultosComprobante"
               Caption         =   "BultosComprobante"
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
            BeginProperty Column10 
               DataField       =   "UnidadesComprobante"
               Caption         =   "UnidadesComprobante"
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
            BeginProperty Column11 
               DataField       =   "TipoRegistro"
               Caption         =   "TipoRegistro"
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
            BeginProperty Column12 
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
            BeginProperty Column13 
               DataField       =   "Descripcion"
               Caption         =   "Descripcion"
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
            BeginProperty Column14 
               DataField       =   "BultosCantidad"
               Caption         =   "BultosCantidad"
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
            BeginProperty Column15 
               DataField       =   "UnidadesCantidad"
               Caption         =   "UnidadesCantidad"
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
            BeginProperty Column16 
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
            BeginProperty Column17 
               DataField       =   "Subtotal"
               Caption         =   "Subtotal"
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
                  ColumnWidth     =   1395,213
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   1305,071
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   3344,882
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1200,189
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   1695,118
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   629,858
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1484,787
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1395,213
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1635,024
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column14 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column15 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1319,811
               EndProperty
               BeginProperty Column16 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1065,26
               EndProperty
               BeginProperty Column17 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1065,26
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   8295
         Left            =   150
         TabIndex        =   6
         Top             =   150
         Width           =   8565
         Begin VB.CheckBox chkImprimir 
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2100
            TabIndex        =   56
            Top             =   6660
            Value           =   1  'Checked
            Width           =   1305
         End
         Begin VB.TextBox txtUnidades 
            Alignment       =   1  'Right Justify
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
            Left            =   5730
            TabIndex        =   55
            Top             =   6630
            Width           =   555
         End
         Begin VB.TextBox txtED_CantidadUnidades 
            Alignment       =   1  'Right Justify
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
            Left            =   4890
            TabIndex        =   53
            Text            =   "0"
            Top             =   2250
            Width           =   705
         End
         Begin VB.TextBox txtBultos 
            Alignment       =   1  'Right Justify
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
            Left            =   5100
            TabIndex        =   48
            Top             =   6630
            Width           =   555
         End
         Begin VB.TextBox txtImporteTotal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
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
            Left            =   7080
            TabIndex        =   47
            Text            =   "9999388,34"
            Top             =   6630
            Width           =   1245
         End
         Begin MSFlexGridLib.MSFlexGrid fgrDetalle 
            Height          =   3855
            Left            =   90
            TabIndex        =   46
            Top             =   2760
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   6800
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            SelectionMode   =   1
         End
         Begin VB.CommandButton cmdNuevaLinea 
            BeginProperty Font 
               Name            =   "Bernard MT Condensed"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   7950
            Picture         =   "frmRegistroVentas.frx":0036
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2220
            UseMaskColor    =   -1  'True
            Width           =   450
         End
         Begin VB.TextBox txtSubTotal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
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
            Left            =   6570
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "0"
            Top             =   2250
            Width           =   1335
         End
         Begin VB.TextBox txtPrecioVenta 
            Alignment       =   1  'Right Justify
            DataField       =   "PrecioBDE"
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
            Height          =   450
            Left            =   5670
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0"
            Top             =   2250
            Width           =   825
         End
         Begin VB.TextBox txtED_CantidadBultos 
            Alignment       =   1  'Right Justify
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
            Left            =   4110
            TabIndex        =   3
            Top             =   2250
            Width           =   705
         End
         Begin VB.TextBox txtDescripcion 
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
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   2250
            Width           =   2985
         End
         Begin VB.TextBox txtED_IdArticulo 
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
            Left            =   90
            TabIndex        =   2
            Top             =   2250
            Width           =   885
         End
         Begin VB.TextBox txtSucursal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "9999#"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
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
            Left            =   4350
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1380
            Width           =   795
         End
         Begin VB.TextBox txtLetra 
            Alignment       =   2  'Center
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
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1380
            Width           =   675
         End
         Begin VB.TextBox txtTipoComprobante 
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
            TabIndex        =   32
            Top             =   1380
            Width           =   3405
         End
         Begin VB.TextBox txtCanal 
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
            Left            =   6090
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   570
            Width           =   2355
         End
         Begin VB.TextBox txtRazonSocial 
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
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   570
            Width           =   4485
         End
         Begin VB.Frame Frame3 
            Height          =   1125
            Left            =   90
            TabIndex        =   8
            Top             =   7080
            Width           =   8385
            Begin VB.CommandButton cmdModificar 
               Caption         =   "Modificar"
               Height          =   795
               Left            =   3540
               Picture         =   "frmRegistroVentas.frx":06A0
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdAnular 
               Caption         =   "Anular"
               Height          =   795
               Left            =   2220
               Picture         =   "frmRegistroVentas.frx":0AE2
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdGuardar 
               Caption         =   "Guardar"
               Height          =   795
               Left            =   900
               Picture         =   "frmRegistroVentas.frx":0F24
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdCancelar 
               Caption         =   "Cancelar"
               Height          =   795
               Left            =   4860
               Picture         =   "frmRegistroVentas.frx":1366
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdSalir 
               Caption         =   "Salir"
               Height          =   795
               Left            =   6180
               Picture         =   "frmRegistroVentas.frx":17A8
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   210
               Width           =   1155
            End
         End
         Begin VB.TextBox txtIdVenta 
            Alignment       =   1  'Right Justify
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
            Left            =   5250
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1380
            Width           =   3165
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
            Left            =   90
            TabIndex        =   1
            Top             =   570
            Width           =   1365
         End
         Begin VB.Label Label14 
            Caption         =   "Unidades"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4920
            TabIndex        =   54
            Top             =   2010
            Width           =   735
         End
         Begin VB.Label lblStockActual 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   90
            TabIndex        =   52
            Top             =   6660
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "Bultos/Unidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3630
            TabIndex        =   50
            Top             =   6720
            Width           =   1785
         End
         Begin VB.Label Label5 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6570
            TabIndex        =   49
            Top             =   6720
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "SubTotal"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6600
            TabIndex        =   44
            Top             =   2010
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Precio"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5700
            TabIndex        =   42
            Top             =   2010
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Bultos"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4140
            TabIndex        =   41
            Top             =   2010
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Descripción"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   39
            Top             =   2010
            Width           =   2115
         End
         Begin VB.Label Label8 
            Caption         =   "Id Artículo"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   2010
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Sucursal"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4380
            TabIndex        =   36
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Letra"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3630
            TabIndex        =   34
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Número de comprobante"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   33
            Top             =   1140
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Canal"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6120
            TabIndex        =   31
            Top             =   330
            Width           =   2235
         End
         Begin VB.Label lbl_Descripcion 
            Caption         =   "Comprobante"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1140
            Width           =   1305
         End
         Begin VB.Label lbl_IdArticulo 
            Caption         =   "Código de cliente"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   330
            Width           =   1275
         End
         Begin VB.Label lbl_IdSegunProveedor 
            Caption         =   "Razón social"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   14
            Top             =   330
            Width           =   2115
         End
      End
   End
   Begin MSAdodcLib.Adodc adoArticulos 
      Height          =   330
      Left            =   8640
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
   Begin MSAdodcLib.Adodc adoClientes 
      Height          =   330
      Left            =   8640
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
   Begin MSAdodcLib.Adodc adoRegistroVentas 
      Height          =   330
      Left            =   11250
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
      RecordSource    =   "RegistroVentas"
      Caption         =   "adoRegistroVentas"
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
      Left            =   11250
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
   Begin MSAdodcLib.Adodc adoRV_Cabecera 
      Height          =   330
      Left            =   13860
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
      RecordSource    =   "RegistroVentas"
      Caption         =   "adoRV_Cabecera"
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
   Begin MSAdodcLib.Adodc adoRV_Item 
      Height          =   330
      Left            =   13860
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
      RecordSource    =   "RegistroVentas"
      Caption         =   "adoRV_Item"
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
   Begin MSAdodcLib.Adodc adoImpresion 
      Height          =   330
      Left            =   16470
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
      RecordSource    =   "RegistroVentas"
      Caption         =   "adoImpresion"
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
   Begin MSAdodcLib.Adodc adoBolsa 
      Height          =   330
      Left            =   16470
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
      RecordSource    =   "Bolsa"
      Caption         =   "adoBolsa"
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
   Begin MSAdodcLib.Adodc adoRV_Devolucion 
      Height          =   330
      Left            =   15660
      Top             =   9330
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
      RecordSource    =   "RegistroVentas"
      Caption         =   "adoRV_Devolucion"
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
      Caption         =   "REGISTRO DE VENTAS"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8625
   End
End
Attribute VB_Name = "frmRegistroVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vsCantidadDeRegistros As Double ' Registra la cantidad de registros en la tabla de ventas
Dim vsUltimoRegistro As Double ' Registra el IdVenta del ultimo registro de la tabla ordenada de menor a mayor
Dim vsComprobante As String ' Tipo de comprobante
Dim vsLetra As String ' Letra de comprobante
Dim vsSucursal As Byte ' Sucursal de facturacion
Dim vsIdVenta As Double ' Número de comprobante (Presupuesto)
Dim vsIdNotaDeCredito As Double ' Número de comprobante (Nota de crédito)
Dim vsPresupuestoActivo As Double ' Número de presupuesto buscado
Dim vsPeriodo As String ' Almacena el periodo(Fecha) de la venta
Dim vsImporteTotal As Single ' Almacena el importe total del comprobante
Dim vsBultos As Integer ' Almacena el total de bultos del comprobante
Dim vsUnidades As Integer ' Almacena el total de unidades del comprobante
Dim vsLineasTotal As Integer ' Almacena cuantas lineas tiene el grid
Dim vsLineaActiva As Integer ' Indica cual es la linea activa
Dim vsI As Integer ' Contador común
Dim vsMovStock As Double ' Número de movimiento de stock
Dim vsTAI As String ' Lista Todos, Activos o Inactivos
Dim vsOrden As String ' Establace el orden de los registros
Dim vsASCDES As String ' Establece si el orden es ASCendente o DEScendente
Dim vsCampo As String ' Establece que campo se usara para el filtro
Dim vsFiltro As String ' Filtra los canales
Dim vsUltimaVenta As Double ' Indica el número de comprobante de la última venta
Dim vsUltimaDevolucion As Double ' Indica el número de comprobante de la última devolución
Dim cReporte As String ' Almacena el query del reporte
Dim vsTipoComprobante As String ' Almacena que tipo de comprobante tiene que filtrar cuando se hace click en el listado de comprobantes

Private Sub Form_Load()
  
  lblTitulo.Top = 0
  lblTitulo.Left = 0
  lblTitulo.Width = Screen.Width
  frcVentas.Left = (Screen.Width - frcVentas.Width) / 2

  Call proLimpiaCampos
  
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
  
  vsComprobante = "Presupuesto"
  vsLetra = "X"
  vsSucursal = 1
  
  txtTipoComprobante.Text = vsComprobante
  txtLetra.Text = vsLetra
  txtSucursal.Text = vsSucursal
  txtIdVenta.Text = vsIdVenta

  fgrDetalle.ColWidth(0) = 870
  fgrDetalle.ColWidth(1) = 3090
  fgrDetalle.ColWidth(2) = 785
  fgrDetalle.ColWidth(3) = 785
  fgrDetalle.ColWidth(4) = 930
  fgrDetalle.ColWidth(5) = 1350
  
  vsLineasTotal = 1
  fgrDetalle.Rows = vsLineasTotal
  vsLineaActiva = 0
  vsImporteTotal = 0
  vsBultos = 0
  vsUnidades = 0
   
  vsTAI = "Todos"
  vsOrden = "IdVenta"
  vsASCDES = "DESC"
  vsCampo = "IdVenta"
  vsFiltro = ""
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub Form_GotFocus()

  If vsVieneDe = "CLIENTES" Then
 
    txtED_IdCliente.Text = vsReturnIdCliente
    txtRazonSocial.Text = vsReturnRazonSocial
    txtCanal.Text = vsReturnCanal
    
    txtED_IdArticulo.SetFocus
  
  End If

  If vsVieneDe = "ARTICULOS" Then
 
    txtED_IdArticulo.Text = vsReturnIdArticulo
    txtDescripcion.Text = vsReturnDescripcion
    txtPrecioVenta.Text = vsReturnPrecioVenta
    
    txtED_CantidadBultos.Text = 1
    txtED_CantidadBultos.SelStart = 0
    txtED_CantidadBultos.SelLength = Len(txtED_CantidadBultos.Text)
    txtED_CantidadBultos.SetFocus
  
  End If

End Sub

Private Sub Form_Activate()

  Call proActualizarVentas

End Sub

Private Sub txtED_IdCliente_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
  
    txtRazonSocial.Text = ""
    txtCanal.Text = ""
    
    If txtED_IdCliente.Text = "" Then
      MsgBox "Debe escribir el código del cliente o las letras para la busqueda sensible", vbInformation, "Dato esperado"
      txtED_IdCliente.SetFocus
      Exit Sub
    End If
    
    If IsNumeric(txtED_IdCliente.Text) = True Then ' Por aquí busca el cliente si se introduce un número
      adoClientes.RecordSource = "SELECT * FROM [Clientes] WHERE [IdCliente]= " & txtED_IdCliente.Text & " ORDER BY [IdCliente]"
      adoClientes.CommandType = adCmdText
      adoClientes.Refresh
      
      If adoClientes.Recordset.RecordCount = 0 Then
        MsgBox "El código ingresado no corresponde a ningún cliente en la base de datos", vbExclamation, "No encontrado"
        txtED_IdCliente.SelStart = 0
        txtED_IdCliente.SelLength = Len(txtED_IdCliente.Text)
        txtED_IdCliente.SetFocus
        Exit Sub
      Else
        txtRazonSocial.Text = adoClientes.Recordset![RazonSocial]
        txtCanal.Text = adoClientes.Recordset![Canal]
        txtED_IdArticulo.SetFocus
      End If
      
    Else ' Por aquí se busca el cliente si se introducen letras
      
      vsQueryRazonSocial = txtED_IdCliente.Text
      
      frmListaClientes.Show
      
    End If
    
    cmdCancelar.Enabled = True
    txtED_IdCliente.Locked = True
     
  End If
  
End Sub

Private Sub txtED_IdArticulo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
  
    txtDescripcion.Text = ""
    txtPrecioVenta.Text = ""
    
    If txtED_IdArticulo.Text = "" Then
      MsgBox "Debe escribir el código del artículo o las letras para la busqueda sensible", vbInformation, "Dato esperado"
      txtED_IdArticulo.SetFocus
      Exit Sub
    End If
    
    If IsNumeric(txtED_IdArticulo.Text) = True Then ' Por aquí busca el artículo si se introduce un número
      adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [Idarticulo]= " & txtED_IdArticulo.Text & " ORDER BY [IdArticulo]"
      adoArticulos.CommandType = adCmdText
      adoArticulos.Refresh
      
      If adoArticulos.Recordset.RecordCount = 0 Then
        MsgBox "El código ingresado no corresponde a ningún artículo en la base de datos", vbExclamation, "No encontrado"
        txtED_IdArticulo.SelStart = 0
        txtED_IdArticulo.SelLength = Len(txtED_IdArticulo.Text)
        txtED_IdArticulo.SetFocus
        Exit Sub
      Else
        txtDescripcion.Text = adoArticulos.Recordset![Descripcion]
        txtPrecioVenta.Text = adoArticulos.Recordset![PrecioBDE]
        lblStockActual.Caption = "Stock: " & adoArticulos.Recordset![StockBultos] & " | " & adoArticulos.Recordset![StockUnidades]
        txtED_CantidadBultos.Text = 1
        txtED_CantidadBultos.SelStart = 0
        txtED_CantidadBultos.SelLength = Len(txtED_CantidadBultos.Text)
        txtED_CantidadBultos.SetFocus
      End If
      
    Else ' Por aquí se busca el artículo si se introducen letras
      
      vsQueryDescripcion = txtED_IdArticulo.Text
      
      frmListaArticulos.Show
      
    End If
    
  End If

End Sub

Private Sub txtED_CantidadBultos_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    If Not IsNumeric(txtED_CantidadBultos.Text) Then
      MsgBox "La cantidad ingresada debe ser un número", vbExclamation, "Dato erroneo"
      txtED_CantidadBultos.Text = 1
      txtED_CantidadBultos.SelStart = 0
      txtED_CantidadBultos.SelLength = Len(txtED_CantidadBultos.Text)
      txtED_CantidadBultos.SetFocus
      Exit Sub
    End If
    If txtED_CantidadBultos.Text < 0 Then
      MsgBox "La cantidad ingresada no es permitida", vbExclamation, "Dato erroneo"
      txtED_CantidadBultos.Text = 1
      txtED_CantidadBultos.SelStart = 0
      txtED_CantidadBultos.SelLength = Len(txtED_CantidadBultos.Text)
      txtED_CantidadBultos.SetFocus
      Exit Sub
    End If
    
    txtED_CantidadUnidades.Text = 0
    txtED_CantidadUnidades.SelStart = 0
    txtED_CantidadUnidades.SelLength = Len(txtED_CantidadUnidades.Text)
    txtED_CantidadUnidades.SetFocus
  
  End If

End Sub

Private Sub txtED_CantidadUnidades_KeyPress(KeyAscii As Integer)


  If KeyAscii = 13 Then
    If Not IsNumeric(txtED_CantidadUnidades.Text) Then
      MsgBox "La cantidad ingresada debe ser un número", vbExclamation, "Dato erroneo"
      txtED_CantidadUnidades.Text = 1
      txtED_CantidadUnidades.SelStart = 0
      txtED_CantidadUnidades.SelLength = Len(txtED_CantidadUnidades.Text)
      txtED_CantidadUnidades.SetFocus
      Exit Sub
    End If
    If (txtED_CantidadBultos.Text + txtED_CantidadUnidades.Text) < 1 Then
      MsgBox "Las cantidades ingresadas no son permitidas", vbExclamation, "Dato erroneo"
      txtED_CantidadUnidades = 1
      txtED_CantidadUnidades.SelStart = 0
      txtED_CantidadUnidades.SelLength = Len(txtED_CantidadUnidades.Text)
      txtED_CantidadUnidades.SetFocus
      Exit Sub
    End If
    
    txtSubTotal.Text = Round((txtED_CantidadBultos.Text * txtPrecioVenta.Text) + (txtED_CantidadUnidades.Text * (txtPrecioVenta.Text / adoArticulos.Recordset![UxB])), 2)
        
    cmdNuevaLinea.Enabled = True
    cmdNuevaLinea.SetFocus
  End If

End Sub

Private Sub cmdNuevaLinea_Click()
  
  fgrDetalle.Row = vsLineaActiva
  fgrDetalle.Col = 0
  fgrDetalle.Text = txtED_IdArticulo.Text
  
  fgrDetalle.Col = 1
  fgrDetalle.CellAlignment = flexAlignLeftCenter
  fgrDetalle.Text = txtDescripcion.Text
  
  fgrDetalle.Col = 2
  fgrDetalle.Text = txtED_CantidadBultos.Text
  
  fgrDetalle.Col = 3
  fgrDetalle.Text = txtED_CantidadUnidades.Text
  
  fgrDetalle.Col = 4
  fgrDetalle.Text = txtPrecioVenta.Text
  
  fgrDetalle.Col = 5
  fgrDetalle.Text = txtSubTotal.Text
  
  vsBultos = vsBultos + txtED_CantidadBultos.Text
  txtBultos.Text = vsBultos
  
  vsUnidades = vsUnidades + txtED_CantidadUnidades.Text
  txtUnidades.Text = vsUnidades
  
  vsImporteTotal = vsImporteTotal + txtSubTotal.Text
  txtImporteTotal.Text = vsImporteTotal
  
  vsLineasTotal = vsLineasTotal + 1
  fgrDetalle.Rows = vsLineasTotal
  vsLineaActiva = vsLineaActiva + 1
  
  txtED_IdArticulo.Text = ""
  txtDescripcion.Text = ""
  txtED_CantidadBultos.Text = 1
  txtED_CantidadUnidades.Text = 0
  txtPrecioVenta.Text = 0
  txtSubTotal.Text = 0
  txtED_IdArticulo.SetFocus
  
  If vsLineasTotal > 12 Then
    MsgBox "Se ha llegado a la cantidad máxima de items", vbExclamation, "Límite de items"
    cmdNuevaLinea.Enabled = False
    Exit Sub
  End If
  
  Call proActivarBotones(True, False, False, True, True)
  
End Sub

Private Sub fgrDetalle_KeyPress(KeyAscii As Integer)

  ' Verifica que se haya presionado la tecla Retroceso (BackSpace)
  If KeyAscii = 8 Then
    
    ' Toma el valor de la celda Bultos de la fila seleccionada y lo resta del total de bultos
    fgrDetalle.Row = fgrDetalle.RowSel
    fgrDetalle.Col = 2
    vsBultos = vsBultos - fgrDetalle.Text
    
    ' Toma el valor de la celda Unidades de la fila seleccionada y lo resta del total de unidades
    fgrDetalle.Row = fgrDetalle.RowSel
    fgrDetalle.Col = 3
    vsUnidades = vsUnidades - fgrDetalle.Text
    
    ' Toma el valor de la celda ImporteTotal de la fila seleccionada y lo resta del total de importe
    fgrDetalle.Row = fgrDetalle.RowSel
    fgrDetalle.Col = 5
    vsImporteTotal = vsImporteTotal - fgrDetalle.Text
    
    txtBultos.Text = vsBultos
    txtUnidades.Text = vsUnidades
    txtImporteTotal.Text = vsImporteTotal
    
    fgrDetalle.RemoveItem (fgrDetalle.RowSel)
    
    vsLineasTotal = vsLineasTotal - 1
    fgrDetalle.Rows = vsLineasTotal
    vsLineaActiva = vsLineaActiva - 1
      
    txtED_IdArticulo.Text = ""
    txtDescripcion.Text = ""
    txtED_CantidadBultos.Text = 1
    txtED_CantidadUnidades.Text = 0
    txtPrecioVenta.Text = 0
    txtSubTotal.Text = 0
    txtED_IdArticulo.SetFocus
  End If
  
End Sub

Private Sub cmdGuardar_Click()

  If txtED_IdCliente.Text = "" Then
    MsgBox "El código de cliente debe tener un valor. Por favor complete el dato", vbInformation, "Dato esperado"
    txtED_IdCliente.SetFocus
    Exit Sub
  End If
  
  ' Almacena los datos en la tabla RegistroVentas para el tipo de registro "CABECERA"
  ' Aqui guarda todos los datos de la cabecera del comprobante, por eso en este registro
  ' los campos de Item de ventas estan vacios
  
  adoRegistroVentas.Recordset.AddNew
  adoRegistroVentas.Recordset![IdVenta] = txtIdVenta.Text
  adoRegistroVentas.Recordset![Fecha] = Date
  adoRegistroVentas.Recordset![Hora] = Time
  adoRegistroVentas.Recordset![Periodo] = DatePart("d", Date) & DatePart("m", Date) & DatePart("yyyy", Date)
  adoRegistroVentas.Recordset![IdCliente] = txtED_IdCliente.Text
  adoRegistroVentas.Recordset![RazonSocial] = txtRazonSocial.Text
  adoRegistroVentas.Recordset![TipoComprobante] = txtTipoComprobante.Text
  adoRegistroVentas.Recordset![Sucursal] = txtSucursal.Text
  adoRegistroVentas.Recordset![Letra] = txtLetra.Text
  adoRegistroVentas.Recordset![ImporteComprobante] = CSng(txtImporteTotal.Text)
  adoRegistroVentas.Recordset![BultosComprobante] = txtBultos.Text
  adoRegistroVentas.Recordset![UnidadesComprobante] = txtUnidades.Text
  adoRegistroVentas.Recordset![TipoRegistro] = "CABECERA"
  adoRegistroVentas.Recordset![IdArticulo] = 0
  adoRegistroVentas.Recordset![Descripcion] = ""
  adoRegistroVentas.Recordset![BultosCantidad] = 0
  adoRegistroVentas.Recordset![unidadesCantidad] = 0
  adoRegistroVentas.Recordset![PrecioVenta] = 0
  adoRegistroVentas.Recordset![SubTotal] = 0
  adoRegistroVentas.Recordset.Update
  
  ' Almacena los datos en la tabla VentaItem y luego en la tabla MovimientoStock en forma intercalada y
  ' actualiza el stock del artículo
  
  adoMovimientoStock.RecordSource = "SELECT * FROM [MovimientoStock] ORDER BY [IdMovStock]"
  adoMovimientoStock.CommandType = adCmdText
  adoMovimientoStock.Refresh
  vsMovStock = adoMovimientoStock.Recordset.RecordCount
  
  For vsI = 0 To vsLineasTotal - 2
    adoRegistroVentas.Recordset.AddNew
    adoRegistroVentas.Recordset![IdVenta] = txtIdVenta.Text
    adoRegistroVentas.Recordset![Fecha] = Date
    adoRegistroVentas.Recordset![Hora] = Time
    adoRegistroVentas.Recordset![IdCliente] = txtED_IdCliente.Text
    adoRegistroVentas.Recordset![RazonSocial] = txtRazonSocial.Text
    adoRegistroVentas.Recordset![TipoComprobante] = txtTipoComprobante.Text
    adoRegistroVentas.Recordset![Sucursal] = txtSucursal.Text
    adoRegistroVentas.Recordset![Letra] = txtLetra.Text
    adoRegistroVentas.Recordset![ImporteComprobante] = CSng(txtImporteTotal.Text)
    adoRegistroVentas.Recordset![BultosComprobante] = txtBultos.Text
    adoRegistroVentas.Recordset![UnidadesComprobante] = txtUnidades.Text
    adoRegistroVentas.Recordset![TipoRegistro] = "ITEM"
  
    adoMovimientoStock.Recordset.AddNew
    adoMovimientoStock.Recordset![IdMovStock] = vsMovStock + 1
    adoMovimientoStock.Recordset![Fecha] = Date
    adoMovimientoStock.Recordset![Hora] = Time
    adoMovimientoStock.Recordset![Motivo] = "VENTA"

    fgrDetalle.Row = vsI
    fgrDetalle.Col = 0
    adoRegistroVentas.Recordset![IdArticulo] = fgrDetalle.Text
    adoMovimientoStock.Recordset![IdArticulo] = fgrDetalle.Text
    
    adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [IdArticulo]= " & fgrDetalle.Text & " ORDER BY [IdArticulo]"
    adoArticulos.CommandType = adCmdText
    adoArticulos.Refresh
    
    adoBolsa.RecordSource = "SELECT * FROM [Bolsa] WHERE [IdArticulo]= " & fgrDetalle.Text & " ORDER BY [IdArticulo]"
    adoBolsa.CommandType = adCmdText
    adoBolsa.Refresh
    
    fgrDetalle.Col = 1
    adoRegistroVentas.Recordset![Descripcion] = fgrDetalle.Text
    adoMovimientoStock.Recordset![Descripcion] = fgrDetalle.Text
    
    fgrDetalle.Col = 2
    adoRegistroVentas.Recordset![BultosCantidad] = fgrDetalle.Text
    adoMovimientoStock.Recordset![StockBultos] = fgrDetalle.Text
    adoArticulos.Recordset![StockBultos] = adoArticulos.Recordset![StockBultos] - fgrDetalle.Text
    adoBolsa.Recordset![BultosCantidad] = adoBolsa.Recordset![BultosCantidad] + fgrDetalle.Text
    fgrDetalle.Col = 3
    adoRegistroVentas.Recordset![unidadesCantidad] = fgrDetalle.Text
    adoMovimientoStock.Recordset![StockUnidades] = fgrDetalle.Text
    If (adoArticulos.Recordset![StockUnidades] - fgrDetalle.Text) < 0 Then
      adoArticulos.Recordset![StockUnidades] = adoArticulos.Recordset![UxB] - fgrDetalle.Text
      adoArticulos.Recordset![StockBultos] = adoArticulos.Recordset![StockBultos] - 1
    Else
      adoArticulos.Recordset![StockUnidades] = adoArticulos.Recordset![StockUnidades] - fgrDetalle.Text
    End If
    adoBolsa.Recordset![unidadesCantidad] = adoBolsa.Recordset![unidadesCantidad] + fgrDetalle.Text
    
    fgrDetalle.Col = 4
    adoRegistroVentas.Recordset![PrecioVenta] = CSng(fgrDetalle.Text)
    
    fgrDetalle.Col = 5
    adoRegistroVentas.Recordset![SubTotal] = CSng(fgrDetalle.Text)
    
    adoRegistroVentas.Recordset.Update
    adoMovimientoStock.Recordset.Update
    adoArticulos.Recordset.Update
    adoBolsa.Recordset.Update
    adoRegistroVentas.Refresh
    adoMovimientoStock.Refresh
    adoArticulos.Refresh
    adoBolsa.Refresh
  
  Next vsI
        
  Call proActualizarVentas

  adoRegistroVentas.Refresh
  
  cmdNuevaLinea.Enabled = False
  
  cmdCancelar_Click
  
  If chkImprimir.Value = vbChecked Then
    
    'cAdoDB =
    'cAdoDB
    adoImpresion.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [IdVenta] =" & vsUltimaVenta & " AND [TipoRegistro]='ITEM' ORDER BY [IdVenta] DESC"
    adoImpresion.CommandType = adCmdText
    adoImpresion.Refresh
  
    If adoImpresion.Recordset.RecordCount = 0 Then
      MsgBox "No hay registros para listar", vbOKOnly, "Resultado"
      Exit Sub
    Else
      For vsI = 0 To 10 ' Puede ser cualquier numero que estimes conveniente
        crcPresupuesto.Formulas(vsI) = ""
      Next vsI
      
      cReporte = "{registroventas.idventa} = " & vsUltimaVenta & " and {registroventas.tiporegistro}='ITEM'"
      
      crcPresupuesto.WindowTitle = "Presupuesto"
      crcPresupuesto.SelectionFormula = (cReporte)
      'crcPresupuesto.WindowState = crptMaximized
      'crcPresupuesto.Action = 2
      
      crcPresupuesto.Destination = crptToPrinter
      crcPresupuesto.PrintReport
      
    End If
   
  End If
      
End Sub

Private Sub cmdAnular_Click()

  ' Almacena los datos en la tabla RegistroVentas para el tipo de registro "CABECERA"
  ' Aqui guarda todos los datos de la cabecera del comprobante, por eso en este registro
  ' los campos de Item de ventas estan vacios

  adoRegistroVentas.Recordset.AddNew
  adoRegistroVentas.Recordset![IdVenta] = txtIdVenta.Text
  adoRegistroVentas.Recordset![Fecha] = Date
  adoRegistroVentas.Recordset![Hora] = Time
  adoRegistroVentas.Recordset![Periodo] = DatePart("d", Date) & DatePart("m", Date) & DatePart("yyyy", Date)
  adoRegistroVentas.Recordset![IdCliente] = txtED_IdCliente.Text
  adoRegistroVentas.Recordset![RazonSocial] = txtRazonSocial.Text
  adoRegistroVentas.Recordset![TipoComprobante] = txtTipoComprobante.Text
  adoRegistroVentas.Recordset![Sucursal] = txtSucursal.Text
  adoRegistroVentas.Recordset![Letra] = txtLetra.Text
  adoRegistroVentas.Recordset![ImporteComprobante] = CSng(txtImporteTotal.Text)
  adoRegistroVentas.Recordset![BultosComprobante] = txtBultos.Text
  adoRegistroVentas.Recordset![UnidadesComprobante] = txtUnidades.Text
  adoRegistroVentas.Recordset![TipoRegistro] = "CABECERA"
  adoRegistroVentas.Recordset![IdArticulo] = 0
  adoRegistroVentas.Recordset![Descripcion] = ""
  adoRegistroVentas.Recordset![BultosCantidad] = 0
  adoRegistroVentas.Recordset![unidadesCantidad] = 0
  adoRegistroVentas.Recordset![PrecioVenta] = 0
  adoRegistroVentas.Recordset![SubTotal] = 0
  adoRegistroVentas.Recordset.Update

  vsUltimaDevolucion = txtIdVenta.Text
  
  ' Almacena los datos en la tabla VentaItem y luego en la tabla MovimientoStock en forma intercalada y
  ' actualiza el stock del artículo
  
  adoMovimientoStock.RecordSource = "SELECT * FROM [MovimientoStock] ORDER BY [IdMovStock]"
  adoMovimientoStock.CommandType = adCmdText
  adoMovimientoStock.Refresh
  vsMovStock = adoMovimientoStock.Recordset.RecordCount
  
  vsLineasTotal = fgrDetalle.Rows
  
  For vsI = 0 To vsLineasTotal - 2
    adoRegistroVentas.Recordset.AddNew
    adoRegistroVentas.Recordset![IdVenta] = txtIdVenta.Text
    adoRegistroVentas.Recordset![Fecha] = Date
    adoRegistroVentas.Recordset![Hora] = Time
    adoRegistroVentas.Recordset![IdCliente] = txtED_IdCliente.Text
    adoRegistroVentas.Recordset![RazonSocial] = txtRazonSocial.Text
    adoRegistroVentas.Recordset![TipoComprobante] = txtTipoComprobante.Text
    adoRegistroVentas.Recordset![Sucursal] = txtSucursal.Text
    adoRegistroVentas.Recordset![Letra] = txtLetra.Text
    adoRegistroVentas.Recordset![ImporteComprobante] = CSng(txtImporteTotal.Text)
    adoRegistroVentas.Recordset![BultosComprobante] = txtBultos.Text
    adoRegistroVentas.Recordset![UnidadesComprobante] = txtUnidades.Text
    adoRegistroVentas.Recordset![TipoRegistro] = "ITEM"
  
    adoMovimientoStock.Recordset.AddNew
    adoMovimientoStock.Recordset![IdMovStock] = vsMovStock + 1
    adoMovimientoStock.Recordset![Fecha] = Date
    adoMovimientoStock.Recordset![Hora] = Time
    adoMovimientoStock.Recordset![Motivo] = "DEVOLUCION"

    fgrDetalle.Row = vsI
    fgrDetalle.Col = 0
    adoRegistroVentas.Recordset![IdArticulo] = fgrDetalle.Text
    adoMovimientoStock.Recordset![IdArticulo] = fgrDetalle.Text
    
    adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [IdArticulo]= " & fgrDetalle.Text & " ORDER BY [IdArticulo]"
    adoArticulos.CommandType = adCmdText
    adoArticulos.Refresh
    
    adoBolsa.RecordSource = "SELECT * FROM [Bolsa] WHERE [IdArticulo]= " & fgrDetalle.Text & " ORDER BY [IdArticulo]"
    adoBolsa.CommandType = adCmdText
    adoBolsa.Refresh
    
    fgrDetalle.Col = 1
    adoRegistroVentas.Recordset![Descripcion] = fgrDetalle.Text
    adoMovimientoStock.Recordset![Descripcion] = fgrDetalle.Text
    
    fgrDetalle.Col = 2
    adoRegistroVentas.Recordset![BultosCantidad] = fgrDetalle.Text
    adoMovimientoStock.Recordset![StockBultos] = fgrDetalle.Text
    adoArticulos.Recordset![StockBultos] = adoArticulos.Recordset![StockBultos] - fgrDetalle.Text
    adoBolsa.Recordset![BultosCantidad] = adoBolsa.Recordset![BultosCantidad] + fgrDetalle.Text
    
    fgrDetalle.Col = 3
    adoRegistroVentas.Recordset![unidadesCantidad] = fgrDetalle.Text
    adoMovimientoStock.Recordset![StockUnidades] = fgrDetalle.Text
    If (adoArticulos.Recordset![StockUnidades] - fgrDetalle.Text) < 0 Then
      adoArticulos.Recordset![StockUnidades] = adoArticulos.Recordset![UxB] - fgrDetalle.Text
      adoArticulos.Recordset![StockBultos] = adoArticulos.Recordset![StockBultos] - 1
    Else
      adoArticulos.Recordset![StockUnidades] = adoArticulos.Recordset![StockUnidades] - fgrDetalle.Text
    End If
    adoBolsa.Recordset![unidadesCantidad] = adoBolsa.Recordset![unidadesCantidad] + fgrDetalle.Text
    
    fgrDetalle.Col = 4
    adoRegistroVentas.Recordset![PrecioVenta] = CSng(fgrDetalle.Text)
    
    fgrDetalle.Col = 5
    adoRegistroVentas.Recordset![SubTotal] = CSng(fgrDetalle.Text)
    
    adoRegistroVentas.Recordset.Update
    adoMovimientoStock.Recordset.Update
    adoArticulos.Recordset.Update
    adoBolsa.Recordset.Update
    adoRegistroVentas.Refresh
    adoMovimientoStock.Refresh
    adoArticulos.Refresh
    adoBolsa.Refresh
  
  Next vsI
  
  Call proActualizarVentas

  adoRegistroVentas.Refresh
  
  cmdNuevaLinea.Enabled = False
  
  cmdCancelar_Click
  
  If chkImprimir.Value = vbChecked Then
    
    'cAdoDB =
    'cAdoDB
    adoImpresion.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [IdVenta] =" & vsUltimaDevolucion & " AND [TipoRegistro]='ITEM' AND [TipoComprobante]='Devolución' ORDER BY [IdVenta] DESC"
    adoImpresion.CommandType = adCmdText
    adoImpresion.Refresh
  
    If adoImpresion.Recordset.RecordCount = 0 Then
      MsgBox "No hay registros para listar", vbOKOnly, "Resultado"
      Exit Sub
    Else
      For vsI = 0 To 10 ' Puede ser cualquier numero que estimes conveniente
        crcPresupuesto.Formulas(vsI) = ""
      Next vsI
      
      cReporte = "{registroventas.idventa} = " & vsUltimaVenta & " and {registroventas.tiporegistro}='ITEM' and {registroventas.tipocomprobante}='Devolución'"
      
      crcPresupuesto.WindowTitle = "Presupuesto"
      crcPresupuesto.SelectionFormula = (cReporte)
      'crcPresupuesto.WindowState = crptMaximized
      'crcPresupuesto.Action = 2
      
      crcPresupuesto.Destination = crptToPrinter
      crcPresupuesto.PrintReport
      
    End If
   
  End If

End Sub

Private Sub cmdModificar_Click()

  Call proActivarBotones(False, True, False, True, True)
  fgrDetalle.Clear
  
  ' Query para buscar cuantas Notas de credito hay entre los registros
  ' y poner el nro de comprobante (Nota de credito) siguiente
  adoRV_Devolucion.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [TipoComprobante]='Devolución' and [TipoRegistro]='CABECERA' ORDER BY [IdVenta]"
  adoRV_Devolucion.CommandType = adCmdText
  adoRV_Devolucion.Refresh
  vsIdNotaDeCredito = adoRV_Devolucion.Recordset.RecordCount + 1
 
  ' Query para buscar los registros de la venta seleccionada
  vsPresupuestoActivo = CDbl(dtgComprobantes.Columns(2).Text)
  adoRV_Devolucion.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [IdVenta]= " & vsPresupuestoActivo & " AND [TipoRegistro]='CABECERA' ORDER BY [IdVenta]"
  adoRV_Devolucion.CommandType = adCmdText
  adoRV_Devolucion.Refresh
  
  If adoRV_Devolucion.Recordset.RecordCount <> 0 Then
  
    ' Se mueve al primer registro donde estan los datos de la cabecera
    adoRV_Devolucion.Recordset.MoveFirst
    txtED_IdCliente.Text = adoRV_Devolucion.Recordset![IdCliente]
    txtRazonSocial.Text = adoRV_Devolucion.Recordset![RazonSocial]
    
    ' Query para obtener el canal de marketing del cliente
    adoClientes.RecordSource = "SELECT * FROM [Clientes] WHERE [IdCliente]= " & CInt(txtED_IdCliente.Text) & " ORDER BY [IdCliente]"
    adoClientes.CommandType = adCmdText
    adoClientes.Refresh
    txtCanal.Text = adoClientes.Recordset![Canal]
    
    ' Continua cargando los datos de la cabecera del comprobante
    txtTipoComprobante = "Devolución"
    txtLetra.Text = adoRV_Devolucion.Recordset![Letra]
    txtSucursal.Text = adoRV_Devolucion.Recordset![Sucursal]
    txtIdVenta.Text = vsIdNotaDeCredito
    txtBultos.Text = adoRV_Devolucion.Recordset![BultosComprobante] * (-1)
    txtUnidades.Text = adoRV_Devolucion.Recordset![UnidadesComprobante] * (-1)
    txtImporteTotal.Text = adoRV_Devolucion.Recordset![ImporteComprobante] * (-1)
    
    ' Query para obtener los registros con los items del comprobante seleccionado
    adoRV_Devolucion.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [IdVenta]= " & vsPresupuestoActivo & " AND [TipoRegistro]='ITEM' ORDER BY [IdVenta]"
    adoRV_Devolucion.CommandType = adCmdText
    adoRV_Devolucion.Refresh
    If adoRV_Devolucion.Recordset.RecordCount <> 0 Then
      adoRV_Devolucion.Recordset.MoveFirst
      
      ' Recorre el Recordset agregando en cada celda el dato correspondiente
      For vsI = 1 To adoRV_Devolucion.Recordset.RecordCount
        fgrDetalle.Row = vsI - 1
        fgrDetalle.Col = 0
        fgrDetalle.Text = adoRV_Devolucion.Recordset![IdArticulo]
        
        fgrDetalle.Col = 1
        fgrDetalle.CellAlignment = vbCenter
        fgrDetalle.Text = adoRV_Devolucion.Recordset![Descripcion]
        
        fgrDetalle.Col = 2
        fgrDetalle.Text = adoRV_Devolucion.Recordset![BultosCantidad] * (-1)
        
        fgrDetalle.Col = 3
        fgrDetalle.Text = adoRV_Devolucion.Recordset![unidadesCantidad] * (-1)

        fgrDetalle.Col = 4
        fgrDetalle.Text = adoRV_Devolucion.Recordset![PrecioVenta] * (-1)
        
        fgrDetalle.Col = 5
        fgrDetalle.Text = adoRV_Devolucion.Recordset![SubTotal] * (-1)
        
        fgrDetalle.Rows = vsI + 1
        adoRV_Devolucion.Recordset.MoveNext
      Next vsI
    
    End If
                 
  End If

End Sub

Private Sub cmdCancelar_Click()
  
  txtED_IdCliente.Text = ""
  txtRazonSocial.Text = ""
  txtCanal.Text = ""
  txtED_IdArticulo.Text = ""
  txtDescripcion.Text = ""
  txtED_CantidadBultos.Text = 1
  txtED_CantidadUnidades.Text = 0
  txtPrecioVenta.Text = 0
  txtSubTotal.Text = 0
  
  txtBultos.Text = 0
  txtUnidades.Text = 0
  txtImporteTotal.Text = 0
  
  fgrDetalle.Clear
  
  vsLineasTotal = 1
  fgrDetalle.Rows = vsLineasTotal
  vsLineaActiva = 0
  vsImporteTotal = 0
  vsBultos = 0
  
  lblStockActual.Caption = ""
  
  txtED_IdCliente.Locked = False
  txtED_IdCliente.SetFocus
  
  Call proListadoFull
  txtIdVenta.Text = vsIdVenta
  
  Call proActivarBotones(False, False, False, False, True)

End Sub

Private Sub txtIdVenta_Change()

  txtIdVenta.Text = Format(txtIdVenta.Text, "0##########")

End Sub

Private Sub txtSucursal_Change()

  txtSucursal.Text = Format(txtSucursal.Text, "0###")

End Sub

Private Sub cmdSalir_Click()
  
  Unload Me
  
End Sub

Private Sub dtgComprobantes_Click()

  vsUltimaVenta = dtgComprobantes.Columns(2).Text
  vsTipoComprobante = dtgComprobantes.Columns(1).Text
  
  adoRV_Item.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [TipoComprobante]='" & vsTipoComprobante & "' AND [TipoRegistro]='ITEM' AND [IdVenta]= " & vsUltimaVenta & " ORDER BY [IdVenta] DESC"
  adoRV_Item.CommandType = adCmdText
  adoRV_Item.Refresh

  Call proActivarBotones(False, False, True, True, True)

End Sub


Private Sub proLimpiaCampos()

  txtED_IdCliente.Text = ""
  txtRazonSocial.Text = ""
  txtCanal.Text = ""
  txtTipoComprobante.Text = ""
  txtLetra.Text = ""
  txtSucursal.Text = ""
  txtIdVenta.Text = ""
  txtED_IdArticulo.Text = ""
  txtDescripcion.Text = ""
  txtED_CantidadBultos.Text = 1
  txtED_CantidadUnidades.Text = 0
  txtPrecioVenta.Text = 0
  txtSubTotal.Text = 0
  cmdNuevaLinea.Enabled = False
  txtBultos.Text = ""
  txtImporteTotal.Text = ""
  
  Call proActivarBotones(False, False, False, False, True)

End Sub

Private Sub proActivarBotones(ByVal G As Boolean, ByVal A As Boolean, ByVal M As Boolean, ByVal C As Boolean, ByVal S As Boolean)

  cmdGuardar.Enabled = G
  cmdAnular.Enabled = A
  cmdModificar.Enabled = M
  cmdCancelar.Enabled = C
  cmdSalir.Enabled = S

End Sub

Private Sub proListadoFull()

  adoRegistroVentas.RecordSource = "SELECT * FROM [RegistroVentas] ORDER BY [IdVenta]"
  adoRegistroVentas.CommandType = adCmdText
  adoRegistroVentas.Refresh
  vsCantidadDeRegistros = adoRegistroVentas.Recordset.RecordCount
  If vsCantidadDeRegistros = 0 Then
    vsIdVenta = 1
  Else
    adoRegistroVentas.Recordset.MoveLast
    vsIdVenta = adoRegistroVentas.Recordset![IdVenta] + 1
  End If
End Sub

Private Sub proListarRegistros(ByVal Estado As String, ByVal Orden As String, ByVal ASCDES As String, ByVal Campo As String, ByVal Filtro As String)

  If Estado = "Todos" Then
    adoRegistroVentas.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [" & Campo & "] LIKE '%" & Filtro & "%' ORDER BY " & Orden & " " & ASCDES
  Else
    adoRegistroVentas.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [Activo]= '" & Estado & "' ORDER BY [" & Orden & "]" & ASCDES
  End If
  adoRegistroVentas.CommandType = adCmdText
  adoRegistroVentas.Refresh

End Sub

Private Sub proActualizarVentas()
  
  adoRV_Cabecera.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [TipoRegistro]='CABECERA' ORDER BY [Fecha] DESC, [IdVenta] DESC"
  adoRV_Cabecera.CommandType = adCmdText
  adoRV_Cabecera.Refresh
  vsUltimaVenta = adoRV_Cabecera.Recordset.RecordCount
  
  If vsUltimaVenta <> 0 Then
    vsUltimaVenta = dtgComprobantes.Columns(2).Text
    
    adoRV_Item.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [TipoRegistro]='ITEM' and [IdVenta]= " & vsUltimaVenta & " ORDER BY [IdVenta] DESC"
    adoRV_Item.CommandType = adCmdText
    adoRV_Item.Refresh
  End If
End Sub
