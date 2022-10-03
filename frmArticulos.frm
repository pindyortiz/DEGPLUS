VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmArticulos 
   Caption         =   "Artículos"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
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
   Begin VB.Frame frcArticulos 
      Height          =   8595
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   19545
      Begin VB.Frame Frame1 
         Height          =   8295
         Left            =   150
         TabIndex        =   21
         Top             =   150
         Width           =   7035
         Begin VB.TextBox txtED_PrecioBDE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
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
            Left            =   1920
            TabIndex        =   8
            Top             =   3030
            Width           =   1695
         End
         Begin VB.TextBox txtDB_PrecioBDE 
            Alignment       =   1  'Right Justify
            DataField       =   "PrecioBDE"
            DataSource      =   "adoArticulos"
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
            Left            =   1920
            TabIndex        =   47
            Top             =   3030
            Width           =   1695
         End
         Begin VB.TextBox txtED_DescuentoChess 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
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
            Left            =   3720
            TabIndex        =   6
            Top             =   2250
            Width           =   1695
         End
         Begin VB.TextBox txtDB_DescuentoChess 
            Alignment       =   1  'Right Justify
            DataField       =   "DescuentoChess"
            DataSource      =   "adoArticulos"
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
            Left            =   3720
            TabIndex        =   45
            Top             =   2250
            Width           =   1695
         End
         Begin VB.TextBox txtED_CantidadOptima 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
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
            Left            =   1920
            TabIndex        =   5
            Top             =   2250
            Width           =   1695
         End
         Begin VB.TextBox txtDB_CantidadOptima 
            Alignment       =   1  'Right Justify
            DataField       =   "CantidadOptima"
            DataSource      =   "adoArticulos"
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
            Left            =   1920
            TabIndex        =   43
            Top             =   2250
            Width           =   1695
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
            TabIndex        =   1
            Top             =   570
            Width           =   2115
         End
         Begin VB.TextBox txtED_IdSegunProveedor 
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
            Left            =   2310
            TabIndex        =   2
            Top             =   570
            Width           =   2115
         End
         Begin VB.TextBox txtED_Descripcion 
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
            Left            =   120
            TabIndex        =   3
            Top             =   1380
            Width           =   6765
         End
         Begin VB.TextBox txtED_PrecioPreventa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
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
            Left            =   120
            TabIndex        =   7
            Top             =   3030
            Width           =   1695
         End
         Begin VB.TextBox txtED_UxB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
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
            Left            =   120
            TabIndex        =   4
            Top             =   2250
            Width           =   1695
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
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3840
            Width           =   1695
         End
         Begin VB.TextBox txtDB_Activo 
            DataField       =   "Activo"
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
            TabIndex        =   42
            Top             =   3840
            Width           =   1695
         End
         Begin VB.TextBox txtDB_PrecioPreventa 
            Alignment       =   1  'Right Justify
            DataField       =   "PrecioPreventa"
            DataSource      =   "adoArticulos"
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
            Left            =   120
            TabIndex        =   39
            Top             =   3030
            Width           =   1695
         End
         Begin VB.TextBox txtDB_UxB 
            Alignment       =   1  'Right Justify
            DataField       =   "UxB"
            DataSource      =   "adoArticulos"
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
            Left            =   120
            TabIndex        =   37
            Top             =   2250
            Width           =   1695
         End
         Begin VB.TextBox txtDB_IdSegunProveedor 
            DataField       =   "IdSegunProveedor"
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
            Left            =   2310
            TabIndex        =   35
            Top             =   570
            Width           =   2115
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
            Height          =   450
            Left            =   90
            TabIndex        =   29
            Top             =   570
            Width           =   2115
         End
         Begin VB.TextBox txtDB_Descripcion 
            DataField       =   "Descripcion"
            DataSource      =   "adoArticulos"
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
            Left            =   120
            TabIndex        =   28
            Top             =   1380
            Width           =   6765
         End
         Begin VB.Frame Frame3 
            Height          =   1125
            Left            =   150
            TabIndex        =   22
            Top             =   7020
            Width           =   6735
            Begin VB.CommandButton cmdSalir 
               Caption         =   "Salir"
               Height          =   795
               Left            =   5430
               Picture         =   "frmArticulos.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdCancelar 
               Caption         =   "Cancelar"
               Height          =   795
               Left            =   4101
               Picture         =   "frmArticulos.frx":0442
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdGuardar 
               Caption         =   "Guardar"
               Height          =   795
               Left            =   2774
               Picture         =   "frmArticulos.frx":0884
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdModificar 
               Caption         =   "Modificar"
               Height          =   795
               Left            =   1447
               Picture         =   "frmArticulos.frx":0CC6
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdNuevo 
               Caption         =   "Nuevo"
               Height          =   795
               Left            =   120
               Picture         =   "frmArticulos.frx":1108
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   210
               Width           =   1155
            End
         End
         Begin VB.Label lbl_PrecioBDE 
            Caption         =   "Precio BDE"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1950
            TabIndex        =   48
            Top             =   2790
            Width           =   1635
         End
         Begin VB.Label lbl_DescuentoChess 
            Caption         =   "Descuento Chess"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3750
            TabIndex        =   46
            Top             =   2010
            Width           =   1665
         End
         Begin VB.Label lbl_CantidadOptima 
            Caption         =   "Cantidad óptima"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1950
            TabIndex        =   44
            Top             =   2010
            Width           =   1365
         End
         Begin VB.Label lbl_PrecioPreventa 
            Caption         =   "Precio preventa"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   40
            Top             =   2790
            Width           =   1635
         End
         Begin VB.Label lbl_UxB 
            Caption         =   "Unidades por bulto"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   38
            Top             =   2010
            Width           =   1695
         End
         Begin VB.Label lbl_IdSegunProveedor 
            Caption         =   "Código segun proveedor"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2340
            TabIndex        =   36
            Top             =   330
            Width           =   2115
         End
         Begin VB.Label lbl_IdArticulo 
            Caption         =   "Código"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   330
            Width           =   2055
         End
         Begin VB.Label lbl_Descripcion 
            Caption         =   "Descripción"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1170
            Width           =   2055
         End
         Begin VB.Label lbl_Activo 
            Caption         =   "Activo"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   90
            TabIndex        =   31
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Label lblRequeridos 
            Caption         =   "Los campos destacados en rojo son obligatorios."
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
            Left            =   2610
            TabIndex        =   30
            Top             =   6690
            Visible         =   0   'False
            Width           =   4275
         End
      End
      Begin VB.Frame Frame2 
         Height          =   8295
         Left            =   7290
         TabIndex        =   10
         Top             =   150
         Width           =   12105
         Begin VB.Frame Frame4 
            Height          =   795
            Left            =   90
            TabIndex        =   11
            Top             =   150
            Width           =   11895
            Begin VB.CommandButton cmdASCDES 
               Caption         =   "ASC"
               Height          =   285
               Left            =   6240
               TabIndex        =   41
               Top             =   270
               Width           =   735
            End
            Begin VB.TextBox txtFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   345
               Left            =   9690
               TabIndex        =   17
               Top             =   270
               Width           =   2085
            End
            Begin VB.ComboBox cmbCamposFiltro 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   7890
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   270
               Width           =   1725
            End
            Begin VB.ComboBox cmbCamposOrden 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Left            =   4440
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   270
               Width           =   1725
            End
            Begin VB.OptionButton optInactivos 
               Caption         =   "Inactivos"
               Height          =   285
               Left            =   2310
               TabIndex        =   14
               Top             =   270
               Width           =   1125
            End
            Begin VB.OptionButton optActivos 
               Caption         =   "Activos"
               Height          =   285
               Left            =   1200
               TabIndex        =   13
               Top             =   270
               Width           =   1125
            End
            Begin VB.OptionButton optTodos 
               Caption         =   "Todos"
               Height          =   285
               Left            =   90
               TabIndex        =   12
               Top             =   270
               Width           =   1125
            End
            Begin VB.Label Label7 
               Caption         =   "Filtro"
               Height          =   255
               Left            =   7440
               TabIndex        =   19
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Label6 
               Caption         =   "Orden"
               Height          =   255
               Left            =   3930
               TabIndex        =   18
               Top             =   300
               Width           =   765
            End
         End
         Begin MSDataGridLib.DataGrid dtgArticulos 
            Bindings        =   "frmArticulos.frx":154A
            Height          =   7185
            Left            =   120
            TabIndex        =   20
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
            ColumnCount     =   11
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
               Caption         =   "Id x proveedor"
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
               DataField       =   "CantidadOptima"
               Caption         =   "Cantidad óptima"
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
               DataField       =   "DescuentoChess"
               Caption         =   "Descuento Chess"
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
               DataField       =   "PrecioPreventa"
               Caption         =   "Precio preventa"
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
               DataField       =   "PrecioBDE"
               Caption         =   "Precio BDE"
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
               DataField       =   "StockBultos"
               Caption         =   "StockBultos"
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
               DataField       =   "StockUnidades"
               Caption         =   "StockUnidades"
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
                  ColumnWidth     =   1154,835
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1154,835
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2594,835
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   404,787
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnWidth     =   1275,024
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnWidth     =   1395,213
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   1275,024
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   1275,024
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   870,236
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1110,047
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   750,047
               EndProperty
            EndProperty
         End
      End
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
   Begin MSAdodcLib.Adodc adoBolsa 
      Height          =   330
      Left            =   11310
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
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "ARTÍCULOS"
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
      TabIndex        =   34
      Top             =   0
      Width           =   8625
   End
End
Attribute VB_Name = "frmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vsPasoNM As String ' Variable de paso por Nuevo o Modificar
Dim vsRegistrosArticulos As Integer ' Cuantos registros tiene la tabla de Articulos
Dim vsIdSegunProveedor As Double '
Dim vsIdArticuloActual As Double '
Dim vsI As Integer ' Contador simple
Dim vsTAI As String ' Lista Todos, Activos o Inactivos
Dim vsOrden As String ' Establace el orden de los registros
Dim vsASCDES As String ' Establece si el orden es ASCendente o DEScendente
Dim vsCampo As String ' Establece que campo se usara para el filtro
Dim vsFiltro As String ' Filtra los canales
Dim vsPrecioPreventa As Single '
Dim vsPrecioBDE As Single '
Dim vsIdArticulo_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsIdSegunProveedor_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsDescripcion_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsUxB_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsPrecioPreventa_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsPrecioBDE_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsCantidadOptima_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio
Dim vsDescuentoChess_Completo As Boolean ' Es TRUE cuando el campo esta NO vacio

Private Sub Form_Load()
  
  lblTitulo.Top = 0
  lblTitulo.Left = 0
  lblTitulo.Width = Screen.Width
  frcArticulos.Left = (Screen.Width - frcArticulos.Width) / 2
  
  proLimpiaCampos
    
  cmbED_Activo.AddItem "Activo"
  cmbED_Activo.AddItem "Inactivo"
  cmbED_Activo.Text = cmbED_Activo.List(0)
  
  cmbCamposOrden.AddItem "Código"
  cmbCamposOrden.AddItem "Código según poveedor"
  cmbCamposOrden.AddItem "Descripción"
  cmbCamposOrden.AddItem "Unidades por bulto"
  cmbCamposOrden.AddItem "Precio preventa"
  cmbCamposOrden.AddItem "Precio BDE"
  cmbCamposOrden.AddItem "Cantidad óptima"
  cmbCamposOrden.AddItem "Descuento Chess"
  cmbCamposOrden.AddItem "Activo"
  cmbCamposOrden.Text = cmbCamposOrden.List(0)
  
  cmbCamposFiltro.AddItem "Código"
  cmbCamposFiltro.AddItem "Código según proveedor"
  cmbCamposFiltro.AddItem "Descripción"
  cmbCamposFiltro.AddItem "Unidades por bulto"
  cmbCamposFiltro.AddItem "Precio preventa"
  cmbCamposFiltro.AddItem "Precio BDE"
  cmbCamposFiltro.AddItem "Cantidad óptima"
  cmbCamposFiltro.AddItem "Descuento Chess"
  cmbCamposFiltro.AddItem "Activo"
  cmbCamposFiltro.Text = cmbCamposFiltro.List(0)
  
  optTodos.Value = True
  
  Call proListadoFull
  
  If vsRegistrosArticulos = 0 Then
    Call proActivarBotones(True, False, False, False, True)
  Else
    Call proActivarBotones(True, True, False, False, True)
  End If
  
  vsTAI = "Todos"
  vsOrden = "IdArticulo"
  vsASCDES = "ASC"
  vsCampo = "IdArticulo"
  vsFiltro = ""
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub txtED_IdArticulo_Change()

  If txtED_IdArticulo.Visible = True Then
    If txtED_IdArticulo.Text <> "" Then
      lbl_IdArticulo.ForeColor = vbBlack
      lbl_IdArticulo.FontBold = False
      vsIdArticulo_Completo = True
    Else
      lbl_IdArticulo.ForeColor = vbRed
      lbl_IdArticulo.FontBold = True
      vsIdArticulo_Completo = False
    End If
  
    If (vsIdArticulo_Completo And vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) = True Then
      lblRequeridos.Visible = False
    Else
      lblRequeridos.Visible = True
    End If
  End If

End Sub

Private Sub txtED_IdSegunProveedor_Change()
  
  If txtED_IdSegunProveedor.Visible = True Then
    If txtED_IdSegunProveedor.Text <> "" Then
      lbl_IdSegunProveedor.ForeColor = vbBlack
      lbl_IdSegunProveedor.FontBold = False
      vsIdSegunProveedor_Completo = True
    Else
      lbl_IdSegunProveedor.ForeColor = vbRed
      lbl_IdSegunProveedor.FontBold = True
      vsIdSegunProveedor_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
    If (vsIdArticulo_Completo And vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If (vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If

End Sub

Private Sub txtED_Descripcion_Change()
  
  If txtED_Descripcion.Visible = True Then
    If txtED_Descripcion.Text <> "" Then
      lbl_Descripcion.ForeColor = vbBlack
      lbl_Descripcion.FontBold = False
      vsDescripcion_Completo = True
    Else
      lbl_Descripcion.ForeColor = vbRed
      lbl_Descripcion.FontBold = True
      vsDescripcion_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
    If (vsIdArticulo_Completo And vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If (vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If

End Sub

Private Sub txtED_UxB_Change()
  
  If txtED_UxB.Visible = True Then
    If txtED_UxB.Text <> "" Then
      lbl_UxB.ForeColor = vbBlack
      lbl_UxB.FontBold = False
      vsUxB_Completo = True
    Else
      lbl_UxB.ForeColor = vbRed
      lbl_UxB.FontBold = True
      vsUxB_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
    If (vsIdArticulo_Completo And vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If (vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If

End Sub

Private Sub txtED_PrecioPreventa_Change()
  
  If txtED_PrecioPreventa.Visible = True Then
    If txtED_PrecioPreventa.Text <> "" Then
      lbl_PrecioPreventa.ForeColor = vbBlack
      lbl_PrecioPreventa.FontBold = False
      vsPrecioPreventa_Completo = True
    Else
      lbl_PrecioPreventa.ForeColor = vbRed
      lbl_PrecioPreventa.FontBold = True
      vsPrecioPreventa_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
    If (vsIdArticulo_Completo And vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If (vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If

End Sub

Private Sub txtED_PrecioBDE_Change()
  
  If txtED_PrecioBDE.Visible = True Then
    If txtED_PrecioBDE.Text <> "" Then
      lbl_PrecioBDE.ForeColor = vbBlack
      lbl_PrecioBDE.FontBold = False
      vsPrecioBDE_Completo = True
    Else
      lbl_PrecioBDE.ForeColor = vbRed
      lbl_PrecioBDE.FontBold = True
      vsPrecioBDE_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
    If (vsIdArticulo_Completo And vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If (vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If

End Sub

Private Sub txtED_CantidadOptima_Change()
  
  If txtED_CantidadOptima.Visible = True Then
    If txtED_CantidadOptima.Text <> "" Then
      lbl_CantidadOptima.ForeColor = vbBlack
      lbl_CantidadOptima.FontBold = False
      vsCantidadOptima_Completo = True
    Else
      lbl_CantidadOptima.ForeColor = vbRed
      lbl_CantidadOptima.FontBold = True
      vsCantidadOptima_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
    If (vsIdArticulo_Completo And vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If (vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If

End Sub

Private Sub txtED_DescuentoChess_Change()
  
  If txtED_DescuentoChess.Visible = True Then
    If txtED_DescuentoChess.Text <> "" Then
      lbl_DescuentoChess.ForeColor = vbBlack
      lbl_DescuentoChess.FontBold = False
      vsDescuentoChess_Completo = True
    Else
      lbl_DescuentoChess.ForeColor = vbRed
      lbl_DescuentoChess.FontBold = True
      vsDescuentoChess_Completo = False
    End If
    
    If vsPasoNM = "NUEVO" Then
    If (vsIdArticulo_Completo And vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) = True Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    Else
      If (vsIdSegunProveedor_Completo And vsDescripcion_Completo And vsUxB_Completo And vsPrecioPreventa_Completo And vsPrecioBDE_Completo And vsCantidadOptima_Completo And vsDescuentoChess_Completo) Then
        lblRequeridos.Visible = False
      Else
        lblRequeridos.Visible = True
      End If
    End If
  End If

End Sub

Private Sub cmdNuevo_Click()

  vsPasoNM = "NUEVO"
  
  txtED_IdArticulo.Text = ""
  txtED_IdSegunProveedor.Text = ""
  txtED_Descripcion.Text = ""
  txtED_UxB.Text = ""
  txtED_PrecioPreventa.Text = ""
  txtED_PrecioBDE.Text = ""
  txtED_CantidadOptima.Text = ""
  txtED_DescuentoChess.Text = ""

  txtED_IdArticulo.Visible = True
  txtED_IdSegunProveedor.Visible = True
  txtED_Descripcion.Visible = True
  txtED_UxB.Visible = True
  txtED_PrecioPreventa.Visible = True
  txtED_PrecioBDE.Visible = True
  txtED_CantidadOptima.Visible = True
  txtED_DescuentoChess.Visible = True
  cmbED_Activo.Visible = True
   
  txtED_IdArticulo.SetFocus
  
  lblRequeridos.Visible = True
  lbl_IdArticulo.ForeColor = vbRed
  lbl_IdArticulo.FontBold = True
  lbl_IdSegunProveedor.ForeColor = vbRed
  lbl_IdSegunProveedor.FontBold = True
  lbl_Descripcion.ForeColor = vbRed
  lbl_Descripcion.FontBold = True
  lbl_UxB.ForeColor = vbRed
  lbl_UxB.FontBold = True
  lbl_PrecioPreventa.ForeColor = vbRed
  lbl_PrecioPreventa.FontBold = True
  lbl_PrecioBDE.ForeColor = vbRed
  lbl_PrecioBDE.FontBold = True
  lbl_CantidadOptima.ForeColor = vbRed
  lbl_CantidadOptima.FontBold = True
  lbl_DescuentoChess.ForeColor = vbRed
  lbl_DescuentoChess.FontBold = True
  
  vsIdArticulo_Completo = False
  vsIdSegunProveedor = False
  vsDescripcion_Completo = False
  vsPrecioPreventa_Completo = False
  vsPrecioBDE_Completo = False
  
  Call proADControlGrid
   
  Call proActivarBotones(False, False, True, True, False)
  
End Sub

Private Sub cmdModificar_Click()

  vsPasoNM = "MODIFICAR"
  
  txtED_IdSegunProveedor.Visible = True
  
  Call proADControlGrid
  
  txtED_Descripcion.Visible = True
  txtED_UxB.Visible = True
  txtED_PrecioPreventa.Visible = True
  txtED_PrecioBDE.Visible = True
  txtED_CantidadOptima.Visible = True
  txtED_DescuentoChess.Visible = True
  cmbED_Activo.Visible = True
  
  txtED_IdSegunProveedor.Text = txtDB_IdSegunProveedor.Text
  txtED_Descripcion.Text = txtDB_Descripcion.Text
  txtED_UxB.Text = txtDB_UxB.Text
  txtED_PrecioPreventa.Text = txtDB_PrecioPreventa.Text
  txtED_PrecioBDE.Text = txtDB_PrecioBDE.Text
  txtED_CantidadOptima.Text = txtDB_CantidadOptima.Text
  txtED_DescuentoChess.Text = txtDB_DescuentoChess.Text
  cmbED_Activo.Text = txtDB_Activo.Text
  
  txtED_IdSegunProveedor.SelStart = 0
  txtED_IdSegunProveedor.SelLength = Len(txtED_IdSegunProveedor.Text)
  
  txtED_IdSegunProveedor.SetFocus
  
  Call proActivarBotones(False, False, True, True, False)
   
End Sub

Private Sub cmdGuardar_Click()

  If vsPasoNM = "NUEVO" Then
    vsIdArticuloActual = CDbl(txtED_IdArticulo.Text)
  Else
    vsIdArticuloActual = CDbl(txtDB_IdArticulo.Text)
  End If
  
  If txtED_IdSegunProveedor.Text = "" Then
    MsgBox "El dato de 'Código según proveedor' no puede estar vacío. Debe completar ese campo. Si su proveedor no asignó un código para este artículo deberá hacerlo usted.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
  If txtED_Descripcion.Text = "" Then
    MsgBox "El dato de 'Descripción' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
  If txtED_UxB.Text = "" Then
    MsgBox "El dato de 'Unidades por bulto' no puedde estar vacío. Debe completar ese campo.", vbExclamation, "Dato erroneo"
    Exit Sub
  End If
  If Not IsNumeric(txtED_UxB.Text) Then
    MsgBox "El dato de 'Unidades por bulto' debe ser un número. Debe coregir ese dato.", vbExclamation, "Dato erroneo"
    Exit Sub
  End If
  If txtED_PrecioPreventa.Text = "" Then
    MsgBox "El dato de 'Precio preventa' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
  If Not IsNumeric(txtED_PrecioPreventa.Text) Then
    MsgBox "El dato de 'Precio preventa' debe ser un número. Debe coregir ese dato.", vbExclamation, "Dato erroneo"
    Exit Sub
  End If
  If txtED_PrecioBDE.Text = "" Then
    MsgBox "El dato de 'Precio BDE' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
  If Not IsNumeric(txtED_PrecioBDE.Text) Then
    MsgBox "El dato de 'Precio BDE' debe ser un número. Debe coregir ese dato.", vbExclamation, "Dato erroneo"
    Exit Sub
  End If
  If txtED_CantidadOptima.Text = "" Then
    MsgBox "El dato de 'Cantidad óptima' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
  If Not IsNumeric(txtED_CantidadOptima.Text) Then
    MsgBox "El dato de 'Cantidad óptima' debe ser un número. Debe coregir ese dato.", vbExclamation, "Dato erroneo"
    Exit Sub
  End If
  If txtED_DescuentoChess.Text = "" Then
    MsgBox "El dato de 'Descuento Chess' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
    Exit Sub
  End If
  If Not IsNumeric(txtED_DescuentoChess.Text) Then
    MsgBox "El dato de 'Descuento Chess' debe ser un número. Debe coregir ese dato.", vbExclamation, "Dato erroneo"
    Exit Sub
  End If
  
  If vsPasoNM = "NUEVO" Then
    If txtED_IdArticulo.Text = "" Then
      MsgBox "El dato de 'Código' no puede estar vacío. Debe completar ese campo.", vbExclamation, "Dato esperado"
      Exit Sub
    End If
    If Not IsNumeric(txtED_IdArticulo) Then
      MsgBox "El dato de 'Código' debe ser un número. Debe corregir este dato.", vbExclamation, "Dato erroneo"
      Exit Sub
    End If
  
    adoArticulos.Recordset.AddNew
    adoArticulos.Recordset![IdArticulo] = txtED_IdArticulo.Text
    adoArticulos.Recordset![StockBultos] = 0
    adoArticulos.Recordset![StockUnidades] = 0
    
  End If
  
' INICIO: Registro del artículo en la tabla Artículos
  
  adoArticulos.Recordset![IdSegunProveedor] = txtED_IdSegunProveedor.Text
  adoArticulos.Recordset![Descripcion] = txtED_Descripcion.Text
  adoArticulos.Recordset![UxB] = txtED_UxB.Text
  adoArticulos.Recordset![Preciopreventa] = Replace(txtED_PrecioPreventa.Text, ",", ".")
  adoArticulos.Recordset![PrecioBDE] = Replace(txtED_PrecioBDE.Text, ",", ".")
  adoArticulos.Recordset![CantidadOptima] = txtED_CantidadOptima.Text
  adoArticulos.Recordset![DescuentoChess] = Replace(txtED_DescuentoChess.Text, ",", ".")
  adoArticulos.Recordset![Activo] = cmbED_Activo.Text
  adoArticulos.Recordset.Update
  adoArticulos.Refresh
  
' FIN: Registro del artículo en la tabla Artículos
  
' INICIO: Registro del artículo en la tabla Bolsa
  
  If vsPasoNM = "NUEVO" Then
  
    adoBolsa.RecordSource = "SELECT * FROM [Bolsa] ORDER BY [IdArticulo]"
    adoBolsa.CommandType = adCmdText
    adoBolsa.Refresh

    adoBolsa.Recordset.AddNew
    adoBolsa.Recordset![IdArticulo] = txtED_IdArticulo.Text
    adoBolsa.Recordset![BultosCantidad] = 0
    adoBolsa.Recordset![unidadesCantidad] = 0
  
  Else
  
    adoBolsa.RecordSource = "SELECT * FROM [Bolsa] WHERE [IdArticulo]=" & vsIdArticuloActual & " ORDER BY [IdArticulo]"
    adoBolsa.CommandType = adCmdText
    adoBolsa.Refresh
    
  End If
    
  adoBolsa.Recordset![Descripcion] = txtED_Descripcion.Text

  adoBolsa.Recordset.Update
  adoBolsa.Refresh
    
' FIN: Registro del artículo en la tabla Bolsa
  
' INICIO: Registro del artículo en la tabla MovimientoStock
  
   If vsPasoNM = "NUEVO" Then
   
    adoMovimientoStock.RecordSource = "SELECT * FROM [MovimientoStock] ORDER BY [IdMovStock]"
    adoMovimientoStock.CommandType = adCmdText
    adoMovimientoStock.Refresh
 
    adoMovimientoStock.Recordset.AddNew
    adoMovimientoStock.Recordset![IdMovStock] = adoMovimientoStock.Recordset.RecordCount
    adoMovimientoStock.Recordset![Fecha] = Date
    adoMovimientoStock.Recordset![Hora] = Time
    adoMovimientoStock.Recordset![Motivo] = "NUEVO ARTICULO"
    adoMovimientoStock.Recordset![IdArticulo] = txtED_IdArticulo
    adoMovimientoStock.Recordset![Descripcion] = txtED_Descripcion
    adoMovimientoStock.Recordset![StockBultos] = 0
    adoMovimientoStock.Recordset![StockUnidades] = 0
    adoMovimientoStock.Recordset.Update
    
  End If
  
' FIN: Registro del artículo en la tabla MovimientoStock
     
  Call proListadoFull
  
  Call cmdCancelar_Click

End Sub

Private Sub cmdCancelar_Click()

  proLimpiaCampos
  
  lbl_IdArticulo.ForeColor = vbBlack
  lbl_IdArticulo.FontBold = False
  lbl_IdSegunProveedor.ForeColor = vbBlack
  lbl_IdSegunProveedor.FontBold = False
  lbl_Descripcion.ForeColor = vbBlack
  lbl_Descripcion.FontBold = False
  lbl_UxB.ForeColor = vbBlack
  lbl_UxB.FontBold = False
  lbl_PrecioPreventa.ForeColor = vbBlack
  lbl_PrecioPreventa.FontBold = False
  lbl_PrecioBDE.ForeColor = vbBlack
  lbl_PrecioBDE.FontBold = False
  lbl_CantidadOptima.ForeColor = vbBlack
  lbl_CantidadOptima.FontBold = False
  lbl_DescuentoChess.ForeColor = vbBlack
  lbl_DescuentoChess.FontBold = False
  lbl_Activo.ForeColor = vbBlack
  lbl_Activo.FontBold = False
  lblRequeridos.Visible = False
  
  If vsRegistrosArticulos = 0 Then
    Call proActivarBotones(True, False, False, False, True)
  Else
    Call proActivarBotones(True, True, False, False, True)
  End If

  txtDB_IdArticulo.SetFocus
  
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
    Case "Código": vsOrden = "IdArticulo"
    Case "Código según proveedor": vsOrden = "IdSegunProveedor"
    Case "Descripción": vsOrden = "Descripcion"
    Case "Unidades por bulto": vsOrden = "UxB"
    Case "Precio preventa": vsOrden = "PrecioPreventa"
    Case "Precio BDE": vsOrden = "PrecioBDE"
    Case "Cantidad óptima": vsOrden = "CantidadOptima"
    Case "Descuento Chess": vsOrden = "DescuentoChess"
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
    Case "Unidades por bulto": vsCampo = "UxB"
    Case "Precio preventa": vsCampo = "PrecioPreventa"
    Case "Precio BDE": vsCampo = "PrecioBDE"
    Case "Cantidad óptima": vsCampo = "CantidadOptima"
    Case "Descuento Chess": vsCampo = "DescuentoChess"
    Case "Activo": vsCampo = "Activo"
  End Select

  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub txtFiltro_Change()
  
  vsFiltro = txtFiltro.Text
  
  Call proListarRegistros(vsTAI, vsOrden, vsASCDES, vsCampo, vsFiltro)
  
End Sub

Private Sub proLimpiaCampos()

  txtED_IdArticulo.Text = ""
  txtED_IdSegunProveedor.Text = ""
  txtED_Descripcion.Text = ""
  txtED_UxB.Text = ""
  txtED_PrecioPreventa.Text = ""
  txtED_PrecioBDE.Text = ""
  txtED_CantidadOptima.Text = ""
  txtED_DescuentoChess.Text = ""
  
  txtED_IdArticulo.Visible = False
  txtED_IdSegunProveedor.Visible = False
  txtED_Descripcion.Visible = False
  txtED_UxB.Visible = False
  txtED_PrecioPreventa.Visible = False
  txtED_PrecioBDE.Visible = False
  txtED_CantidadOptima.Visible = False
  txtED_DescuentoChess.Visible = False
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

  adoArticulos.RecordSource = "SELECT * FROM [Articulos] ORDER BY [IdArticulo]"
  adoArticulos.CommandType = adCmdText
  adoArticulos.Refresh
  vsRegistrosArticulos = adoArticulos.Recordset.RecordCount
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



