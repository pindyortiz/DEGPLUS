VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaArticulos 
   Caption         =   "Listado de artículos"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleMode       =   0  'User
   ScaleWidth      =   21903.53
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   795
      Left            =   9810
      Picture         =   "frmListaArticulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5490
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   795
      Left            =   11055
      Picture         =   "frmListaArticulos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5490
      Width           =   1155
   End
   Begin VB.Frame frcListaArticulos 
      Height          =   4755
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   12165
      Begin MSDataGridLib.DataGrid dtgListaArticulos 
         Bindings        =   "frmListaArticulos.frx":0884
         Height          =   4455
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   7858
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
            Caption         =   "CantidadOptima"
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
            Caption         =   "DescuentoChess"
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
            Caption         =   "PrecioPreventa"
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
            Caption         =   "Stock bultos"
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
            Caption         =   "Stock unidades"
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
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4004,788
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   404,787
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1140,095
            EndProperty
         EndProperty
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
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "LISTADO DE ARTÍCULOS"
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
      TabIndex        =   4
      Top             =   0
      Width           =   8625
   End
End
Attribute VB_Name = "frmListaArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vsTotalFilas As Integer ' Total de filas con clientes filtrados
Dim vsFilaActual As Integer ' Fila actual seleccionada

Private Sub Form_Load()

  lblTitulo.Top = 0
  lblTitulo.Left = 0
  lblTitulo.Width = Screen.Width
  frcListaArticulos.Left = (Screen.Width - frcListaArticulos.Width) / 2
  
  vsReturnIdArticulo = 0
  vsReturnDescripcion = ""
  vsReturnPrecioVenta = 0
  
  adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [Descripcion] LIKE '%" & vsQueryDescripcion & "%' ORDER BY [IdArticulo]"
  adoArticulos.CommandType = adCmdText
  adoArticulos.Refresh
  
  
  
  'vsTotalFilas = adoClientes.Recordset.RecordCount
      
  'If vsTotalFilas <> 0 Then
    
  '  dtgListaClientes.RowBookmark (0)
  '  vsFilaActual = 0
  
  'End If
  
  'dtgListaClientes.SelBookmarks.Add (dtgListaClientes.Bookmark)
  
End Sub

Private Sub Form_Activate()

  dtgListaArticulos.SetFocus

End Sub

'Private Sub dtgListaClientes_KeyPress(KeyAscii As Integer)
'
'  If KeyAscii = 13 Then
'
'    vsQueryRazonSocial = ""
'    vsReturnIdCliente = dtgListaClientes.Columns(0).Text
'   vsReturnRazonSocial = dtgListaClientes.Columns(1).Text
'    vsReturnCanal = dtgListaClientes.Columns(2).Text
'
'    Unload frmListaClientes
'
'  End If
'
'End Sub

Private Sub cmdSeleccionar_Click()
    
  vsQueryDescripcion = ""
  vsReturnIdArticulo = dtgListaArticulos.Columns(0).Text
  vsReturnDescripcion = dtgListaArticulos.Columns(2).Text
  vsReturnPrecioVenta = dtgListaArticulos.Columns(4).Text
  
  vsVieneDe = "ARTICULOS"
  
  Unload frmListaArticulos

End Sub

Private Sub cmdSalir_Click()

  Unload Me

End Sub

