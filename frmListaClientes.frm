VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaClientes 
   Caption         =   "Listado de clientes"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleMode       =   0  'User
   ScaleWidth      =   11265
   Begin VB.Frame frcListaClientes 
      Height          =   4755
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   11115
      Begin MSDataGridLib.DataGrid dtgListaClientes 
         Bindings        =   "frmListaClientes.frx":0000
         Height          =   4455
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   10935
         _ExtentX        =   19288
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "IdCliente"
            Caption         =   "Código de cliente"
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
               ColumnWidth     =   1500,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4500,284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3000,189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1244,976
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   795
      Left            =   9945
      Picture         =   "frmListaClientes.frx":001A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5490
      Width           =   1155
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   795
      Left            =   8610
      Picture         =   "frmListaClientes.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5490
      Width           =   1155
   End
   Begin MSAdodcLib.Adodc adoClientes 
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
      Caption         =   "LISTADO DE CLIENTES"
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
Attribute VB_Name = "frmListaClientes"
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
  frcListaClientes.Left = (Screen.Width - frcListaClientes.Width) / 2
  
  vsReturnIdCliente = 0
  vsReturnRazonSocial = ""
  vsReturnCanal = ""
  
  adoClientes.RecordSource = "SELECT * FROM [Clientes] WHERE [RazonSocial] LIKE '%" & vsQueryRazonSocial & "%' ORDER BY [IdCliente]"
  adoClientes.CommandType = adCmdText
  adoClientes.Refresh
  'vsTotalFilas = adoClientes.Recordset.RecordCount
      
  'If vsTotalFilas <> 0 Then
    
  '  dtgListaClientes.RowBookmark (0)
  '  vsFilaActual = 0
  
  'End If
  
  'dtgListaClientes.SelBookmarks.Add (dtgListaClientes.Bookmark)
  
End Sub

Private Sub Form_Activate()

  dtgListaClientes.SetFocus

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
    
  vsQueryRazonSocial = ""
  vsReturnIdCliente = dtgListaClientes.Columns(0).Text
  vsReturnRazonSocial = dtgListaClientes.Columns(1).Text
  vsReturnCanal = dtgListaClientes.Columns(2).Text
  
  vsVieneDe = "CLIENTES"
  
  Unload frmListaClientes

End Sub

Private Sub cmdSalir_Click()

  Unload Me

End Sub

