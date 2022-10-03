VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCajaDiaria 
   Caption         =   "Caja diaria"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   8130
   WindowState     =   2  'Maximized
   Begin VB.Frame frcCajaDiaria 
      Height          =   8625
      Left            =   60
      TabIndex        =   0
      Top             =   630
      Width           =   11835
      Begin VB.Frame Frame3 
         Height          =   2715
         Left            =   60
         TabIndex        =   3
         Top             =   5850
         Width           =   11685
         Begin VB.TextBox txtFecha 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2970
            TabIndex        =   9
            Top             =   450
            Width           =   2715
         End
         Begin VB.TextBox txtED_Recaudacion 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2970
            TabIndex        =   8
            Top             =   1410
            Width           =   2715
         End
         Begin VB.CommandButton cmdFiltrar 
            Caption         =   "Filtrar"
            Height          =   795
            Left            =   9090
            Picture         =   "frmCajaDiaria.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1830
            Width           =   1155
         End
         Begin VB.CommandButton cmdSalir 
            Caption         =   "Salir"
            Height          =   795
            Left            =   10410
            Picture         =   "frmCajaDiaria.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1800
            Width           =   1155
         End
         Begin MSComCtl2.MonthView mvcFecha 
            Height          =   2370
            Left            =   90
            TabIndex        =   6
            Top             =   210
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            StartOfWeek     =   125829121
            CurrentDate     =   44830
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha seleccionada"
            Height          =   285
            Left            =   2970
            TabIndex        =   10
            Top             =   210
            Width           =   2715
         End
         Begin VB.Label Label1 
            Caption         =   "Recaudación"
            Height          =   195
            Left            =   2970
            TabIndex        =   7
            Top             =   1170
            Width           =   2715
         End
      End
      Begin MSDataGridLib.DataGrid dtgVentaDiaria 
         Bindings        =   "frmCajaDiaria.frx":0884
         Height          =   5625
         Left            =   60
         TabIndex        =   2
         Top             =   180
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   9922
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
         Caption         =   "Detalle de ventas"
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
            DataField       =   "Hora"
            Caption         =   "Hora"
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
            DataField       =   "TipoComprobante"
            Caption         =   "Comprobante"
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
            DataField       =   "Letra"
            Caption         =   "Letra"
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
         BeginProperty Column05 
            DataField       =   "IdVenta"
            Caption         =   "Nro comprobante"
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
         BeginProperty Column07 
            DataField       =   "RazonSocial"
            Caption         =   "Razón Social"
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
         BeginProperty Column09 
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
         BeginProperty Column10 
            DataField       =   "ImporteComprobante"
            Caption         =   "Importe"
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
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   494,929
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1995,024
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   1305,071
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
   Begin MSAdodcLib.Adodc adoRegistroVentas 
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
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "CAJA DIARIA"
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
      TabIndex        =   1
      Top             =   0
      Width           =   8625
   End
End
Attribute VB_Name = "frmCajaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vsI As Integer ' Contador
Dim vsRecaudacion As Single ' Almacena la recaudación de la fecha seleccionada
Dim vsPeriodo As String ' Periodo(Fecha) para filtrar ventas

Private Sub Form_Load()
  
  lblTitulo.Top = 0
  lblTitulo.Left = 0
  lblTitulo.Width = Screen.Width
  frcCajaDiaria.Left = (Screen.Width - frcCajaDiaria.Width) / 2

  adoRegistroVentas.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [TipoRegistro]='CABECERA' ORDER BY [IdVenta] DESC"
  adoRegistroVentas.CommandType = adCmdText
  adoRegistroVentas.Refresh
  
  vsPeriodo = DatePart("d", Date) & DatePart("m", Date) & DatePart("yyyy", Date)
  txtFecha.Text = Date

End Sub

Private Sub mvcFecha_DateClick(ByVal DateClicked As Date)

  vsPeriodo = DatePart("d", mvcFecha.Value) & DatePart("m", mvcFecha.Value) & DatePart("yyyy", mvcFecha.Value)
  txtFecha = mvcFecha.Value
  cmdFiltrar.SetFocus

End Sub

Private Sub cmdFiltrar_Click()
  
  vsRecaudacion = 0
  
  adoRegistroVentas.RecordSource = "SELECT * FROM [RegistroVentas] WHERE [TipoRegistro]='CABECERA' and [Periodo]='" & vsPeriodo & "' ORDER BY [IdVenta] DESC"
  adoRegistroVentas.CommandType = adCmdText
  adoRegistroVentas.Refresh
  If adoRegistroVentas.Recordset.RecordCount <> 0 Then
    
    adoRegistroVentas.Recordset.MoveFirst
    
    For vsI = 1 To adoRegistroVentas.Recordset.RecordCount
      vsRecaudacion = vsRecaudacion + adoRegistroVentas.Recordset![ImporteComprobante]
      adoRegistroVentas.Recordset.MoveNext
    Next vsI
  
    txtED_Recaudacion.Text = Format(vsRecaudacion, "$ #####,##")
   Else
    
    MsgBox "No se registraron ventas en el período seleccionado. Por favor seleccione otra fecha.", vbExclamation, "Aviso"
    Exit Sub
   End If
End Sub

Private Sub cmdSalir_Click()

  Unload Me

End Sub
