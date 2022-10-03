VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDistribuir 
   Caption         =   "Distribuir ventas"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6743.808
   ScaleMode       =   0  'User
   ScaleWidth      =   16326.91
   WindowState     =   2  'Maximized
   Begin VB.Frame frcBolsa 
      Height          =   8595
      Left            =   60
      TabIndex        =   1
      Top             =   630
      Width           =   7335
      Begin VB.Frame Frame1 
         Height          =   8295
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   7035
         Begin ComctlLib.ProgressBar prbAvance 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   6720
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Frame Frame3 
            Height          =   1125
            Left            =   90
            TabIndex        =   3
            Top             =   7020
            Width           =   6855
            Begin VB.CommandButton cmdSalir 
               Caption         =   "Salir"
               Height          =   795
               Left            =   3510
               Picture         =   "frmDistribuir.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   210
               Width           =   1155
            End
            Begin VB.CommandButton cmdProcesar 
               Caption         =   "Procesar"
               Height          =   795
               Left            =   2190
               Picture         =   "frmDistribuir.frx":0442
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   210
               Width           =   1155
            End
         End
         Begin MSDataGridLib.DataGrid dtgBolsa 
            Bindings        =   "frmDistribuir.frx":0884
            Height          =   6135
            Left            =   90
            TabIndex        =   6
            Top             =   180
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   10821
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
            Caption         =   "Stock para distribuir"
            ColumnCount     =   4
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
            BeginProperty Column02 
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
            BeginProperty Column03 
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1200,189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3404,977
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   794,835
               EndProperty
            EndProperty
         End
         Begin MSForms.Label lblProcesando 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   6390
            Width           =   6825
            Caption         =   "Procesando"
            Size            =   "12039;503"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
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
   Begin MSAdodcLib.Adodc adoBolsa 
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
   Begin MSAdodcLib.Adodc adoArticulos 
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
   Begin MSAdodcLib.Adodc adoTemporal 
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
      RecordSource    =   "Temporal"
      Caption         =   "adoTemporal"
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
      Caption         =   "DISTRIBUIR VENTAS"
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
Attribute VB_Name = "frmDistribuir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vsTotalAProcesar As Integer ' Cantidad de bultos y unidades para procesar
Dim vsProcesados As Integer ' Cantidad de bultos y unidades procesados
Dim vsI As Integer ' Contador
Dim vsJ As Integer ' Contador
Dim vsClienteAzar As Integer ' Código de cliente generado al azar
Dim vsMin As Integer ' Minimo código de cliente
Dim vsMax As Integer ' Máximo código de cliente
Dim vsPedidosxArt As Integer ' Cuenta la cantidad de pedidos de cada artículo que se puede realizar de acuerdo al stock de Bolsa
Dim vsNroPedido As Integer ' Número de pedido generado para distribuir
Dim vsLinea As String ' Almacena el texto que se guardara en el archivo de proceso
Dim vsNombreArchivo As String ' Almacena el nombre del archivo de proceso
Dim vsNoAction As String ' Indica si habia algún artículo para distribuir. Valor [SiHay] como Verdadero

Private Sub Form_Load()

  vsNoAction = ""
  lblTitulo.Top = 0
  lblTitulo.Left = 0
  lblTitulo.Width = Screen.Width
  frcBolsa.Left = (Screen.Width - frcBolsa.Width) / 2
  
  vsTotalAProcesar = 0
  vsNroPedido = 1
  
  adoBolsa.RecordSource = "SELECT * FROM [Bolsa] WHERE [BultosCantidad]>0 ORDER BY [BultosCantidad] DESC, [UnidadesCantidad] DESC"
  adoBolsa.CommandType = adCmdText
  adoBolsa.Refresh
  adoBolsa.Recordset.MoveFirst
  For vsI = 1 To adoBolsa.Recordset.RecordCount
    vsTotalAProcesar = vsTotalAProcesar + adoBolsa.Recordset![BultosCantidad] + adoBolsa.Recordset![unidadesCantidad]
    adoBolsa.Recordset.MoveNext
  Next vsI
  adoBolsa.Recordset.MoveFirst
  
  lblProcesando.Caption = "Cantidad a procesar: " & vsTotalAProcesar
  vsProcesados = 0
  If vsTotalAProcesar > 0 Then
    cmdProcesar.Enabled = True
  Else
    cmdProcesar.Enabled = False
  End If
  
End Sub

Private Sub cmdProcesar_Click()
    
  ' Lista todos los artículos de la tabla Bolsa (Solo los bultos,
  ' hay que hacer lo mismo para las unidades)
  adoBolsa.RecordSource = "SELECT * FROM [Bolsa] WHERE [BultosCantidad]>0 ORDER BY [BultosCantidad] DESC"
  adoBolsa.CommandType = adCmdText
  adoBolsa.Refresh
    
  If adoBolsa.Recordset.RecordCount > 0 Then
      
    ' Se posiciona en el primer registro
    adoBolsa.Recordset.MoveFirst
    
    ' Inicia el bucle de los artículos
    For vsI = 1 To adoBolsa.Recordset.RecordCount
      
      adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [IdArticulo]=" & adoBolsa.Recordset![IdArticulo] & " ORDER BY [IdArticulo]"
      adoArticulos.CommandType = adCmdText
      adoArticulos.Refresh
      
      ' Calcula cuantos pedidos puede generar con el stock disponible en la tabla Bolsa
      vsPedidosxArt = adoBolsa.Recordset![BultosCantidad] \ adoArticulos.Recordset![CantidadOptima]
      
      If vsPedidosxArt > 0 Then
       
        ' Indica que si hay artículos en condiciones para distribuir
        vsNoAction = "SiHay"
        
        For vsJ = 1 To vsPedidosxArt
        
          ' Genera el indice aleatorio para buscar el cliente en base
          adoClientes.RecordSource = "SELECT * FROM [Clientes] WHERE [HabilitaDistribucion]= 'Si' ORDER BY [IdCliente]"
          adoClientes.CommandType = adCmdText
          adoClientes.Refresh
          vsMin = 1
          vsMax = adoClientes.Recordset.RecordCount - 1
          
          Randomize Timer
          vsClienteAzar = vsMin + Int(Rnd() * (vsMax - vsMin + 1))
          
          ' Se posiciona en el registro generado aleatoriamente
          adoClientes.Recordset.Move vsClienteAzar
  
          ' Carga los datos en la tabla Temporal
          adoTemporal.RecordSource = "SELECT * FROM [Temporal] ORDER BY [nropedido]"
          adoTemporal.CommandType = adCmdText
          adoTemporal.Refresh
          
          adoTemporal.Recordset.AddNew
          adoTemporal.Recordset![nropedido] = vsNroPedido
          adoTemporal.Recordset![pdv] = adoClientes.Recordset![IdCliente]
          adoTemporal.Recordset![articulo] = adoArticulos.Recordset![IdArticulo]
          adoTemporal.Recordset![cantidad] = adoArticulos.Recordset![CantidadOptima]
          adoTemporal.Recordset![descuento] = adoArticulos.Recordset![DescuentoChess]
          adoTemporal.Recordset![vendedor] = adoClientes.Recordset![CodigoVendedor]
          adoTemporal.Recordset.Update
          adoTemporal.Refresh
          
          vsNroPedido = vsNroPedido + 1
          vsProcesados = vsProcesados + 1
          prbAvance.Value = (vsProcesados / vsTotalAProcesar) * 100
          
          ' Actualiza el stock de la tabla Bolsa
          adoBolsa.Recordset![BultosCantidad] = adoBolsa.Recordset![BultosCantidad] - adoArticulos.Recordset![CantidadOptima]
          adoBolsa.Recordset.Update
          'adoBolsa.Refresh
        
        Next vsJ
      End If
        
      adoBolsa.Recordset.MoveNext
    Next vsI

  End If
  ' Aqui inicia el mismo proceso
  ' pero para distrbuir las unidades
  adoBolsa.RecordSource = "SELECT * FROM [Bolsa] WHERE [UnidadesCantidad]>0 ORDER BY [UnidadesCantidad] DESC"
  adoBolsa.CommandType = adCmdText
  adoBolsa.Refresh
    
  If adoBolsa.Recordset.RecordCount > 0 Then
    ' Se posiciona en el primer registro
    adoBolsa.Recordset.MoveFirst
    
    ' Inicia el bucle de los artículos
    For vsI = 1 To adoBolsa.Recordset.RecordCount
      
      adoArticulos.RecordSource = "SELECT * FROM [Articulos] WHERE [IdArticulo]=" & adoBolsa.Recordset![IdArticulo] & " ORDER BY [IdArticulo]"
      adoArticulos.CommandType = adCmdText
      adoArticulos.Refresh
      
      ' Calcula cuantos pedidos puede generar con el stock disponible en la tabla Bolsa
      vsPedidosxArt = adoBolsa.Recordset![unidadesCantidad] \ adoArticulos.Recordset![CantidadOptima]
      
      If vsPedidosxArt > 0 Then
          
        ' Indica que si hay artículos en condiciones para distribuir
        vsNoAction = "SiHay"
        
        For vsJ = 1 To vsPedidosxArt
        
          ' Genera el indice aleatorio para buscar el cliente en base
          adoClientes.RecordSource = "SELECT * FROM [Clientes] WHERE [HabilitaDistribucion]= 'Si' ORDER BY [IdCliente]"
          adoClientes.CommandType = adCmdText
          adoClientes.Refresh
          vsMin = 1
          vsMax = adoClientes.Recordset.RecordCount - 1
                  
          Randomize Timer
          vsClienteAzar = vsMin + Int(Rnd() * (vsMax - vsMin + 1))
          
          ' Se posiciona en el registro generado aleatoriamente
          adoClientes.Recordset.Move vsClienteAzar
  
          ' Carga los datos en la tabla Temporal
          adoTemporal.RecordSource = "SELECT * FROM [Temporal] ORDER BY [nropedido]"
          adoTemporal.CommandType = adCmdText
          adoTemporal.Refresh
          
          adoTemporal.Recordset.AddNew
          adoTemporal.Recordset![nropedido] = vsNroPedido
          adoTemporal.Recordset![pdv] = adoClientes.Recordset![IdCliente]
          adoTemporal.Recordset![articulo] = adoArticulos.Recordset![IdArticulo]
          adoTemporal.Recordset![cantidad] = adoArticulos.Recordset![CantidadOptima]
          adoTemporal.Recordset![descuento] = adoArticulos.Recordset![DescuentoChess]
          adoTemporal.Recordset![vendedor] = adoClientes.Recordset![CodigoVendedor]
          adoTemporal.Recordset.Update
          adoTemporal.Refresh
          
          vsNroPedido = vsNroPedido + 1
          
          ' Actualiza el stock de la tabla Bolsa
          adoBolsa.Recordset![unidadesCantidad] = adoBolsa.Recordset![unidadesCantidad] - adoArticulos.Recordset![CantidadOptima]
          adoBolsa.Recordset.Update
          'adoBolsa.Refresh
        
        Next vsJ
      End If
        
      adoBolsa.Recordset.MoveNext
    Next vsI
  
    ' Lista todos los registros de la tabla Temporal
    adoTemporal.RecordSource = "SELECT * FROM [Temporal] ORDER BY [nropedido]"
    adoTemporal.CommandType = adCmdText
    adoTemporal.Refresh
  
  End If
  
  ' Verifica que el archivo de texto no exista anteriormente
  ' If Dir("C:\Archivo.TXT") = "" Then Kill ("C:\Archivo.TXT")
 
  vsNombreArchivo = Trim(Trim(Str(DatePart("d", Date))) & Trim(Str(DatePart("m", Date))) & Trim(Str(DatePart("yyyy", Date))))
  
  ' Crea el archivo de texto
  Open "C:\pedidosDEG_" & vsNombreArchivo & ".odb" For Output As #1
  
  ' Comienza a recorrer el Recordset y generar las lineas del archivo de texto
  If vsNoAction = "SiHay" Then
  
    If adoTemporal.Recordset.RecordCount > 0 Then
      adoTemporal.Recordset.MoveFirst
      For vsI = 1 To adoTemporal.Recordset.RecordCount
        vsLinea = Chr$(34) & adoTemporal.Recordset![nropedido] & Chr$(34) & "," & Chr$(34) & adoTemporal.Recordset![pdv] & Chr$(34) & "," & Chr$(34) & adoTemporal.Recordset![articulo] & Chr$(34) & "," & Chr$(34) & adoTemporal.Recordset![cantidad] & Chr$(34) & "," & Chr$(34) & adoTemporal.Recordset![descuento] & Chr$(34) & "," & Chr$(34) & adoTemporal.Recordset![vendedor] & Chr$(34) & "," & Chr$(34) & "0" & Chr$(34) & "," & Chr$(34) & "0" & Chr$(34)
        Print #1, vsLinea
        adoTemporal.Recordset.MoveNext
      Next vsI
      Close #1
    End If
  Else
    MsgBox "En este momento no hay artículos con las condiciones necesarias para poder distribuir. Intentelo nuevamente despúes de generar algunas ventas.", vbExclamation, "Aviso"
    cmdSalir.SetFocus
  End If
  prbAvance.Value = 0
  adoBolsa.RecordSource = "SELECT * FROM [Bolsa] WHERE [BultosCantidad]>0 ORDER BY [BultosCantidad] DESC, [UnidadesCantidad] DESC"
  adoBolsa.CommandType = adCmdText
  adoBolsa.Refresh
  If adoBolsa.Recordset.RecordCount > 0 Then
    adoBolsa.Recordset.MoveFirst
  End If
End Sub

Private Sub cmdSalir_Click()

  Unload Me

End Sub

