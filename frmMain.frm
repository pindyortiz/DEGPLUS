VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "DEG PLUS - Presupuestos"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   12870
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList 
      Left            =   360
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1085A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":210B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3190E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42168
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   1588
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artículos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stock"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ventas"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Opciones"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu m_Articulos 
      Caption         =   "Artículos"
      Begin VB.Menu sm_Articulos 
         Caption         =   "Artículos"
      End
      Begin VB.Menu sm_ControlDeStock 
         Caption         =   "Control de stock"
      End
   End
   Begin VB.Menu m_Clientes 
      Caption         =   "Clientes"
      Begin VB.Menu sm_Clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu sm_Canales 
         Caption         =   "Canales de marketing"
      End
   End
   Begin VB.Menu m_ventas 
      Caption         =   "Ventas"
      Begin VB.Menu sm_Ventas 
         Caption         =   "Ventas"
      End
      Begin VB.Menu sm_Distribuir 
         Caption         =   "Distribuir"
      End
      Begin VB.Menu sm_CajaDiaria 
         Caption         =   "Caja diaria"
      End
   End
   Begin VB.Menu m_Configuracion 
      Caption         =   "Configuración"
      Begin VB.Menu sm_Opciones 
         Caption         =   "Opciones"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sm_Articulos_Click()

  frmArticulos.Show

End Sub

Private Sub sm_CajaDiaria_Click()

  frmCajaDiaria.Show

End Sub

Private Sub sm_Canales_Click()

  frmCanales.Show

End Sub

Private Sub sm_Clientes_Click()
  
  frmClientes.Show
  
End Sub

Private Sub sm_ControlDeStock_Click()

  frmStock.Show

End Sub

Private Sub sm_Distribuir_Click()

  frmDistribuir.Show

End Sub

Private Sub sm_Ventas_Click()

  frmRegistroVentas.Show

End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
  
 Select Case Button.Index

    Case 1: frmArticulos.Show
    Case 4: frmClientes.Show
    Case 6: frmRegistroVentas.Show
  End Select

End Sub
