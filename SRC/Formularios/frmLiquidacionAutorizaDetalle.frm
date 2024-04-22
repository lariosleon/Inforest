VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLiquidacionAutorizacionDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuadres de Liquidaciones"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab 
      Height          =   6210
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   10954
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos Principales"
      TabPicture(0)   =   "frmLiquidacionAutorizaDetalle.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label(9)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label(10)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblCN"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCE"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblLN"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblLE"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDN"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDE"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label(13)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label(7)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtFondoCN"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtFondoCE"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtEfectivoCN"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtEfectivoCE"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCreditoCN"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCreditoCE"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtChequeCN"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtChequeCE"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtOtroCN"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtOtroCE"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtPuntoCN"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtPuntoCE"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtFondoLN"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtFondoLE"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtEfectivoDE"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtEfectivoDN"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtEfectivoLE"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtCreditoDE"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtCreditoDN"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtEfectivoLN"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtCreditoLE"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtCreditoLN"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtChequeDE"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtChequeDN"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtChequeLE"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtChequeLN"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtOtroLN"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtOtroLE"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtOtroDN"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtOtroDE"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtPuntoLN"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtPuntoLE"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtPuntoDN"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtPuntoDE"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtDiferencia"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtTotalCE"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtTotalCN"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "cmdOpcion(2)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cmdOpcion(3)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).ControlCount=   57
      TabCaption(1)   =   "Tarjetas de Crédito"
      TabPicture(1)   =   "frmLiquidacionAutorizaDetalle.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTarjeta(1)"
      Tab(1).Control(1)=   "lblTarjeta(2)"
      Tab(1).Control(2)=   "lblTarjeta(3)"
      Tab(1).Control(3)=   "lblTarjeta(4)"
      Tab(1).Control(4)=   "lblTarjeta(5)"
      Tab(1).Control(5)=   "lblTarjeta(6)"
      Tab(1).Control(6)=   "lblTarjeta(7)"
      Tab(1).Control(7)=   "lblTarjeta(8)"
      Tab(1).Control(8)=   "Label1"
      Tab(1).Control(9)=   "Label2"
      Tab(1).Control(10)=   "Label(21)"
      Tab(1).Control(11)=   "Label(33)"
      Tab(1).Control(12)=   "Label5"
      Tab(1).Control(13)=   "Label6"
      Tab(1).Control(14)=   "Label(22)"
      Tab(1).Control(15)=   "Label7"
      Tab(1).Control(16)=   "Label8"
      Tab(1).Control(17)=   "Label(36)"
      Tab(1).Control(18)=   "txtTarjetaCE(1)"
      Tab(1).Control(19)=   "txtTarjetaCN(1)"
      Tab(1).Control(20)=   "txtTarjetaCN(2)"
      Tab(1).Control(21)=   "txtTarjetaCN(3)"
      Tab(1).Control(22)=   "txtTarjetaCN(4)"
      Tab(1).Control(23)=   "txtTarjetaCN(5)"
      Tab(1).Control(24)=   "txtTarjetaCN(6)"
      Tab(1).Control(25)=   "txtTarjetaCN(7)"
      Tab(1).Control(26)=   "txtTarjetaCE(2)"
      Tab(1).Control(27)=   "txtTarjetaCE(3)"
      Tab(1).Control(28)=   "txtTarjetaCE(4)"
      Tab(1).Control(29)=   "txtTarjetaCE(5)"
      Tab(1).Control(30)=   "txtTarjetaCE(6)"
      Tab(1).Control(31)=   "txtTarjetaCE(7)"
      Tab(1).Control(32)=   "txtTarjetaCE(8)"
      Tab(1).Control(33)=   "txtTarjetaCN(8)"
      Tab(1).Control(34)=   "txtTotalTarjetaCE"
      Tab(1).Control(35)=   "txtTotalTarjetaCN"
      Tab(1).Control(36)=   "txtTarjetaLE(1)"
      Tab(1).Control(37)=   "txtTarjetaLN(1)"
      Tab(1).Control(38)=   "txtTarjetaLN(2)"
      Tab(1).Control(39)=   "txtTarjetaLN(3)"
      Tab(1).Control(40)=   "txtTarjetaLN(4)"
      Tab(1).Control(41)=   "txtTarjetaLN(5)"
      Tab(1).Control(42)=   "txtTarjetaLN(6)"
      Tab(1).Control(43)=   "txtTarjetaLN(7)"
      Tab(1).Control(44)=   "txtTarjetaLE(2)"
      Tab(1).Control(45)=   "txtTarjetaLE(3)"
      Tab(1).Control(46)=   "txtTarjetaLE(4)"
      Tab(1).Control(47)=   "txtTarjetaLE(5)"
      Tab(1).Control(48)=   "txtTarjetaLE(6)"
      Tab(1).Control(49)=   "txtTarjetaLE(7)"
      Tab(1).Control(50)=   "txtTarjetaLE(8)"
      Tab(1).Control(51)=   "txtTarjetaLN(8)"
      Tab(1).Control(52)=   "txtTarjetaDE(1)"
      Tab(1).Control(53)=   "txtTarjetaDN(1)"
      Tab(1).Control(54)=   "txtTarjetaDN(2)"
      Tab(1).Control(55)=   "txtTarjetaDN(3)"
      Tab(1).Control(56)=   "txtTarjetaDN(4)"
      Tab(1).Control(57)=   "txtTarjetaDN(5)"
      Tab(1).Control(58)=   "txtTarjetaDN(6)"
      Tab(1).Control(59)=   "txtTarjetaDN(7)"
      Tab(1).Control(60)=   "txtTarjetaDE(2)"
      Tab(1).Control(61)=   "txtTarjetaDE(3)"
      Tab(1).Control(62)=   "txtTarjetaDE(4)"
      Tab(1).Control(63)=   "txtTarjetaDE(5)"
      Tab(1).Control(64)=   "txtTarjetaDE(6)"
      Tab(1).Control(65)=   "txtTarjetaDE(7)"
      Tab(1).Control(66)=   "txtTarjetaDE(8)"
      Tab(1).Control(67)=   "txtTarjetaDN(8)"
      Tab(1).Control(68)=   "txtTotalTarjetaLE"
      Tab(1).Control(69)=   "txtTotalTarjetaLN"
      Tab(1).Control(70)=   "txtTotalTarjetaDE"
      Tab(1).Control(71)=   "txtTotalTarjetaDN"
      Tab(1).ControlCount=   72
      TabCaption(2)   =   "Otras Cancelaciones"
      TabPicture(2)   =   "frmLiquidacionAutorizaDetalle.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label(34)"
      Tab(2).Control(1)=   "txtOTotalCE"
      Tab(2).Control(2)=   "txtOTotalCN"
      Tab(2).Control(3)=   "txtOTotalLE"
      Tab(2).Control(4)=   "txtOTotalLN"
      Tab(2).Control(5)=   "txtOTotalDE"
      Tab(2).Control(6)=   "txtOTotalDN"
      Tab(2).Control(7)=   "SSTab1"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Aprobaciones"
      TabPicture(3)   =   "frmLiquidacionAutorizaDetalle.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtObservacion"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdOpcion(1)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdOpcion(0)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "chkControler"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "chkAdministrador"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtControler"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtAdministrador"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label3"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.TextBox txtObservacion 
         Height          =   1380
         Left            =   -74820
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   315
         Top             =   825
         Width           =   9510
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   156
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   9340
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Otros"
         TabPicture(0)   =   "frmLiquidacionAutorizaDetalle.frx":0070
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label(38)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblODN(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblODE(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label(37)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblOLN(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblOLE(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblOtro(10)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label(32)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblOCN(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lblOCE(0)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lblOtro(9)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lblOtro(8)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lblOtro(7)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "lblOtro(6)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lblOtro(5)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "lblOtro(4)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "lblOtro(3)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "lblOtro(2)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "lblOtro(1)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txtODE(10)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txtODE(9)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txtODE(6)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtODE(4)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtODE(2)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "txtODN(2)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txtODN(10)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txtODN(9)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txtODN(8)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "txtODE(8)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txtODE(7)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txtODE(5)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "txtODE(3)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txtODN(7)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "txtODN(6)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "txtODN(5)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "txtODN(4)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "txtODN(3)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "txtODN(1)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "txtODE(1)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "txtOLE(10)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "txtOLE(9)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "txtOLE(6)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "txtOLE(4)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "txtOLE(2)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "txtOLN(2)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "txtOLN(10)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "txtOLN(9)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "txtOLN(8)"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "txtOLE(8)"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "txtOLE(7)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "txtOLE(5)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "txtOLE(3)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "txtOLN(7)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "txtOLN(6)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "txtOLN(5)"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "txtOLN(4)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "txtOLN(3)"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "txtOLN(1)"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "txtOLE(1)"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "txtOCE(10)"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "txtOCE(9)"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "txtOCE(6)"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "txtOCE(4)"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "txtOCE(2)"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "txtOCN(2)"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "txtOCN(10)"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "txtOCN(9)"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).Control(67)=   "txtOCN(8)"
         Tab(0).Control(67).Enabled=   0   'False
         Tab(0).Control(68)=   "txtOCE(8)"
         Tab(0).Control(68).Enabled=   0   'False
         Tab(0).Control(69)=   "txtOCE(7)"
         Tab(0).Control(69).Enabled=   0   'False
         Tab(0).Control(70)=   "txtOCE(5)"
         Tab(0).Control(70).Enabled=   0   'False
         Tab(0).Control(71)=   "txtOCE(3)"
         Tab(0).Control(71).Enabled=   0   'False
         Tab(0).Control(72)=   "txtOCN(7)"
         Tab(0).Control(72).Enabled=   0   'False
         Tab(0).Control(73)=   "txtOCN(6)"
         Tab(0).Control(73).Enabled=   0   'False
         Tab(0).Control(74)=   "txtOCN(5)"
         Tab(0).Control(74).Enabled=   0   'False
         Tab(0).Control(75)=   "txtOCN(4)"
         Tab(0).Control(75).Enabled=   0   'False
         Tab(0).Control(76)=   "txtOCN(3)"
         Tab(0).Control(76).Enabled=   0   'False
         Tab(0).Control(77)=   "txtOCN(1)"
         Tab(0).Control(77).Enabled=   0   'False
         Tab(0).Control(78)=   "txtOCE(1)"
         Tab(0).Control(78).Enabled=   0   'False
         Tab(0).ControlCount=   79
         TabCaption(1)   =   "Otros"
         TabPicture(1)   =   "frmLiquidacionAutorizaDetalle.frx":008C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label(16)"
         Tab(1).Control(1)=   "lblODN(1)"
         Tab(1).Control(2)=   "lblODE(1)"
         Tab(1).Control(3)=   "Label(19)"
         Tab(1).Control(4)=   "lblOLN(1)"
         Tab(1).Control(5)=   "lblOLE(1)"
         Tab(1).Control(6)=   "lblOtro(20)"
         Tab(1).Control(7)=   "Label(20)"
         Tab(1).Control(8)=   "lblOCN(1)"
         Tab(1).Control(9)=   "lblOCE(1)"
         Tab(1).Control(10)=   "lblOtro(19)"
         Tab(1).Control(11)=   "lblOtro(18)"
         Tab(1).Control(12)=   "lblOtro(17)"
         Tab(1).Control(13)=   "lblOtro(16)"
         Tab(1).Control(14)=   "lblOtro(15)"
         Tab(1).Control(15)=   "lblOtro(14)"
         Tab(1).Control(16)=   "lblOtro(13)"
         Tab(1).Control(17)=   "lblOtro(12)"
         Tab(1).Control(18)=   "lblOtro(11)"
         Tab(1).Control(19)=   "txtODE(20)"
         Tab(1).Control(20)=   "txtODE(19)"
         Tab(1).Control(21)=   "txtODE(16)"
         Tab(1).Control(22)=   "txtODE(14)"
         Tab(1).Control(23)=   "txtODE(12)"
         Tab(1).Control(24)=   "txtODN(12)"
         Tab(1).Control(25)=   "txtODN(20)"
         Tab(1).Control(26)=   "txtODN(19)"
         Tab(1).Control(27)=   "txtODN(18)"
         Tab(1).Control(28)=   "txtODE(18)"
         Tab(1).Control(29)=   "txtODE(17)"
         Tab(1).Control(30)=   "txtODE(15)"
         Tab(1).Control(31)=   "txtODE(13)"
         Tab(1).Control(32)=   "txtODN(17)"
         Tab(1).Control(33)=   "txtODN(16)"
         Tab(1).Control(34)=   "txtODN(15)"
         Tab(1).Control(35)=   "txtODN(14)"
         Tab(1).Control(36)=   "txtODN(13)"
         Tab(1).Control(37)=   "txtODN(11)"
         Tab(1).Control(38)=   "txtODE(11)"
         Tab(1).Control(39)=   "txtOLE(20)"
         Tab(1).Control(40)=   "txtOLE(19)"
         Tab(1).Control(41)=   "txtOLE(16)"
         Tab(1).Control(42)=   "txtOLE(14)"
         Tab(1).Control(43)=   "txtOLE(12)"
         Tab(1).Control(44)=   "txtOLN(12)"
         Tab(1).Control(45)=   "txtOLN(20)"
         Tab(1).Control(46)=   "txtOLN(19)"
         Tab(1).Control(47)=   "txtOLN(18)"
         Tab(1).Control(48)=   "txtOLE(18)"
         Tab(1).Control(49)=   "txtOLE(17)"
         Tab(1).Control(50)=   "txtOLE(15)"
         Tab(1).Control(51)=   "txtOLE(13)"
         Tab(1).Control(52)=   "txtOLN(17)"
         Tab(1).Control(53)=   "txtOLN(16)"
         Tab(1).Control(54)=   "txtOLN(15)"
         Tab(1).Control(55)=   "txtOLN(14)"
         Tab(1).Control(56)=   "txtOLN(13)"
         Tab(1).Control(57)=   "txtOLN(11)"
         Tab(1).Control(58)=   "txtOLE(11)"
         Tab(1).Control(59)=   "txtOCE(20)"
         Tab(1).Control(60)=   "txtOCE(19)"
         Tab(1).Control(61)=   "txtOCE(16)"
         Tab(1).Control(62)=   "txtOCE(14)"
         Tab(1).Control(63)=   "txtOCE(12)"
         Tab(1).Control(64)=   "txtOCN(12)"
         Tab(1).Control(65)=   "txtOCN(20)"
         Tab(1).Control(66)=   "txtOCN(19)"
         Tab(1).Control(67)=   "txtOCN(18)"
         Tab(1).Control(68)=   "txtOCE(18)"
         Tab(1).Control(69)=   "txtOCE(17)"
         Tab(1).Control(70)=   "txtOCE(15)"
         Tab(1).Control(71)=   "txtOCE(13)"
         Tab(1).Control(72)=   "txtOCN(17)"
         Tab(1).Control(73)=   "txtOCN(16)"
         Tab(1).Control(74)=   "txtOCN(15)"
         Tab(1).Control(75)=   "txtOCN(14)"
         Tab(1).Control(76)=   "txtOCN(13)"
         Tab(1).Control(77)=   "txtOCN(11)"
         Tab(1).Control(78)=   "txtOCE(11)"
         Tab(1).ControlCount=   79
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   11
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   295
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   11
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   294
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   13
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   293
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   14
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   292
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   15
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   291
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   16
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   290
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   17
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   289
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   13
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   288
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   15
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   287
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   17
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   286
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   18
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   285
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   18
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   284
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   19
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   283
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   20
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   282
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   12
            Left            =   -72885
            Locked          =   -1  'True
            TabIndex        =   281
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   12
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   280
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   14
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   279
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   16
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   278
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   19
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   277
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   20
            Left            =   -71445
            Locked          =   -1  'True
            TabIndex        =   276
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   11
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   275
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   11
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   274
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   13
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   273
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   14
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   272
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   15
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   271
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   16
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   270
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   17
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   269
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   13
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   268
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   15
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   267
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   17
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   266
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   18
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   265
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   18
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   264
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   19
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   263
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   20
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   262
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   12
            Left            =   -69780
            Locked          =   -1  'True
            TabIndex        =   261
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   12
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   260
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   14
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   259
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   16
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   258
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   19
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   257
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   20
            Left            =   -68340
            Locked          =   -1  'True
            TabIndex        =   256
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   11
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   255
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   11
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   254
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   13
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   253
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   14
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   252
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   15
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   251
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   16
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   250
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   17
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   249
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   13
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   248
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   15
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   247
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   17
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   246
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   18
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   245
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   18
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   244
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   19
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   243
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   20
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   242
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   12
            Left            =   -66675
            Locked          =   -1  'True
            TabIndex        =   241
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   12
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   240
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   14
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   239
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   16
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   238
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   19
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   237
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   20
            Left            =   -65235
            Locked          =   -1  'True
            TabIndex        =   236
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   1
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   216
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   1
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   215
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   3
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   214
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   4
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   213
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   5
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   212
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   6
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   211
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   7
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   210
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   3
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   209
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   5
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   208
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   7
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   207
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   8
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   206
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   8
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   205
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   9
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   204
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   10
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   203
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtOCN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   2
            Left            =   2115
            Locked          =   -1  'True
            TabIndex        =   202
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   2
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   201
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   4
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   200
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   6
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   199
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   9
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   198
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtOCE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   10
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   197
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   1
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   196
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   1
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   195
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   3
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   194
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   4
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   193
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   5
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   192
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   6
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   191
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   7
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   190
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   3
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   189
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   5
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   188
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   7
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   187
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   8
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   186
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   8
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   185
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   9
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   184
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   10
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   183
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtOLN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   2
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   182
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   2
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   181
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   4
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   180
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   6
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   179
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   9
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   178
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtOLE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Index           =   10
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   177
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   1
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   176
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   1
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   175
            Text            =   "0.00"
            Top             =   1170
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   3
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   174
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   4
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   173
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   5
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   172
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   6
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   171
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   7
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   170
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   3
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   169
            Text            =   "0.00"
            Top             =   1980
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   5
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   168
            Text            =   "0.00"
            Top             =   2790
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   7
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   167
            Text            =   "0.00"
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   8
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   166
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   8
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   165
            Text            =   "0.00"
            Top             =   4005
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   9
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   164
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   10
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   163
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.TextBox txtODN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   2
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   162
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   2
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   161
            Text            =   "0.00"
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   4
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   160
            Text            =   "0.00"
            Top             =   2385
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   6
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   159
            Text            =   "0.00"
            Top             =   3195
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   9
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   158
            Text            =   "0.00"
            Top             =   4410
            Width           =   1365
         End
         Begin VB.TextBox txtODE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   10
            Left            =   9765
            Locked          =   -1  'True
            TabIndex        =   157
            Text            =   "0.00"
            Top             =   4815
            Width           =   1365
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   11
            Left            =   -74880
            TabIndex        =   314
            Top             =   1215
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   12
            Left            =   -74880
            TabIndex        =   313
            Top             =   1620
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   13
            Left            =   -74880
            TabIndex        =   312
            Top             =   2025
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   14
            Left            =   -74880
            TabIndex        =   311
            Top             =   2430
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   15
            Left            =   -74880
            TabIndex        =   310
            Top             =   2835
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   16
            Left            =   -74880
            TabIndex        =   309
            Top             =   3240
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   17
            Left            =   -74880
            TabIndex        =   308
            Top             =   3645
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   18
            Left            =   -74880
            TabIndex        =   307
            Top             =   4050
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   19
            Left            =   -74880
            TabIndex        =   306
            Top             =   4455
            Width           =   1920
         End
         Begin VB.Label lblOCE 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   -71175
            TabIndex        =   305
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblOCN 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   -72615
            TabIndex        =   304
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Otras Cancelaciones Comprobantes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   20
            Left            =   -73560
            TabIndex        =   303
            Top             =   360
            Width           =   3480
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   20
            Left            =   -74880
            TabIndex        =   302
            Top             =   4860
            Width           =   1920
         End
         Begin VB.Label lblOLE 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   -68070
            TabIndex        =   301
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblOLN 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   -69510
            TabIndex        =   300
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Otras Cancelaciones Liquidación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   19
            Left            =   -69825
            TabIndex        =   299
            Top             =   360
            Width           =   2850
         End
         Begin VB.Label lblODE 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   -64965
            TabIndex        =   298
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblODN 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   -66405
            TabIndex        =   297
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Otras Cancelaciones Diferencias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   16
            Left            =   -66720
            TabIndex        =   296
            Top             =   360
            Width           =   2850
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   235
            Top             =   1215
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   234
            Top             =   1620
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   233
            Top             =   2025
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   232
            Top             =   2430
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   231
            Top             =   2835
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   230
            Top             =   3240
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   229
            Top             =   3645
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   228
            Top             =   4050
            Width           =   1920
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   9
            Left            =   120
            TabIndex        =   227
            Top             =   4455
            Width           =   1920
         End
         Begin VB.Label lblOCE 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   3825
            TabIndex        =   226
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblOCN 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2385
            TabIndex        =   225
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Otras Cancelaciones Comprobantes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   32
            Left            =   1440
            TabIndex        =   224
            Top             =   360
            Width           =   3480
         End
         Begin VB.Label lblOtro 
            Alignment       =   1  'Right Justify
            Caption         =   "Fondo de Caja :"
            Height          =   240
            Index           =   10
            Left            =   120
            TabIndex        =   223
            Top             =   4860
            Width           =   1920
         End
         Begin VB.Label lblOLE 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   6930
            TabIndex        =   222
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblOLN 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   5490
            TabIndex        =   221
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Otras Cancelaciones Liquidación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   37
            Left            =   5175
            TabIndex        =   220
            Top             =   360
            Width           =   2850
         End
         Begin VB.Label lblODE 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   10035
            TabIndex        =   219
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblODN 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   8595
            TabIndex        =   218
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Otras Cancelaciones Diferencias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   38
            Left            =   8280
            TabIndex        =   217
            Top             =   360
            Width           =   2850
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   3690
         Picture         =   "frmLiquidacionAutorizaDetalle.frx":00A8
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   5400
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   9945
         Picture         =   "frmLiquidacionAutorizaDetalle.frx":05DA
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   5400
         Width           =   1305
      End
      Begin VB.TextBox txtTotalCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   150
         Text            =   "0.00"
         Top             =   4260
         Width           =   1365
      End
      Begin VB.TextBox txtTotalCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   149
         Text            =   "0.00"
         Top             =   4260
         Width           =   1365
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Aprobación"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   -65145
         Picture         =   "frmLiquidacionAutorizaDetalle.frx":06CC
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   5475
         Width           =   1305
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Aprobación"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   -65145
         Picture         =   "frmLiquidacionAutorizaDetalle.frx":0C56
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   3517
         Width           =   1305
      End
      Begin VB.TextBox txtDiferencia 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   8370
         Locked          =   -1  'True
         TabIndex        =   145
         Text            =   "0.00"
         Top             =   4260
         Width           =   1365
      End
      Begin VB.CheckBox chkControler 
         Caption         =   "Controler"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74820
         TabIndex        =   144
         Top             =   4365
         Width           =   2130
      End
      Begin VB.CheckBox chkAdministrador 
         Caption         =   "Administrador"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74820
         TabIndex        =   143
         Top             =   2385
         Width           =   2130
      End
      Begin VB.TextBox txtControler 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74820
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   142
         Top             =   4635
         Width           =   9510
      End
      Begin VB.TextBox txtAdministrador 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74820
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   141
         Top             =   2677
         Width           =   9510
      End
      Begin VB.TextBox txtOTotalDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   -66510
         Locked          =   -1  'True
         TabIndex        =   133
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtOTotalDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   -65070
         Locked          =   -1  'True
         TabIndex        =   132
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtOTotalLN 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   -69615
         Locked          =   -1  'True
         TabIndex        =   131
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtOTotalLE 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   -68175
         Locked          =   -1  'True
         TabIndex        =   130
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtTotalTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtTotalTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   -65190
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtTotalTarjetaLN 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtTotalTarjetaLE 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   126
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   8
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   125
         Text            =   "0.00"
         Top             =   4485
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   8
         Left            =   -65190
         Locked          =   -1  'True
         TabIndex        =   124
         Text            =   "0.00"
         Top             =   4485
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   -65190
         Locked          =   -1  'True
         TabIndex        =   120
         Text            =   "0.00"
         Top             =   4080
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   -65190
         Locked          =   -1  'True
         TabIndex        =   119
         Text            =   "0.00"
         Top             =   3675
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   -65190
         Locked          =   -1  'True
         TabIndex        =   118
         Text            =   "0.00"
         Top             =   3270
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   4
         Left            =   -65190
         Locked          =   -1  'True
         TabIndex        =   117
         Text            =   "0.00"
         Top             =   2865
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   -65190
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "0.00"
         Top             =   2460
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   -65190
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "0.00"
         Top             =   2055
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "0.00"
         Top             =   4080
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "0.00"
         Top             =   3675
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "0.00"
         Top             =   3270
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   4
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "0.00"
         Top             =   2865
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   110
         Text            =   "0.00"
         Top             =   2460
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   "0.00"
         Top             =   2055
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   -66630
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   -65190
         TabIndex        =   107
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   8
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   "0.00"
         Top             =   4485
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   8
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "0.00"
         Top             =   4485
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   7
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "0.00"
         Top             =   4080
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   6
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "0.00"
         Top             =   3675
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   5
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "0.00"
         Top             =   3270
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   4
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "0.00"
         Top             =   2865
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   3
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "0.00"
         Top             =   2460
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   2
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "0.00"
         Top             =   2055
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   7
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   95
         Text            =   "0.00"
         Top             =   4080
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   6
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   94
         Text            =   "0.00"
         Top             =   3675
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   5
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "0.00"
         Top             =   3270
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   4
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   92
         Text            =   "0.00"
         Top             =   2865
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   3
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   91
         Text            =   "0.00"
         Top             =   2460
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   2
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "0.00"
         Top             =   2055
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   1
         Left            =   -69735
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   1
         Left            =   -68295
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtOTotalCN 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtOTotalCE 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   -71280
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtTotalTarjetaCN 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtTotalTarjetaCE 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "0.00"
         Top             =   5745
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   8
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "0.00"
         Top             =   4485
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   8
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "0.00"
         Top             =   4485
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   7
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "0.00"
         Top             =   4080
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   6
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "0.00"
         Top             =   3675
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   5
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "0.00"
         Top             =   3270
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   4
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "0.00"
         Top             =   2865
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   3
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "0.00"
         Top             =   2460
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   2
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "0.00"
         Top             =   2055
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   7
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "0.00"
         Top             =   4080
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   6
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   3675
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   5
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "0.00"
         Top             =   3270
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   4
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "0.00"
         Top             =   2865
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   3
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "0.00"
         Top             =   2460
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   2
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "0.00"
         Top             =   2055
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   1
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtTarjetaCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   1
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtPuntoDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   9810
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   3765
         Width           =   1365
      End
      Begin VB.TextBox txtPuntoDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8370
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   3765
         Width           =   1365
      End
      Begin VB.TextBox txtPuntoLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "0.00"
         Top             =   3765
         Width           =   1365
      End
      Begin VB.TextBox txtPuntoLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "0.00"
         Top             =   3765
         Width           =   1365
      End
      Begin VB.TextBox txtOtroDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   9810
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   3342
         Width           =   1365
      End
      Begin VB.TextBox txtOtroDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8370
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "0.00"
         Top             =   3342
         Width           =   1365
      End
      Begin VB.TextBox txtOtroLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "0.00"
         Top             =   3342
         Width           =   1365
      End
      Begin VB.TextBox txtOtroLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   3342
         Width           =   1365
      End
      Begin VB.TextBox txtChequeLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   2919
         Width           =   1365
      End
      Begin VB.TextBox txtChequeLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "0.00"
         Top             =   2919
         Width           =   1365
      End
      Begin VB.TextBox txtChequeDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8370
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   2919
         Width           =   1365
      End
      Begin VB.TextBox txtChequeDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   9810
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   2919
         Width           =   1365
      End
      Begin VB.TextBox txtCreditoLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   2496
         Width           =   1365
      End
      Begin VB.TextBox txtCreditoLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   2496
         Width           =   1365
      End
      Begin VB.TextBox txtEfectivoLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   2073
         Width           =   1365
      End
      Begin VB.TextBox txtCreditoDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8370
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   2496
         Width           =   1365
      End
      Begin VB.TextBox txtCreditoDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   9810
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   2496
         Width           =   1365
      End
      Begin VB.TextBox txtEfectivoLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   2073
         Width           =   1365
      End
      Begin VB.TextBox txtEfectivoDN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8370
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   2073
         Width           =   1365
      End
      Begin VB.TextBox txtEfectivoDE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   9810
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   2073
         Width           =   1365
      End
      Begin VB.TextBox txtFondoLE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   330
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtFondoLN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   330
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtPuntoCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   3765
         Width           =   1365
      End
      Begin VB.TextBox txtPuntoCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   3765
         Width           =   1365
      End
      Begin VB.TextBox txtOtroCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   3342
         Width           =   1365
      End
      Begin VB.TextBox txtOtroCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   3342
         Width           =   1365
      End
      Begin VB.TextBox txtChequeCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   2919
         Width           =   1365
      End
      Begin VB.TextBox txtChequeCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   2919
         Width           =   1365
      End
      Begin VB.TextBox txtCreditoCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   2496
         Width           =   1365
      End
      Begin VB.TextBox txtCreditoCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   2496
         Width           =   1365
      End
      Begin VB.TextBox txtEfectivoCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   2073
         Width           =   1365
      End
      Begin VB.TextBox txtEfectivoCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   2073
         Width           =   1365
      End
      Begin VB.TextBox txtFondoCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtFondoCN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   330
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1650
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74775
         TabIndex        =   154
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Diferencia :"
         Height          =   240
         Index           =   6
         Left            =   6480
         TabIndex        =   153
         Top             =   4305
         Width           =   1590
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo Final en Caja :"
         Height          =   240
         Index           =   7
         Left            =   495
         TabIndex        =   151
         Top             =   4305
         Width           =   1590
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "( Cobrado + Recibos - Egresos )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   13
         Left            =   90
         TabIndex        =   138
         Top             =   2280
         Width           =   1980
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Tarjetas de Crédito Diferencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   36
         Left            =   -66630
         TabIndex        =   123
         Top             =   840
         Width           =   2805
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -66360
         TabIndex        =   122
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Propinas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -64920
         TabIndex        =   121
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Tarjetas de Crédito Liquidación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   -69735
         TabIndex        =   104
         Top             =   840
         Width           =   2805
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69465
         TabIndex        =   103
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Propinas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -68025
         TabIndex        =   102
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Cancelaciones :"
         Height          =   240
         Index           =   34
         Left            =   -74565
         TabIndex        =   87
         Top             =   5790
         Width           =   1770
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Tarjetas :"
         Height          =   240
         Index           =   33
         Left            =   -74640
         TabIndex        =   84
         Top             =   5790
         Width           =   1725
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Tarjetas de Crédito Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   -72885
         TabIndex        =   79
         Top             =   840
         Width           =   2850
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -72570
         TabIndex        =   78
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Propinas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71130
         TabIndex        =   77
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   8
         Left            =   -74640
         TabIndex        =   64
         Top             =   4530
         Width           =   1725
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   7
         Left            =   -74640
         TabIndex        =   63
         Top             =   4125
         Width           =   1725
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   6
         Left            =   -74640
         TabIndex        =   62
         Top             =   3720
         Width           =   1725
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   5
         Left            =   -74640
         TabIndex        =   61
         Top             =   3315
         Width           =   1725
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   4
         Left            =   -74640
         TabIndex        =   60
         Top             =   2910
         Width           =   1725
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   3
         Left            =   -74640
         TabIndex        =   59
         Top             =   2505
         Width           =   1725
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   2
         Left            =   -74640
         TabIndex        =   58
         Top             =   2100
         Width           =   1725
      End
      Begin VB.Label lblTarjeta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   1
         Left            =   -74640
         TabIndex        =   57
         Top             =   1695
         Width           =   1725
      End
      Begin VB.Label lblDE 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10080
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblDN 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8640
         TabIndex        =   28
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblLE 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6975
         TabIndex        =   27
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblLN 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5490
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblCE 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3870
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblCN 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2475
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Diferencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   9585
         TabIndex        =   23
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Liquidación Cajero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   6030
         TabIndex        =   22
         Top             =   840
         Width           =   2040
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Comprobante Cajero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   2880
         TabIndex        =   21
         Top             =   840
         Width           =   2085
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Puntos :"
         Height          =   240
         Index           =   5
         Left            =   495
         TabIndex        =   6
         Top             =   3810
         Width           =   1590
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Otros Tipo de Pago :"
         Height          =   240
         Index           =   4
         Left            =   495
         TabIndex        =   5
         Top             =   3390
         Width           =   1590
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque / Depósito :"
         Height          =   240
         Index           =   3
         Left            =   495
         TabIndex        =   4
         Top             =   2970
         Width           =   1590
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Tarjetas de Crédito :"
         Height          =   240
         Index           =   2
         Left            =   495
         TabIndex        =   3
         Top             =   2535
         Width           =   1590
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Efectivo :"
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   2
         Top             =   2070
         Width           =   1590
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Fondo de Caja :"
         Height          =   240
         Index           =   0
         Left            =   495
         TabIndex        =   1
         Top             =   1695
         Width           =   1590
      End
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha Final :"
      Height          =   240
      Index           =   15
      Left            =   3825
      TabIndex        =   148
      Top             =   427
      Width           =   1545
   End
   Begin VB.Label lblTipoCambio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10080
      TabIndex        =   140
      Top             =   45
      Width           =   1320
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo de Cambio :"
      Height          =   240
      Index           =   14
      Left            =   8505
      TabIndex        =   139
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label lblCaja 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10080
      TabIndex        =   137
      Top             =   405
      Width           =   1320
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Caja :"
      Height          =   240
      Index           =   12
      Left            =   9180
      TabIndex        =   136
      Top             =   427
      Width           =   780
   End
   Begin VB.Label lblTurno 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   990
      TabIndex        =   135
      Top             =   45
      Width           =   1995
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario :"
      Height          =   240
      Index           =   11
      Left            =   135
      TabIndex        =   134
      Top             =   427
      Width           =   780
   End
   Begin VB.Label lblFechaFin 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5467
      TabIndex        =   54
      Top             =   405
      Width           =   2130
   End
   Begin VB.Label lblFechaIni 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5460
      TabIndex        =   53
      Top             =   45
      Width           =   2130
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha Inicial :"
      Height          =   240
      Index           =   18
      Left            =   3825
      TabIndex        =   52
      Top             =   60
      Width           =   1545
   End
   Begin VB.Label lblUsuario 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   990
      TabIndex        =   51
      Top             =   405
      Width           =   1995
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Turno :"
      Height          =   240
      Index           =   17
      Left            =   135
      TabIndex        =   50
      Top             =   60
      Width           =   780
   End
End
Attribute VB_Name = "frmLiquidacionAutorizacionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTarjetas As Recordset
Dim RsOtros As Recordset
Dim RsLiquidacion As Recordset
Dim RsIngreso As Recordset

Dim sTurno As String
Dim nFondoN As Double
Dim nFondoE As Double
Dim nEfectivoCN As Double
Dim nEfectivoLN As Double
Dim nEfectivoCE As Double
Dim nEfectivoLE As Double
Dim nCreditoCN As Double
Dim nCreditoLN As Double
Dim nCreditoCE As Double
Dim nCreditoLE As Double
Dim nChequeCN As Double
Dim nChequeLN As Double
Dim nChequeCE As Double
Dim nChequeLE As Double
Dim nOTotalCN As Double
Dim nOTotalLN As Double
Dim nOTotalCE As Double
Dim nOTotalLE As Double
Dim nPuntoCN As Double
Dim nPuntoLN As Double
Dim nPuntoCE As Double
Dim nPuntoLE As Double
Dim nIngresoN As Double
Dim nIngresoE As Double
Dim nEgresoLE As Double
Dim nEgresoLN As Double

Dim nTotalCN As Double
Dim nTotalLN As Double
Dim nTotalCE As Double
Dim nTotalLE As Double

Dim tTarjeta(8) As String
Dim nTarjetaCN(8) As Double
Dim nTarjetaLN(8) As Double
Dim nTarjetaCE(8) As Double
Dim nTarjetaLE(8) As Double

Dim tOtro(20) As String
Dim nOtroCN(20) As Double
Dim nOtroLN(20) As Double
Dim nOtroCE(20) As Double
Dim nOtroLE(20) As Double

'
'Dim tOtro(10) As String
'Dim nOtroCN(10) As Double
'Dim nOtroLN(10) As Double
'Dim nOtroCE(10) As Double
'Dim nOtroLE(10) As Double

Dim nSaldoCN As Double
Dim nSaldoLN As Double
Dim nSaldoCE As Double
Dim nSaldoLE As Double
Dim nTC As Double
Dim nDolar As Double
Dim nDiferencia As Double

Dim i As Integer

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0
               If Supervisor("11") = False Then
                 MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               Cn.Execute "update MTURNO set lAdministrador=1, tAdministrador='" & txtAdministrador.Text & "', tAdministradorUsuario='" & sVar1 & "', nDiferencia= " & nDiferencia & " where tTurno='" & sTurno & "'"
               frmLiquidacion.RsCabecera.Requery
               
          Case Is = 1
               If Supervisor("11") = False Then
                 MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
               Cn.Execute "update MTURNO set lControler=1, tControler='" & txtControler.Text & "', tControlerUsuario='" & sVar1 & "', nDiferencia= " & nDiferencia & " where tTurno='" & sTurno & "'"
               frmLiquidacion.RsCabecera.Requery
               
          Case Is = 3
              ImprimeLiquidacion
   End Select
   
   Unload Me
End Sub

Private Sub Form_Load()
   Centrar Me
   Tarjetas
   Otros
   AsignarComprobante
   AsignarLiquidacion
   diferencia
   If sTipo = "ADM" Then
      txtAdministrador.Enabled = True
      chkAdministrador.Enabled = True
      cmdOpcion(0).Enabled = True
   Else
      txtControler.Enabled = True
      chkControler.Enabled = True
      cmdOpcion(1).Enabled = True
   End If
   
End Sub

Public Sub AsignarComprobante()
   Screen.MousePointer = vbHourglass
      
   With frmLiquidacion.RsCabecera
        'Cuadro de Texto
        sTurno = IIf(IsNull(!tTurno), "", !tTurno)
        lblTurno.Caption = IIf(IsNull(!tTurno), "", !tTurno)
        lblUsuario.Caption = IIf(IsNull(!tUsuario), "", !tUsuario)
        lblFechaIni.Caption = IIf(IsNull(!finicial), "", !finicial)
        lblFechaFin.Caption = IIf(IsNull(!fFinal), "", !fFinal)
        lblCaja.Caption = IIf(IsNull(!tCaja), "", !tCaja)
        
        chkAdministrador.value = IIf(IsNull(!lAdministrador), 0, IIf(!lAdministrador = True, 1, 0))
        chkControler.value = IIf(IsNull(!lControler), 0, IIf(!lControler = True, 1, 0))
        txtAdministrador.Text = IIf(IsNull(!tAdministrador), "", !tAdministrador)
        txtControler.Text = IIf(IsNull(!tControler), "", !tControler)
        txtObservacion.Text = IIf(IsNull(!tObservacion), "", !tObservacion)
   End With
                
   'Descripciones MN y ME
   lblCN.Caption = sMonN
   lblLN.Caption = sMonN
   lblDN.Caption = sMonN
   lblCE.Caption = sMonE
   lblLE.Caption = sMonE
   lblDE.Caption = sMonE
   
   lblOCN(0).Caption = sMonN
   lblOLN(0).Caption = sMonN
   lblODN(0).Caption = sMonN
   lblOCE(0).Caption = sMonE
   lblOLE(0).Caption = sMonE
   lblODE(0).Caption = sMonE
   
   lblOCN(1).Caption = sMonN
   lblOLN(1).Caption = sMonN
   lblODN(1).Caption = sMonN
   lblOCE(1).Caption = sMonE
   lblOLE(1).Caption = sMonE
   lblODE(1).Caption = sMonE
   
   'Quitar ME
   If sMonE = "" And sMonN = sMonE Then
      QuitarME
      nTC = 1
   Else
      nTC = Calcular("select nVenta as Codigo from TTIPOCAMBIO where fFecha='" & Format(lblFechaIni.Caption, "yyyy/mm/dd") & "'", Cn)
   End If
   'lblTipoCambio.Caption = Format(nTC, "##0.000")
   lblTipoCambio.Caption = Format(nTC, "##0.00")
                                   
   With frmLiquidacion.RsCabecera
        'Fondo de Caja
        nFondoN = IIf(IsNull(!nMontoIN), 0, !nMontoIN)
        nFondoE = IIf(IsNull(!nMontoIE), 0, !nMontoIE)
        txtFondoCN.Text = Format(nFondoN, "###,###,##0.00")
        txtFondoCE.Text = Format(nFondoE, "###,###,##0.00")
        
        'Efectivo
        nEfectivoCN = IIf(IsNull(!nMontoEN), 0, !nMontoEN)
        nEfectivoCE = IIf(IsNull(!nMontoEE), 0, !nMontoEE)
        txtEfectivoCN.Text = Format(nEfectivoCN, "###,###,##0.00")
        txtEfectivoCE.Text = Format(nEfectivoCE, "###,###,##0.00")
        
        'Tarjeta
        nTarjetaCN(1) = IIf(IsNull(!nTarjeta1), 0, !nTarjeta1)
        nTarjetaCN(2) = IIf(IsNull(!nTarjeta2), 0, !nTarjeta2)
        nTarjetaCN(3) = IIf(IsNull(!nTarjeta3), 0, !nTarjeta3)
        nTarjetaCN(4) = IIf(IsNull(!nTarjeta4), 0, !nTarjeta4)
        nTarjetaCN(5) = IIf(IsNull(!nTarjeta5), 0, !nTarjeta5)
        nTarjetaCN(6) = IIf(IsNull(!nTarjeta6), 0, !nTarjeta6)
        nTarjetaCN(7) = IIf(IsNull(!nTarjeta7), 0, !nTarjeta7)
        nTarjetaCN(8) = IIf(IsNull(!nTarjeta8), 0, !nTarjeta8)
        nCreditoCN = nTarjetaCN(1) + nTarjetaCN(2) + nTarjetaCN(3) + nTarjetaCN(4) + nTarjetaCN(5) + nTarjetaCN(6) + nTarjetaCN(7) + nTarjetaCN(8)
        
        txtTarjetaCN(1).Text = Format(nTarjetaCN(1), "###,###,##0.00")
        txtTarjetaCN(2).Text = Format(nTarjetaCN(2), "###,###,##0.00")
        txtTarjetaCN(3).Text = Format(nTarjetaCN(3), "###,###,##0.00")
        txtTarjetaCN(4).Text = Format(nTarjetaCN(4), "###,###,##0.00")
        txtTarjetaCN(5).Text = Format(nTarjetaCN(5), "###,###,##0.00")
        txtTarjetaCN(6).Text = Format(nTarjetaCN(6), "###,###,##0.00")
        txtTarjetaCN(7).Text = Format(nTarjetaCN(7), "###,###,##0.00")
        txtTarjetaCN(8).Text = Format(nTarjetaCN(8), "###,###,##0.00")
        txtCreditoCN.Text = Format(nCreditoCN, "###,###,##0.00")
        txtTotalTarjetaCN.Text = Format(nCreditoCN, "###,###,##0.00")
                
        nTarjetaCE(1) = IIf(IsNull(!nPropina1), 0, !nPropina1)
        nTarjetaCE(2) = IIf(IsNull(!nPropina2), 0, !nPropina2)
        nTarjetaCE(3) = IIf(IsNull(!nPropina3), 0, !nPropina3)
        nTarjetaCE(4) = IIf(IsNull(!nPropina4), 0, !nPropina4)
        nTarjetaCE(5) = IIf(IsNull(!nPropina5), 0, !nPropina5)
        nTarjetaCE(6) = IIf(IsNull(!nPropina6), 0, !nPropina6)
        nTarjetaCE(7) = IIf(IsNull(!nPropina7), 0, !nPropina7)
        nTarjetaCE(8) = IIf(IsNull(!nPropina8), 0, !nPropina8)
        nCreditoCE = nTarjetaCE(1) + nTarjetaCE(2) + nTarjetaCE(3) + nTarjetaCE(4) + nTarjetaCE(5) + nTarjetaCE(6) + nTarjetaCE(7) + nTarjetaCE(8)

        txtTarjetaCE(1).Text = Format(nTarjetaCE(1), "###,###,##0.00")
        txtTarjetaCE(2).Text = Format(nTarjetaCE(2), "###,###,##0.00")
        txtTarjetaCE(3).Text = Format(nTarjetaCE(3), "###,###,##0.00")
        txtTarjetaCE(4).Text = Format(nTarjetaCE(4), "###,###,##0.00")
        txtTarjetaCE(5).Text = Format(nTarjetaCE(5), "###,###,##0.00")
        txtTarjetaCE(6).Text = Format(nTarjetaCE(6), "###,###,##0.00")
        txtTarjetaCE(7).Text = Format(nTarjetaCE(7), "###,###,##0.00")
        txtTarjetaCE(8).Text = Format(nTarjetaCE(8), "###,###,##0.00")
        txtCreditoCE.Text = Format(nCreditoCE, "###,###,##0.00")
        txtTotalTarjetaCE.Text = Format(nCreditoCE, "###,###,##0.00")

        'Cheque
        nChequeCN = IIf(IsNull(!nMontoCN), 0, !nMontoCN)
        nChequeCE = IIf(IsNull(!nMontoCE), 0, !nMontoCE)
        txtChequeCN.Text = Format(nChequeCN, "###,###,##0.00")
        txtChequeCE.Text = Format(nChequeCE, "###,###,##0.00")
        
        'Otro
        nOtroCN(1) = IIf(IsNull(!notron1), 0, !notron1)
        nOtroCN(2) = IIf(IsNull(!nOtroN2), 0, !nOtroN2)
        nOtroCN(3) = IIf(IsNull(!nOtroN3), 0, !nOtroN3)
        nOtroCN(4) = IIf(IsNull(!nOtroN4), 0, !nOtroN4)
        nOtroCN(5) = IIf(IsNull(!nOtroN5), 0, !nOtroN5)
        nOtroCN(6) = IIf(IsNull(!nOtroN6), 0, !nOtroN6)
        nOtroCN(7) = IIf(IsNull(!nOtroN7), 0, !nOtroN7)
        nOtroCN(8) = IIf(IsNull(!nOtroN8), 0, !nOtroN8)
        nOtroCN(9) = IIf(IsNull(!nOtroN9), 0, !nOtroN9)
        nOtroCN(10) = IIf(IsNull(!nOtroN10), 0, !nOtroN10)
        
        nOtroCN(11) = IIf(IsNull(!notron11), 0, !notron11)
        nOtroCN(12) = IIf(IsNull(!nOtroN12), 0, !nOtroN12)
        nOtroCN(13) = IIf(IsNull(!nOtroN13), 0, !nOtroN13)
        nOtroCN(14) = IIf(IsNull(!nOtroN14), 0, !nOtroN14)
        nOtroCN(15) = IIf(IsNull(!nOtroN15), 0, !nOtroN15)
        nOtroCN(16) = IIf(IsNull(!nOtroN16), 0, !nOtroN16)
        nOtroCN(17) = IIf(IsNull(!nOtroN17), 0, !nOtroN17)
        nOtroCN(18) = IIf(IsNull(!nOtroN18), 0, !nOtroN18)
        nOtroCN(19) = IIf(IsNull(!nOtroN19), 0, !nOtroN19)
        nOtroCN(20) = IIf(IsNull(!nOtroN20), 0, !nOtroN20)
        nOTotalCN = nOtroCN(1) + nOtroCN(2) + nOtroCN(3) + nOtroCN(4) + nOtroCN(5) + nOtroCN(6) + nOtroCN(7) + nOtroCN(8) + nOtroCN(9) + nOtroCN(10) + nOtroCN(11) + nOtroCN(12) + nOtroCN(13) + nOtroCN(14) + nOtroCN(15) + nOtroCN(16) + nOtroCN(17) + nOtroCN(18) + nOtroCN(19) + nOtroCN(20)
        
        txtOCN(1).Text = Format(nOtroCN(1), "###,###,##0.00")
        txtOCN(2).Text = Format(nOtroCN(2), "###,###,##0.00")
        txtOCN(3).Text = Format(nOtroCN(3), "###,###,##0.00")
        txtOCN(4).Text = Format(nOtroCN(4), "###,###,##0.00")
        txtOCN(5).Text = Format(nOtroCN(5), "###,###,##0.00")
        txtOCN(6).Text = Format(nOtroCN(6), "###,###,##0.00")
        txtOCN(7).Text = Format(nOtroCN(7), "###,###,##0.00")
        txtOCN(8).Text = Format(nOtroCN(8), "###,###,##0.00")
        txtOCN(9).Text = Format(nOtroCN(9), "###,###,##0.00")
        txtOCN(10).Text = Format(nOtroCN(10), "###,###,##0.00")
        
        txtOCN(11).Text = Format(nOtroCN(11), "###,###,##0.00")
        txtOCN(12).Text = Format(nOtroCN(12), "###,###,##0.00")
        txtOCN(13).Text = Format(nOtroCN(13), "###,###,##0.00")
        txtOCN(14).Text = Format(nOtroCN(14), "###,###,##0.00")
        txtOCN(15).Text = Format(nOtroCN(15), "###,###,##0.00")
        txtOCN(16).Text = Format(nOtroCN(16), "###,###,##0.00")
        txtOCN(17).Text = Format(nOtroCN(17), "###,###,##0.00")
        txtOCN(18).Text = Format(nOtroCN(18), "###,###,##0.00")
        txtOCN(19).Text = Format(nOtroCN(19), "###,###,##0.00")
        txtOCN(20).Text = Format(nOtroCN(20), "###,###,##0.00")
        
        txtOTotalCN.Text = Format(nOTotalCN, "###,###,##0.00")
        
        txtOtroCN.Text = Format(nOTotalCN, "###,###,##0.00")
        
        
        nOtroCE(1) = IIf(IsNull(!nOtroE1), 0, !nOtroE1)
        nOtroCE(2) = IIf(IsNull(!nOtroE2), 0, !nOtroE2)
        nOtroCE(3) = IIf(IsNull(!nOtroE3), 0, !nOtroE3)
        nOtroCE(4) = IIf(IsNull(!nOtroE4), 0, !nOtroE4)
        nOtroCE(5) = IIf(IsNull(!nOtroE5), 0, !nOtroE5)
        nOtroCE(6) = IIf(IsNull(!nOtroE6), 0, !nOtroE6)
        nOtroCE(7) = IIf(IsNull(!nOtroE7), 0, !nOtroE7)
        nOtroCE(8) = IIf(IsNull(!nOtroE8), 0, !nOtroE8)
        nOtroCE(9) = IIf(IsNull(!nOtroE9), 0, !nOtroE9)
        nOtroCE(10) = IIf(IsNull(!nOtroE10), 0, !nOtroE10)
        
        nOtroCE(11) = IIf(IsNull(!nOtroE11), 0, !nOtroE11)
        nOtroCE(12) = IIf(IsNull(!nOtroE12), 0, !nOtroE12)
        nOtroCE(13) = IIf(IsNull(!nOtroE13), 0, !nOtroE13)
        nOtroCE(14) = IIf(IsNull(!nOtroE14), 0, !nOtroE14)
        nOtroCE(15) = IIf(IsNull(!nOtroE15), 0, !nOtroE15)
        nOtroCE(16) = IIf(IsNull(!nOtroE16), 0, !nOtroE16)
        nOtroCE(17) = IIf(IsNull(!nOtroE17), 0, !nOtroE17)
        nOtroCE(18) = IIf(IsNull(!nOtroE18), 0, !nOtroE18)
        nOtroCE(19) = IIf(IsNull(!nOtroE19), 0, !nOtroE19)
        nOtroCE(20) = IIf(IsNull(!nOtroE20), 0, !nOtroE20)
        
        nOTotalCE = nOtroCE(1) + nOtroCE(2) + nOtroCE(3) + nOtroCE(4) + nOtroCE(5) + nOtroCE(6) + nOtroCE(7) + nOtroCE(8) + nOtroCE(9) + nOtroCE(10) + nOtroCE(11) + nOtroCE(12) + nOtroCE(13) + nOtroCE(14) + nOtroCE(15) + nOtroCE(16) + nOtroCE(17) + nOtroCE(18) + nOtroCE(19) + nOtroCE(20)
        
        txtOCE(1).Text = Format(nOtroCE(1), "###,###,##0.00")
        txtOCE(2).Text = Format(nOtroCE(2), "###,###,##0.00")
        txtOCE(3).Text = Format(nOtroCE(3), "###,###,##0.00")
        txtOCE(4).Text = Format(nOtroCE(4), "###,###,##0.00")
        txtOCE(5).Text = Format(nOtroCE(5), "###,###,##0.00")
        txtOCE(6).Text = Format(nOtroCE(6), "###,###,##0.00")
        txtOCE(7).Text = Format(nOtroCE(7), "###,###,##0.00")
        txtOCE(8).Text = Format(nOtroCE(8), "###,###,##0.00")
        txtOCE(9).Text = Format(nOtroCE(9), "###,###,##0.00")
        txtOCE(10).Text = Format(nOtroCE(10), "###,###,##0.00")
        
        txtOCE(11).Text = Format(nOtroCE(11), "###,###,##0.00")
        txtOCE(12).Text = Format(nOtroCE(12), "###,###,##0.00")
        txtOCE(13).Text = Format(nOtroCE(13), "###,###,##0.00")
        txtOCE(14).Text = Format(nOtroCE(14), "###,###,##0.00")
        txtOCE(15).Text = Format(nOtroCE(15), "###,###,##0.00")
        txtOCE(16).Text = Format(nOtroCE(16), "###,###,##0.00")
        txtOCE(17).Text = Format(nOtroCE(17), "###,###,##0.00")
        txtOCE(18).Text = Format(nOtroCE(18), "###,###,##0.00")
        txtOCE(19).Text = Format(nOtroCE(19), "###,###,##0.00")
        txtOCE(20).Text = Format(nOtroCE(20), "###,###,##0.00")
        
        txtOTotalCE.Text = Format(nOTotalCE, "###,###,##0.00")
        txtOtroCE.Text = Format(nOTotalCE, "###,###,##0.00")
        
        'Punto
        nPuntoCN = IIf(IsNull(!nMontoPN), 0, !nMontoPN)
        nPuntoCE = IIf(IsNull(!nMontoPE), 0, !nMontoPE)
        txtPuntoCN.Text = Format(nPuntoCN, "###,###,##0.00")
        txtPuntoCE.Text = Format(nPuntoCE, "###,###,##0.00")
                
        'Saldo
        nSaldoCN = IIf(IsNull(!nMontoFN), 0, !nMontoFN)
        nSaldoCE = IIf(IsNull(!nMontoFE), 0, !nMontoFE)
        txtTotalCN.Text = Format(nSaldoCN, "###,###,##0.00")
        txtTotalCE.Text = Format(nSaldoCE, "###,###,##0.00")
        
   End With
   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignarLiquidacion()
   Screen.MousePointer = vbHourglass
   Isql = "select tTipoPago, tOtroTipoPago, tMoneda, tTarjeta, sum(nMonto) as nVenta1, sum(nPropina) as nVenta2, sum(nDolar) as nVenta3 from dpagodocumento " & _
          "where tTurno='" & sTurno & "' Group by tTipoPago, tOtroTipoPago, tMoneda, tTarjeta"
         
   Set RsLiquidacion = Lib.OpenRecordset(Isql, Cn)
     
   With RsLiquidacion
        'Efectivo
        nEfectivoLN = 0
        nEfectivoLE = 0
        nDolar = 0
        .Filter = "tTipoPago='01'"
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              If !tMoneda = "01" Then
                 nEfectivoLN = IIf(IsNull(!nVenta1), 0, !nVenta1)
              Else
                 nEfectivoLE = IIf(IsNull(!nVenta1), 0, !nVenta1)
                 nDolar = IIf(IsNull(!nVenta3), 0, !nVenta3)
              End If
              .MoveNext
           Loop
        End If
        
        If nEfectivoLE > 0 Then
           nEfectivoLN = nEfectivoLN - ((nDolar - nEfectivoLE) * nTC)
        End If
           
        'Tarjeta
        nCreditoLN = 0
        nCreditoLE = 0
        For i = 1 To 8
            nTarjetaLN(i) = 0
            nTarjetaLE(i) = 0
        Next i
        .Filter = "tTipoPago='02'"
        If .RecordCount > 0 Then
           For i = 1 To 8
               .MoveFirst
               .Find "tTarjeta='" & tTarjeta(i) & "'"
               If Not .EOF Then
                  nTarjetaLN(i) = IIf(IsNull(!nVenta1), 0, !nVenta1)
                  nTarjetaLE(i) = IIf(IsNull(!nVenta2), 0, !nVenta2)
               Else
                  nTarjetaLN(i) = 0
                  nTarjetaLE(i) = 0
               End If
               nCreditoLN = nCreditoLN + nTarjetaLN(i)
               nCreditoLE = nCreditoLE + nTarjetaLE(i)
           Next i
        End If
           
        'Deposito
        nChequeLN = 0
        nChequeLE = 0
        For i = 1 To 20
            nOtroLN(i) = 0
            nOtroLE(i) = 0
        Next i
        .Filter = "tTipoPago='03'"
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              If !tMoneda = "01" Then
                 nChequeLN = IIf(IsNull(!nVenta1), 0, !nVenta1)
              Else
                 nChequeLE = IIf(IsNull(!nVenta2), 0, !nVenta2)
              End If
              .MoveNext
           Loop
        End If
                
        'Otros
        nOTotalLN = 0
        nOTotalLE = 0
        .Filter = "tTipoPago='04'"
        If .RecordCount > 0 Then
           For i = 1 To 20
               .MoveFirst
               .Find "tOtroTipoPago='" & tOtro(i) & "'"
               If Not .EOF Then
                  nOtroLN(i) = IIf(IsNull(!nVenta1), 0, !nVenta1)
                  nOtroLE(i) = IIf(IsNull(!nVenta2), 0, !nVenta2)
               Else
                  nOtroLN(i) = 0
                  nOtroLE(i) = 0
               End If
               nOTotalLN = nOTotalLN + nOtroLN(i)
               nOTotalLE = nOTotalLE + nOtroLE(i)
           Next i
        End If
        
        'Puntos
        nPuntoLN = 0
        nPuntoLE = 0
        .Filter = "tTipoPago='05'"
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              If !tMoneda = "01" Then
                 nPuntoLN = IIf(IsNull(!nVenta1), 0, !nVenta1)
              Else
                 nPuntoLE = IIf(IsNull(!nVenta2), 0, !nVenta2)
              End If
              .MoveNext
           Loop
        End If
   End With
   
   'Recibos de Ingreso
   nIngresoN = 0
   nIngresoE = 0
   'Isql = "select tMoneda, sum(nMonto) as nMonto From mingreso where tEstadoDocumento='01' and tTurno='" & sTurno & "' Group by tMoneda"
   Isql = "select ttipopago,tMoneda,tTarjeta, sum(nMonto) as nMonto From mingreso where tEstadoDocumento='01' and tTurno='" & sTurno & "' Group by tMoneda,ttipopago,tTarjeta"
   Set RsIngreso = Lib.OpenRecordset(Isql, Cn)
   With RsIngreso
    If RsIngreso.RecordCount > 0 Then
       .Filter = "tTipoPago='01'"
       Do While Not RsIngreso.EOF
          If RsIngreso!tMoneda = "01" Then
             nIngresoN = IIf(IsNull(RsIngreso!nMonto), 0, RsIngreso!nMonto)
          Else
             nIngresoE = IIf(IsNull(RsIngreso!nMonto), 0, RsIngreso!nMonto)
          End If
          RsIngreso.MoveNext
       Loop
       .Filter = "tTipoPago='02'"
        If .RecordCount > 0 Then
           For i = 1 To 8
               .MoveFirst
               .Find "tTarjeta='" & tTarjeta(i) & "'"
               If Not .EOF Then
                 nTarjetaLN(i) = nTarjetaLN(i) + IIf(IsNull(!nMonto), 0, !nMonto)
                 If !tMoneda = "01" Then
                  nCreditoLN = nCreditoLN + IIf(IsNull(!nMonto), 0, !nMonto)
                 End If
               End If
           Next i
        End If
    End If
   End With
'   Set RsIngreso = Lib.OpenRecordset(Isql, Cn)
'   If RsIngreso.RecordCount > 0 Then
'      Do While Not RsIngreso.EOF
'         If RsIngreso!tMoneda = "01" Then
'            nIngresoN = IIf(IsNull(RsIngreso!nMonto), 0, RsIngreso!nMonto)
'         Else
'            nIngresoE = IIf(IsNull(RsIngreso!nMonto), 0, RsIngreso!nMonto)
'         End If
'         RsIngreso.MoveNext
'      Loop
'   End If
   
    'nEgresoLN = Calcular("select sum(nMonto) as Codigo from MEGRESO where tTurno='" & sTurno & "' and tMoneda='01'", Cn)
    'nEgresoLE = Calcular("select sum(nMonto) as Codigo from MEGRESO where tTurno='" & sTurno & "' and tMoneda='02'", Cn)
   
    nEgresoLN = Calcular("select sum(nMonto) as Codigo from MEGRESO where tTurno='" & sTurno & "' and tMoneda='01' AND TESTADODOCUMENTO<>'04'", Cn)
    nEgresoLE = Calcular("select sum(nMonto) as Codigo from MEGRESO where tTurno='" & sTurno & "' and tMoneda='02' AND TESTADODOCUMENTO<>'04'", Cn)

    nEfectivoLN = nEfectivoLN - nEgresoLN + nIngresoN
    nDolar = nDolar - nEgresoLE + nIngresoE
    
    'Fondo de Caja
    txtFondoLN.Text = Format(nFondoN, "###,###,##0.00")
    txtFondoLE.Text = Format(nFondoE, "###,###,##0.00")
        
    'Efectivo
    txtEfectivoLN.Text = Format(nEfectivoLN, "###,###,##0.00")
    txtEfectivoLE.Text = Format(nDolar, "###,###,##0.00")
        
    'Tarjeta
    txtTarjetaLN(1).Text = Format(nTarjetaLN(1), "###,###,##0.00")
    txtTarjetaLN(2).Text = Format(nTarjetaLN(2), "###,###,##0.00")
    txtTarjetaLN(3).Text = Format(nTarjetaLN(3), "###,###,##0.00")
    txtTarjetaLN(4).Text = Format(nTarjetaLN(4), "###,###,##0.00")
    txtTarjetaLN(5).Text = Format(nTarjetaLN(5), "###,###,##0.00")
    txtTarjetaLN(6).Text = Format(nTarjetaLN(6), "###,###,##0.00")
    txtTarjetaLN(7).Text = Format(nTarjetaLN(7), "###,###,##0.00")
    txtTarjetaLN(8).Text = Format(nTarjetaLN(8), "###,###,##0.00")
    txtCreditoLN.Text = Format(nCreditoLN, "###,###,##0.00")
    txtTotalTarjetaLN.Text = Format(nCreditoLN, "###,###,##0.00")
                
    txtTarjetaLE(1).Text = Format(nTarjetaLE(1), "###,###,##0.00")
    txtTarjetaLE(2).Text = Format(nTarjetaLE(2), "###,###,##0.00")
    txtTarjetaLE(3).Text = Format(nTarjetaLE(3), "###,###,##0.00")
    txtTarjetaLE(4).Text = Format(nTarjetaLE(4), "###,###,##0.00")
    txtTarjetaLE(5).Text = Format(nTarjetaLE(5), "###,###,##0.00")
    txtTarjetaLE(6).Text = Format(nTarjetaLE(6), "###,###,##0.00")
    txtTarjetaLE(7).Text = Format(nTarjetaLE(7), "###,###,##0.00")
    txtTarjetaLE(8).Text = Format(nTarjetaLE(8), "###,###,##0.00")
    txtCreditoLE.Text = Format(nCreditoLE, "###,###,##0.00")
    txtTotalTarjetaLE.Text = Format(nCreditoLE, "###,###,##0.00")

    'Cheque
    txtChequeLN.Text = Format(nChequeLN, "###,###,##0.00")
    txtChequeLE.Text = Format(nChequeLE, "###,###,##0.00")
    
    'Otro
    txtOLN(1).Text = Format(nOtroLN(1), "###,###,##0.00")
    txtOLN(2).Text = Format(nOtroLN(2), "###,###,##0.00")
    txtOLN(3).Text = Format(nOtroLN(3), "###,###,##0.00")
    txtOLN(4).Text = Format(nOtroLN(4), "###,###,##0.00")
    txtOLN(5).Text = Format(nOtroLN(5), "###,###,##0.00")
    txtOLN(6).Text = Format(nOtroLN(6), "###,###,##0.00")
    txtOLN(7).Text = Format(nOtroLN(7), "###,###,##0.00")
    txtOLN(8).Text = Format(nOtroLN(8), "###,###,##0.00")
    txtOLN(9).Text = Format(nOtroLN(9), "###,###,##0.00")
    txtOLN(10).Text = Format(nOtroLN(10), "###,###,##0.00")
    
    txtOLN(11).Text = Format(nOtroLN(11), "###,###,##0.00")
    txtOLN(12).Text = Format(nOtroLN(12), "###,###,##0.00")
    txtOLN(13).Text = Format(nOtroLN(13), "###,###,##0.00")
    txtOLN(14).Text = Format(nOtroLN(14), "###,###,##0.00")
    txtOLN(15).Text = Format(nOtroLN(15), "###,###,##0.00")
    txtOLN(16).Text = Format(nOtroLN(16), "###,###,##0.00")
    txtOLN(17).Text = Format(nOtroLN(17), "###,###,##0.00")
    txtOLN(18).Text = Format(nOtroLN(18), "###,###,##0.00")
    txtOLN(19).Text = Format(nOtroLN(19), "###,###,##0.00")
    txtOLN(20).Text = Format(nOtroLN(20), "###,###,##0.00")
    
    txtOTotalLN.Text = Format(nOTotalLN, "###,###,##0.00")
    txtOtroLN.Text = Format(nOTotalLN, "###,###,##0.00")
    
    txtOLE(1).Text = Format(nOtroLE(1), "###,###,##0.00")
    txtOLE(2).Text = Format(nOtroLE(2), "###,###,##0.00")
    txtOLE(3).Text = Format(nOtroLE(3), "###,###,##0.00")
    txtOLE(4).Text = Format(nOtroLE(4), "###,###,##0.00")
    txtOLE(5).Text = Format(nOtroLE(5), "###,###,##0.00")
    txtOLE(6).Text = Format(nOtroLE(6), "###,###,##0.00")
    txtOLE(7).Text = Format(nOtroLE(7), "###,###,##0.00")
    txtOLE(8).Text = Format(nOtroLE(8), "###,###,##0.00")
    txtOLE(9).Text = Format(nOtroLE(9), "###,###,##0.00")
    txtOLE(10).Text = Format(nOtroLE(10), "###,###,##0.00")
    
    txtOLE(11).Text = Format(nOtroLE(11), "###,###,##0.00")
    txtOLE(12).Text = Format(nOtroLE(12), "###,###,##0.00")
    txtOLE(13).Text = Format(nOtroLE(13), "###,###,##0.00")
    txtOLE(14).Text = Format(nOtroLE(14), "###,###,##0.00")
    txtOLE(15).Text = Format(nOtroLE(15), "###,###,##0.00")
    txtOLE(16).Text = Format(nOtroLE(16), "###,###,##0.00")
    txtOLE(17).Text = Format(nOtroLE(17), "###,###,##0.00")
    txtOLE(18).Text = Format(nOtroLE(18), "###,###,##0.00")
    txtOLE(19).Text = Format(nOtroLE(19), "###,###,##0.00")
    txtOLE(20).Text = Format(nOtroLE(20), "###,###,##0.00")
    
    txtOTotalLE.Text = Format(nOTotalLE, "###,###,##0.00")
    txtOtroLE.Text = Format(nOTotalLE, "###,###,##0.00")
    
    'Punto
    txtPuntoLN.Text = Format(nPuntoLN, "###,###,##0.00")
    txtPuntoLE.Text = Format(nPuntoLE, "###,###,##0.00")
   
   Screen.MousePointer = vbDefault
End Sub

Public Sub diferencia()
   Screen.MousePointer = vbHourglass
            
    'Efectivo
    txtEfectivoDN.Text = Format(nEfectivoLN - nEfectivoCN, "###,###,##0.00")
    txtEfectivoDE.Text = Format(nDolar - nEfectivoCE, "###,###,##0.00")
        
    'Tarjeta
    txtTarjetaDN(1).Text = Format(nTarjetaLN(1) - nTarjetaCN(1), "###,###,##0.00")
    txtTarjetaDN(2).Text = Format(nTarjetaLN(2) - nTarjetaCN(2), "###,###,##0.00")
    txtTarjetaDN(3).Text = Format(nTarjetaLN(3) - nTarjetaCN(3), "###,###,##0.00")
    txtTarjetaDN(4).Text = Format(nTarjetaLN(4) - nTarjetaCN(4), "###,###,##0.00")
    txtTarjetaDN(5).Text = Format(nTarjetaLN(5) - nTarjetaCN(5), "###,###,##0.00")
    txtTarjetaDN(6).Text = Format(nTarjetaLN(6) - nTarjetaCN(6), "###,###,##0.00")
    txtTarjetaDN(7).Text = Format(nTarjetaLN(7) - nTarjetaCN(7), "###,###,##0.00")
    txtTarjetaDN(8).Text = Format(nTarjetaLN(8) - nTarjetaCN(8), "###,###,##0.00")
    txtCreditoDN.Text = Format(nCreditoLN - nCreditoCN, "###,###,##0.00")
    txtTotalTarjetaDN.Text = Format(nCreditoLN - nCreditoCN, "###,###,##0.00")
                
    txtTarjetaDE(1).Text = Format(nTarjetaLE(1) - nTarjetaCE(1), "###,###,##0.00")
    txtTarjetaDE(2).Text = Format(nTarjetaLE(2) - nTarjetaCE(2), "###,###,##0.00")
    txtTarjetaDE(3).Text = Format(nTarjetaLE(3) - nTarjetaCE(3), "###,###,##0.00")
    txtTarjetaDE(4).Text = Format(nTarjetaLE(4) - nTarjetaCE(4), "###,###,##0.00")
    txtTarjetaDE(5).Text = Format(nTarjetaLE(5) - nTarjetaCE(5), "###,###,##0.00")
    txtTarjetaDE(6).Text = Format(nTarjetaLE(6) - nTarjetaCE(6), "###,###,##0.00")
    txtTarjetaDE(7).Text = Format(nTarjetaLE(7) - nTarjetaCE(7), "###,###,##0.00")
    txtTarjetaDE(8).Text = Format(nTarjetaLE(8) - nTarjetaCE(8), "###,###,##0.00")
    txtCreditoDE.Text = Format(nCreditoLE - nCreditoCE, "###,###,##0.00")
    txtTotalTarjetaDE.Text = Format(nCreditoLE - nCreditoCE, "###,###,##0.00")

    'Cheque
    txtChequeDN.Text = Format(nChequeLN - nChequeCN, "###,###,##0.00")
    txtChequeDE.Text = Format(nChequeLE - nChequeCE, "###,###,##0.00")
    
    'Otro
    txtODN(1).Text = Format(nOtroLN(1) - nOtroCN(1), "###,###,##0.00")
    txtODN(2).Text = Format(nOtroLN(2) - nOtroCN(2), "###,###,##0.00")
    txtODN(3).Text = Format(nOtroLN(3) - nOtroCN(3), "###,###,##0.00")
    txtODN(4).Text = Format(nOtroLN(4) - nOtroCN(4), "###,###,##0.00")
    txtODN(5).Text = Format(nOtroLN(5) - nOtroCN(5), "###,###,##0.00")
    txtODN(6).Text = Format(nOtroLN(6) - nOtroCN(6), "###,###,##0.00")
    txtODN(7).Text = Format(nOtroLN(7) - nOtroCN(7), "###,###,##0.00")
    txtODN(8).Text = Format(nOtroLN(8) - nOtroCN(8), "###,###,##0.00")
    txtODN(9).Text = Format(nOtroLN(9) - nOtroCN(9), "###,###,##0.00")
    txtODN(10).Text = Format(nOtroLN(10) - nOtroCN(10), "###,###,##0.00")
    
    txtODN(11).Text = Format(nOtroLN(11) - nOtroCN(11), "###,###,##0.00")
    txtODN(12).Text = Format(nOtroLN(12) - nOtroCN(12), "###,###,##0.00")
    txtODN(13).Text = Format(nOtroLN(13) - nOtroCN(13), "###,###,##0.00")
    txtODN(14).Text = Format(nOtroLN(14) - nOtroCN(14), "###,###,##0.00")
    txtODN(15).Text = Format(nOtroLN(15) - nOtroCN(15), "###,###,##0.00")
    txtODN(16).Text = Format(nOtroLN(16) - nOtroCN(16), "###,###,##0.00")
    txtODN(17).Text = Format(nOtroLN(17) - nOtroCN(17), "###,###,##0.00")
    txtODN(18).Text = Format(nOtroLN(18) - nOtroCN(18), "###,###,##0.00")
    txtODN(19).Text = Format(nOtroLN(19) - nOtroCN(19), "###,###,##0.00")
    txtODN(20).Text = Format(nOtroLN(20) - nOtroCN(20), "###,###,##0.00")
    
    txtOTotalDN.Text = Format(nOTotalLN - nOTotalCN, "###,###,##0.00")
    
    txtOtroDN.Text = Format(nOTotalLN - nOTotalCN, "###,###,##0.00")
    
    txtODE(1).Text = Format(nOtroLE(1) - nOtroCE(1), "###,###,##0.00")
    txtODE(2).Text = Format(nOtroLE(2) - nOtroCE(2), "###,###,##0.00")
    txtODE(3).Text = Format(nOtroLE(3) - nOtroCE(3), "###,###,##0.00")
    txtODE(4).Text = Format(nOtroLE(4) - nOtroCE(4), "###,###,##0.00")
    txtODE(5).Text = Format(nOtroLE(5) - nOtroCE(5), "###,###,##0.00")
    txtODE(6).Text = Format(nOtroLE(6) - nOtroCE(6), "###,###,##0.00")
    txtODE(7).Text = Format(nOtroLE(7) - nOtroCE(7), "###,###,##0.00")
    txtODE(8).Text = Format(nOtroLE(8) - nOtroCE(8), "###,###,##0.00")
    txtODE(9).Text = Format(nOtroLE(9) - nOtroCE(9), "###,###,##0.00")
    txtODE(10).Text = Format(nOtroLE(10) - nOtroCE(10), "###,###,##0.00")
    
    txtODE(11).Text = Format(nOtroLE(11) - nOtroCE(11), "###,###,##0.00")
    txtODE(12).Text = Format(nOtroLE(12) - nOtroCE(12), "###,###,##0.00")
    txtODE(13).Text = Format(nOtroLE(13) - nOtroCE(13), "###,###,##0.00")
    txtODE(14).Text = Format(nOtroLE(14) - nOtroCE(14), "###,###,##0.00")
    txtODE(15).Text = Format(nOtroLE(15) - nOtroCE(15), "###,###,##0.00")
    txtODE(16).Text = Format(nOtroLE(16) - nOtroCE(16), "###,###,##0.00")
    txtODE(17).Text = Format(nOtroLE(17) - nOtroCE(17), "###,###,##0.00")
    txtODE(18).Text = Format(nOtroLE(18) - nOtroCE(18), "###,###,##0.00")
    txtODE(19).Text = Format(nOtroLE(19) - nOtroCE(19), "###,###,##0.00")
    txtODE(20).Text = Format(nOtroLE(20) - nOtroCE(20), "###,###,##0.00")
    
    txtOTotalDE.Text = Format(nOTotalLE - nOTotalCE, "###,###,##0.00")
    txtOtroDE.Text = Format(nOTotalLE - nOTotalCE, "###,###,##0.00")
    
    'Punto
    txtPuntoDN.Text = Format(nPuntoLN - nPuntoCN, "###,###,##0.00")
    txtPuntoDE.Text = Format(nPuntoLE - nPuntoCE, "###,###,##0.00")
        
    'Total Diferencia
    nDiferencia = (nEfectivoLN - nEfectivoCN) + (nCreditoLN - nCreditoCN) + (nChequeLN - nChequeCN) + (nOTotalLN - nOTotalCN) + (nPuntoLN - nPuntoCN) + (((nDolar - nEfectivoCE) + (nChequeLE - nChequeCE) + (nOTotalLE - nOTotalCE) + (nPuntoLE - nPuntoCE)) * nTC)
    
    txtDiferencia.Text = Format(nDiferencia, "###,###,##0.00")
   
   Screen.MousePointer = vbDefault
End Sub
Public Sub Tarjetas()
   Set RsTarjetas = Lib.OpenRecordset("select tCodigoTarjeta, tDetallado, tResumido from tTarjetaCredito", Cn)
   For i = 1 To 8
       RsTarjetas.MoveFirst
       RsTarjetas.Find "tcodigoTarjeta = '0" & Trim(str(i)) & "'"
       If RsTarjetas.EOF Then
          lblTarjeta(i).Visible = False
          txtTarjetaCN(i).Visible = False
          txtTarjetaCE(i).Visible = False
          txtTarjetaLN(i).Visible = False
          txtTarjetaLE(i).Visible = False
          txtTarjetaDN(i).Visible = False
          txtTarjetaDE(i).Visible = False
       Else
          lblTarjeta(i).Caption = RsTarjetas!tDetallado
          tTarjeta(i) = RsTarjetas!tCodigoTarjeta
       End If
   Next i
End Sub

Public Sub Otros()
   Set RsOtros = Lib.OpenRecordset("select * from vTipoCancelacion where lActivo=1", Cn)
   If RsOtros.RecordCount = 0 Then
      Exit Sub
   End If
   
   RsOtros.MoveFirst
   
   For i = 1 To 20
       If RsOtros.EOF Then
          lblOtro(i).Visible = False
          txtOCN(i).Visible = False
          txtOCE(i).Visible = False
          txtOLN(i).Visible = False
          txtOLE(i).Visible = False
          txtODN(i).Visible = False
          txtODE(i).Visible = False
       Else
          lblOtro(i).Caption = RsOtros!Descripcion
          tOtro(i) = RsOtros!codigo
          RsOtros.MoveNext
       End If
   Next i
End Sub

Public Sub QuitarME()
    lblCE.Visible = False
    lblLE.Visible = False
    lblDE.Visible = False
    lblOCE(0).Visible = False
    lblOLE(0).Visible = False
    lblODE(0).Visible = False
    lblOCE(1).Visible = False
    lblOLE(1).Visible = False
    lblODE(1).Visible = False
    
    txtFondoCE.Visible = False
    txtFondoLE.Visible = False
    txtEfectivoCE.Visible = False
    txtEfectivoLE.Visible = False
    txtEfectivoDE.Visible = False
    txtCreditoCE.Visible = False
    txtCreditoLE.Visible = False
    txtCreditoDE.Visible = False
    txtChequeCE.Visible = False
    txtChequeLE.Visible = False
    txtChequeDE.Visible = False
    txtOtroCE.Visible = False
    txtOtroLE.Visible = False
    txtOtroDE.Visible = False
    txtPuntoCE.Visible = False
    txtPuntoLE.Visible = False
    txtPuntoDE.Visible = False
    txtTotalCE.Visible = False
    
    
    For i = 1 To 20
        txtOCE(i).Visible = False
        txtOLE(i).Visible = False
        txtODE(i).Visible = False
    Next i
End Sub

Public Sub ImprimeLiquidacion()
   Dim sTitulo1 As String
   Dim sTitulo2 As String
   Dim sTitulo3 As String
   
   Screen.MousePointer = vbHourglass
   'Configura la impresora la impresion Font
   Imprimir (sPreCuenta)
   Printer.FontName = sFont
   Printer.FontBold = False
   
   'Cabecera
   ImprimeXCentro "Liquidación de Cajero", 40
   ImprimeXCentro sRazonSocial, 40
   Printer.Print ""
      
   sTitulo1 = ""
   sTitulo2 = ""
   sTitulo1 = "Turno   : " & sTurno
   sTitulo2 = "Fecha   : " & Format(Now, "dd MMMM yyyy HH:mm") & " Hrs "
                
   Printer.Print sTitulo1
   Printer.Print sTitulo2
   Printer.Print ""
   
   Printer.Print String(40, "-")
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "CONCEPTO                  " & sMonN & "        " & sMonE
   Else
      Printer.Print "CONCEPTO                  " & sMonN
   End If
   
   Printer.Print String(40, "-")
               
   'Fondo de Caja
   Printer.Print
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Fondo de Caja   : " & Right(String(11, " ") & Format(nFondoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nFondoE, "###,##0.00"), 11)
   Else
      Printer.Print "Fondo de Caja   : " & Right(String(11, " ") & Format(nFondoN, "###,##0.00"), 11)
   End If
      
   'Efectivo
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "       Efectivo : " & Right(String(11, " ") & Format(nEfectivoCN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nEfectivoCE, "###,##0.00"), 11)
   Else
      Printer.Print "       Efectivo : " & Right(String(11, " ") & Format(nEfectivoCN, "###,##0.00"), 11)
   End If
   
   'Total de Tarjetas
   Printer.Print ""
   Printer.Print "Tarjetas Credito: " & Right(String(11, " ") & Format(nCreditoCN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nCreditoCE, "###,##0.00"), 11)
   For i = 1 To 8
       If lblTarjeta(i).Visible = True Then
          Printer.Print " - " & Mid(IIf(lblTarjeta(i).Caption = "", "(no definido)", lblTarjeta(i).Caption) & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nTarjetaCN(i), "###,##0.00"), 11) & Right(String(11, " ") & Format(nTarjetaCE(i), "###,##0.00"), 11)
       End If
   Next i

   'Total de Cheques
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Cheque/Deposito : " & Right(String(11, " ") & Format(nChequeCN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nChequeCE, "###,##0.00"), 11)
   Else
      Printer.Print "Cheque/Deposito : " & Right(String(11, " ") & Format(nChequeCN, "###,##0.00"), 11)
   End If

   'Otros Tipos de Pago
   Printer.Print ""
   Printer.Print "Otros Tipos Pago: " & Right(String(11, " ") & Format(nOTotalCN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nOTotalCE, "###,##0.00"), 11)
   For i = 1 To 20
       If sMonE <> "" And sMonN <> sMonE Then
          If lblOtro(i).Visible = True Then
             Printer.Print " - " & Mid(IIf(lblOtro(i).Caption = "", "(no definido)", lblOtro(i).Caption) & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nOtroCN(i), "###,##0.00"), 11) & Right(String(11, " ") & Format(nOtroCE(i), "###,##0.00"), 11)
          End If
       Else
          If lblOtro(i).Visible = True Then
             Printer.Print " - " & Mid(IIf(lblOtro(i).Caption = "", "(no definido)", lblOtro(i).Caption) & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nOtroCN(i), "###,##0.00"), 11)
          End If
       End If
   Next i
 
   'Puntos
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Puntos          : " & Right(String(11, " ") & Format(nPuntoCN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nPuntoCE, "###,##0.00"), 11)
   Else
      Printer.Print "Puntos          : " & Right(String(11, " ") & Format(nPuntoCN, "###,##0.00"), 11)
   End If
   
   'Total Ingreso en Caja
   Printer.Print ""
   Printer.Print String(40, "-")
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Total Ingreso   : " & Right(String(11, " ") & Format(nEfectivoCN + nCreditoCN + nChequeCN + nOTotalCN + nPuntoCN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nEfectivoCE + nChequeCE + nOTotalCE, "###,##0.00"), 11)
   Else
      Printer.Print "Total Ingreso   : " & Right(String(11, " ") & Format(nEfectivoCN + nCreditoCN + nChequeCN + nOTotalCN + nPuntoCN, "###,##0.00"), 11)
   End If

   'Total Efectivo en Caja
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Total Efectivo  : " & Right(String(11, " ") & Format(nEfectivoCN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nEfectivoCE, "###,##0.00"), 11)
   Else
      Printer.Print "Total Efectivo  : " & Right(String(11, " ") & Format(nEfectivoCN, "###,##0.00"), 11)
   End If

   'Retiro de Caja
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Retiro Efectivo : " & Right(String(11, " ") & Format(0, "###,##0.00"), 11) & Right(String(11, " ") & Format(0, "###,##0.00"), 11)
   Else
      Printer.Print "Retiro Efectivo : " & Right(String(11, " ") & Format(0, "###,##0.00"), 11)
   End If

   'Saldo de Caja
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Saldo en Caja   : " & Right(String(11, " ") & Format(nSaldoCN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nSaldoCE, "###,##0.00"), 11)
   Else
      Printer.Print "Saldo en Caja   : " & Right(String(11, " ") & Format(nSaldoCN, "###,##0.00"), 11)
   End If

   Printer.Print " "
   Printer.Print String(40, "-")
   Printer.Print " "
   Printer.Print " "
   Printer.Print " "
   Printer.Print " "
   Printer.Print String(25, " ") & "---------------"
   Printer.Print String(25, " ") & String((15 - Len(RTrim(sUsuario))) / 2, " ") & sUsuario
   Printer.Print " "
   Printer.Print String(40, "-")
   
   Dim sDoc As String
   Dim sTipoDoc As String
   Dim nCantidadDoc As Integer
   
   For i = 1 To Calcular("select count(tCodigo) as codigo From TTABLA Where TTABLA = 'TIPODOCUMENTO'", Cn)
       sTipoDoc = Mid("00", 1, 2 - Len(Trim(str(i)))) & Trim(str(i))
       sDoc = Calcular("Select tDetallado As Codigo from TTABLA Where TTABLA = 'TIPODOCUMENTO' AND tCODIGO = '" & sTipoDoc & "'", Cn)
       nCantidadDoc = Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn)
       If sDoc <> "0" And nCantidadDoc > 0 Then
          Printer.Print " "
          Printer.Print sDoc
          Printer.Print "Del     : " & Format(Calcular("select min(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Al      : " & Format(Calcular("select max(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Emitido : " & Format(Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tEstadoDocumento<>'04' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Anulado : " & Format(Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tEstadoDocumento='04' and tTurno='" & sTurno & "'", Cn), "##,##0")
       End If
   Next i
   Printer.Print ""
   If Len(Trim(txtObservacion.Text)) > 0 Then
      Printer.Print String(40, "-")
      ImprimeXLinea txtObservacion.Text, 40, 0
   End If
   Printer.Print String(40, "-")
   Printer.Print "Caja    : " & sCaja
   Printer.Print " "
   Printer.EndDoc
   Screen.MousePointer = vbDefault
End Sub

