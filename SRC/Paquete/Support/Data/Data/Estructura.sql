
CREATE TABLE [dbo].[APEDIDO] (
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItem] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoGrupo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoSubGrupo] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPrecioNeto] [float] NULL ,
	[nPrecioImpuesto1] [float] NULL ,
	[nPrecioImpuesto2] [float] NULL ,
	[nPrecioImpuesto3] [float] NULL ,
	[nPrecioVenta] [float] NULL ,
	[nRecargo] [float] NULL ,
	[nDescuento] [float] NULL ,
	[nPrecioOficial] [float] NULL ,
	[nCantidad] [float] NULL ,
	[nImpuesto1] [float] NULL ,
	[nImpuesto2] [float] NULL ,
	[nImpuesto3] [float] NULL ,
	[nVenta] [float] NULL ,
	[tObservacion] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL ,
	[lImprime] [bit] NULL ,
	[tEstadoItem] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[lImprimeArea] [bit] NULL ,
	[tComanda] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tMotivoEliminacion] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tObservacionAnulado] [nvarchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuarioAnulado] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistroAnulado] [smalldatetime] NULL ,
	[tTurnoAnulado] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[fDiaContable] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CPEDIDO] (
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItem] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItemCombo] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tProductoCombo] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[nCantidad] [float] NULL ,
	[tCodigoGrupo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoSubGrupo] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPrecioNeto] [float] NULL ,
	[nImpuesto1] [float] NULL ,
	[nImpuesto2] [float] NULL ,
	[nImpuesto3] [float] NULL ,
	[nVenta] [float] NULL ,
	[nInsumo] [float] NULL ,
	[nGasto] [float] NULL ,
	[nManoObra] [float] NULL ,
	[lImprimeArea] [bit] NULL ,
	[lImprime] [bit] NULL ,
	[nOrden] [int] NULL ,
	[tObservacion] [nvarchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,
	[lCorte] [bit] NULL ,
	[lAtendidoC] [bit] NULL ,
	[fAtendidoC] [datetime] NULL ,
	[tUsuarioAtendio] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DDOCUMENTO] (
	[tDocumento] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItem] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPrecioNeto] [float] NULL ,
	[nPrecioImpuesto1] [float] NULL ,
	[nPrecioImpuesto2] [float] NULL ,
	[nPrecioImpuesto3] [float] NULL ,
	[nPrecioVenta] [float] NULL ,
	[nPrecioOficial] [float] NULL ,
	[nRecargo] [float] NULL ,
	[nDescuento] [float] NULL ,
	[nCantidad] [float] NULL ,
	[nImpuesto1] [float] NULL ,
	[nImpuesto2] [float] NULL ,
	[nImpuesto3] [float] NULL ,
	[nVenta] [float] NULL ,
	[tGuia] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DPAGODOCUMENTO] (
	[tDocumento] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCorrelativo] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoPago] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
        [tOtroTipoPago] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMoneda] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTipoCambio] [float] NULL ,
	[nMonto] [float] NULL ,
	[nPropina] [float] NULL ,
	[tTarjeta] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tReferencia] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBanco] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNumero] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFechaVencimiento] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[nDolar] [float] NULL ,
	[lReplica] [bit] NULL ,
	[fDiaContable] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DPEDIDO] (
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItem] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tTipoPedido] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoGrupo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoSubGrupo] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMoneda] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPrecioNeto] [float] NULL ,
	[nPrecioImpuesto1] [float] NULL ,
	[nPrecioImpuesto2] [float] NULL ,
	[nPrecioImpuesto3] [float] NULL ,
	[nPrecioVenta] [float] NULL ,
	[nRecargo] [float] NULL ,
	[nDescuento] [float] NULL ,
	[nPrecioOficial] [float] NULL ,
	[nCantidad] [float] NULL ,
	[nImpuesto1] [float] NULL ,
	[nImpuesto2] [float] NULL ,
	[nImpuesto3] [float] NULL ,
	[nVenta] [float] NULL ,
	[tObservacion] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCortesia] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[lImprime] [bit] NULL ,
	[tEstadoItem] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[lCombinacion] [bit] NULL ,
	[nCombinacion] [smallint] NULL ,
	[lImprimeArea] [bit] NULL ,
	[tFacturado] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDocumento] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[lTransferido] [bit] NULL ,
	[tComanda] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tMozoD] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuarioD] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[nInsumo] [float] NULL ,
	[nGasto] [float] NULL ,
	[nManoObra] [float] NULL ,
	[nOrden] [int] NULL ,
	[lCorte] [bit] NULL ,
	[tPosicion] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[fEnvio] [smalldatetime] NULL ,
	[nEnvio] [int] NULL ,
	[tUnidadNegocio] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tOferta] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[tAutorizaOferta] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL  ,
	[tSubalmacen] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL,
	[lRegistroVenta] [bit] NULL ,
	[fDiaContable] [smalldatetime] NULL ,
	[lAtendidoC] [bit] NULL ,
	[fAtendidoC] [datetime] NULL ,
	[tUsuarioAtendio] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL   ,
	[lCantadoC] [bit] NULL ,
	[fCantadoC] [datetime] NULL ,
	[lTipoEnvio] [bit] NULL ,
	[tGuiaTransporte] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL,
	[tCodigoEtiqueta] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL,
    [tCajaD] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL,
    [lNoCantado] [bit] NULL   
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DPREPAGO] (
	[tDocumento] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCorrelativo] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoPago] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
    [tOtroTipoPago] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMoneda] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTipoCambio] [float] NULL ,
	[nMonto] [float] NULL ,
	[nVuelto] [float] NULL ,
	[nPropina] [float] NULL ,
	[tTarjeta] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tReferencia] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBanco] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNumero] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFechaVencimiento] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tObservacion] [nvarchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[fDiaContable] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MCIERRE] (
	[tPeriodo] [nvarchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[lCierre] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MDOCUMENTO] (
	[tDocumento] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tTipoDocumento] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoCliente] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCortesia] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[nNeto] [float] NULL ,
	[nPrecioImpuesto1] [float] NULL ,
	[nPrecioImpuesto2] [float] NULL ,
	[nPrecioImpuesto3] [float] NULL ,
	[nVenta] [float] NULL ,
	[nRecargo] [float] NULL ,
	[nDescuento] [float] NULL ,
	[nPrecioOficial] [float] NULL ,
	[nPropina] [float] NULL ,
	[ntotal] [float] NULL ,
	[nAbono] [float] NULL ,
	[nVuelto] [float] NULL ,
	[tCortesiaPago] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEstadoDocumento] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tClientePago] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[fPago] [smalldatetime] NULL ,
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuarioAutoriza] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tUsuarioAnulado] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistroAnulado] [smalldatetime] NULL ,
	[tObservacion] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL ,
	[tSalon] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEmision] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[tConsumo] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL ,
	[lReplica] [bit] NULL  ,
	[tAutorizacion] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoControl] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL,
	[tImpresora] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tSerieImpresora] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL  ,
	[fDiaContable] [smalldatetime] NULL ,
	[tDescuento] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[fInicio] [smalldatetime] NULL ,	
	[fCaducidad] [smalldatetime] NULL ,
	[tContribuyenteEspecial] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL 	
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MEGRESO] (
	[tRecibo] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[fFecha] [smalldatetime] NULL ,
	[tMoneda] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTipoCambio] [float] NULL ,
	[nMonto] [float] NULL ,
	[tDescripcion] [nvarchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[tAutoriza] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEstadoDocumento] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lReplica] [bit] NULL ,
	[fDiaContable] [smalldatetime] NULL ,
	[tTipoEgreso] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MINGRESO] (
	[tRecibo] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fFecha] [smalldatetime] NULL ,
	[tMoneda] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoPago] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTarjeta] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tReferencia] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTipoCambio] [float] NULL ,
	[nMonto] [float] NULL ,
	[tDescripcion] [nvarchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[tAutoriza] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[lAnticipo] [bit] NULL ,
	[tEstadoDocumento] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lReplica] [bit] NULL ,
	[fDiaContable] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MNOTACREDITO] (
	[tNotaCredito] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fFecha] [smalldatetime] NULL ,
	[tDocumento] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[nNeto] [float] NULL ,
	[nImpuesto1] [float] NULL ,
	[nImpuesto2] [float] NULL ,
	[nImpuesto3] [float] NULL ,
	[nVenta] [float] NULL ,
	[tEstadoDocumento] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tUsuarioAnulado] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistroAnulado] [smalldatetime] NULL ,
	[tObservacion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[lReplica] [bit] NULL ,
	[fDiaContable] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MPEDIDO] (
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nCorrelativo] [int] NULL ,
	[tClienteDelivery] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[tClienteCtaCte] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[fFecha] [smalldatetime] NULL ,
	[tMoneda] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[nMonto] [float] NULL ,
	[tEstadoPedido] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoAtencion] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoPedido] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lPrioridad] [bit] NULL ,
	[tAnulacionPedido] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMesa] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[nMesa] [smallint] NULL ,
	[tMozo] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMotorizado] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tSalon] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[fProgramacion] [smalldatetime] NULL ,
	[nTiempo] [int] NULL ,
	[tObservacion] [nvarchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[nAdulto] [smallint] NULL ,
	[nNino] [smallint] NULL ,
    [tMotivoAnulacion] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tusuarioAnulado] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegAnulado] [smalldatetime] NULL ,
	[tObservacionAnulado] [nvarchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTurnoAnulado] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tClienteCorp] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTienda] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegCuenta] [smalldatetime] NULL ,
	[nPrecuenta] [int] NULL ,
	[nReimpresion] [int] NULL ,
	[tCajaAnterior] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTurnoAnterior] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tComanda] [nvarchar] (12) COLLATE Modern_Spanish_CI_AS NULL ,
	[tPuntoVenta] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tHabitacion] [nvarchar] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[tReserva] [nvarchar] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[tPasajero] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCompania] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[tContacto] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFichaPasajero] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoComanda] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nDescuento] [float] NULL ,
	[tDescuento] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tObservacionDescuento] [nvarchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuarioDescuento] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEmpacador] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[fEmpacador] [smalldatetime] NULL ,
	[fAsignacion] [smalldatetime] NULL ,
	[fSalida] [smalldatetime] NULL ,
	[fEntrega] [smalldatetime] NULL ,
	[fLlegada] [smalldatetime] NULL ,
	[tCodigoPedidoCD] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTiempoDelivery] [int] NULL ,
	[lReplica] [bit] NULL ,
	[fRegistroCD] [smalldatetime] NULL ,
	[fEntregaClienteCD] [smalldatetime] NULL ,
	[fDiaContable] [smalldatetime] NULL ,
	[lAtendidoC] [bit] NULL ,
	[fAtendidoC] [datetime] NULL ,
	[tUsuarioAtendio] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL  ,
	[tMaitre] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTarifaMotorizado] [float] NULL ,
	[nTarifaExtra] [int] NULL ,
	[tMotorizadoN] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTarifaMotorizadoN] [float] NULL ,
	[nTarifaExtraN] [int] NULL ,
	[tCodigoInvitado] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCodigoPariente] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MTURNO] (
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCaja] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tSalon] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[fInicial] [smalldatetime] NULL ,
	[fFinal] [smalldatetime] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[lCierre] [bit] NULL ,
	[nMontoIN] [float] NULL ,
	[nMontoIE] [float] NULL ,
	[nMontoEN] [float] NULL ,
	[nMontoEE] [float] NULL ,
	[nMontoCN] [float] NULL ,
	[nMontoCE] [float] NULL ,
	[nMontoPN] [float] NULL ,
	[nMontoPE] [float] NULL ,
	[nMontoFN] [float] NULL ,
	[nMontoFE] [float] NULL ,
	[nTarjeta1] [float] NULL ,
	[nTarjeta2] [float] NULL ,
	[nTarjeta3] [float] NULL ,
	[nTarjeta4] [float] NULL ,
	[nTarjeta5] [float] NULL ,
	[nTarjeta6] [float] NULL ,
	[nTarjeta7] [float] NULL ,
	[nTarjeta8] [float] NULL ,
	[nPropina1] [float] NULL ,
	[nPropina2] [float] NULL ,
	[nPropina3] [float] NULL ,
	[nPropina4] [float] NULL ,
	[nPropina5] [float] NULL ,
	[nPropina6] [float] NULL ,
	[nPropina7] [float] NULL ,
	[nPropina8] [float] NULL ,
	[nOtroN1] [float] NULL ,
	[nOtroN2] [float] NULL ,
	[nOtroN3] [float] NULL ,
	[nOtroN4] [float] NULL ,
	[nOtroN5] [float] NULL ,
	[nOtroN6] [float] NULL ,
	[nOtroN7] [float] NULL ,
	[nOtroN8] [float] NULL ,
	[nOtroN9] [float] NULL ,
	[nOtroN10] [float] NULL ,
	[nOtroE1] [float] NULL ,
	[nOtroE2] [float] NULL ,
	[nOtroE3] [float] NULL ,
	[nOtroE4] [float] NULL ,
	[nOtroE5] [float] NULL ,
	[nOtroE6] [float] NULL ,
	[nOtroE7] [float] NULL ,
	[nOtroE8] [float] NULL ,
	[nOtroE9] [float] NULL ,
	[nOtroE10] [float] NULL ,
	[tObservacion] [nvarchar] (2000) COLLATE Modern_Spanish_CI_AS NULL ,
	[nDiferencia] [float] NULL ,
	[lAdministrador] [bit] NULL ,
	[tAdministrador] [nvarchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,
	[tAdministradorUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[lControler] [bit] NULL ,
	[tControler] [nvarchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,
	[tControlerUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[lReplica] [bit] NULL ,
	[nOtroN11] [float] NULL ,
	[nOtroN12] [float] NULL ,
	[nOtroN13] [float] NULL ,
	[nOtroN14] [float] NULL ,
	[nOtroN15] [float] NULL ,
	[nOtroN16] [float] NULL ,
	[nOtroN17] [float] NULL ,
	[nOtroN18] [float] NULL ,
	[nOtroN19] [float] NULL ,
	[nOtroN20] [float] NULL ,
	[nOtroE11] [float] NULL ,
	[nOtroE12] [float] NULL ,
	[nOtroE13] [float] NULL ,
	[nOtroE14] [float] NULL ,
	[nOtroE15] [float] NULL ,
	[nOtroE16] [float] NULL ,
	[nOtroE17] [float] NULL ,
	[nOtroE18] [float] NULL ,
	[nOtroE19] [float] NULL ,
	[nOtroE20] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MGUIATRANSPORTE](
	[tGuiaTransporte] [nvarchar](15) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fFecha] [smalldatetime] NULL,
	[tCodigoDelivery] [nvarchar](7) COLLATE Modern_Spanish_CI_AS NULL ,	
	[tDestinatario] [nvarchar](5) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTienda] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NULL ,	
	[tTransportista] [nvarchar](5) COLLATE Modern_Spanish_CI_AS NULL ,	
	[tUnidadTransporte] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDocumento] [nvarchar](20) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEstadoGuia] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTotal] [float] NULL,
	[tCaja] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTurno] [nvarchar](10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tUsuarioAnulado] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistroAnulado] [smalldatetime] NULL 
	
 ) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DGUIATRANSPORTE](
	[tGuiaTransporte] [nvarchar](15) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItem] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoProducto] [nvarchar](7) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPrecioVenta] [float] NULL,
	[nCantidad] [float] NULL,
	[nVenta] [float] NULL,
	[tDocumento] [nvarchar](20) COLLATE Modern_Spanish_CI_AS NULL 
 ) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TAREAIMPRESORA] (
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tImpresora] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TCLIENTE] (
	[tCodigoCliente] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tEmpresa] [nvarchar] (80) COLLATE Modern_Spanish_CI_AS NULL ,
	[tIdentidad] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDireccion] [nvarchar] (200) COLLATE Modern_Spanish_CI_AS NULL ,
	[lActivo] [bit] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lReplica] [bit] NULL ,
	[tCorreo] [nvarchar] (400) COLLATE Modern_Spanish_CI_AS NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TCOMBO] (
	[tCombo] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nCantidad] [float] NULL ,
	[lFijo] [bit] NULL ,
	[lUnico] [bit] NULL ,
	[tEtiqueta] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nAumento] [float] NULL ,
	[lReplica] [bit] NULL ,
	[lEliminaFijo] [bit] NULL ,
	[nValor] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TCOMPANIA] (
	[tCodigoCliente] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[lEmpresa] [bit] NULL ,
	[tApeCom] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNomSoc] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoIdentidad] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tIdentidad] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDireccion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTelefono1] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTelefono2] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEmail] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[nConsumo] [float] NULL ,
	[nLinea] [float] NULL ,
	[lActivo] [bit] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lCliente] [bit] NULL ,
	[tTipoCtaCte] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tSubTipoCtaCte] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[lReplica] [bit] NULL 
) ON [PRIMARY]
GO
	
CREATE TABLE [dbo].[TDELIVERY](
	[tCodigoDelivery] [nvarchar](7) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tTipoCliente] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL,
	[tApellido] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[tNombre] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[tDireccion] [nvarchar](100) COLLATE Modern_Spanish_CI_AS NULL,
	[tTelefono] [nvarchar](15) COLLATE Modern_Spanish_CI_AS NULL,
	[tReferencia] [nvarchar](250) COLLATE Modern_Spanish_CI_AS NULL,
	[tZona] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NULL,
	[tDistrito] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NULL,
	[tCodigoCliente] [nvarchar](5) COLLATE Modern_Spanish_CI_AS NULL,
	[tCodigoTarjeta] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL,
	[tNumeroTarjeta] [nvarchar](16) COLLATE Modern_Spanish_CI_AS NULL,
	[tFechaTarjeta] [nvarchar](4) COLLATE Modern_Spanish_CI_AS NULL,
	[nDescuento] [float] NULL,
	[fNacimiento] [smalldatetime] NULL,
	[tEmail] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[tObservacion] [nvarchar](250) COLLATE Modern_Spanish_CI_AS NULL,
	[lActivo] [bit] NULL,
	[lPuntos] [bit] NULL,
	[nAcumulado] [float] NULL,
	[nUtilizado] [float] NULL,
	[nDisponible] [float] NULL,
	[tUsuario] [nvarchar](15) COLLATE Modern_Spanish_CI_AS NULL,
	[fRegistro] [smalldatetime] NULL,
	[fModificacion] [smalldatetime] NULL,
	[lReplica] [bit] NULL,
	[tEstadoFrecuente] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL,
	[lExcluyeProductos] [bit] NULL,
	[lClienteCtaCte] [bit] NULL,
	[nConsumo] [float] NULL,
	[nLinea] [float] NULL,
	[tTipoCtaCte] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL,
	[tSubTipoCtaCte] [nvarchar](4) COLLATE Modern_Spanish_CI_AS NULL,
	[tTipoIdentidad] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL,
	[tIdentidad] [nvarchar](20) COLLATE Modern_Spanish_CI_AS NULL ,	
	[iFoto] [image] NULL,
	[tAccionSocio] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[nLineaPorCobrar] [float] NULL,
	[nConsumoPorCobrar] [float] NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TGRUPO] (
	[tCodigoGrupo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (24) COLLATE Modern_Spanish_CI_AS NULL ,
	[nBoton] [int] NULL ,
	[lActivo] [bit] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[lReplica] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TIMPRESORA] (
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tImpresora] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDescripcion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDevice] [nvarchar] (80) COLLATE Modern_Spanish_CI_AS NULL ,
	[tRuta] [nvarchar] (80) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFont] [nvarchar] (40) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tFont1] [nvarchar] (40) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFont2] [nvarchar] (40) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNumeroSerie] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL  ,
	[nFontSizePrecuenta] [float] NULL,
	[nFontSizeEnvio] [float] NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TMESA] (
	[tCodigoMesa] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tSalon] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lFumador] [bit] NULL ,
	[tX] [int] NULL ,
	[tY] [int] NULL ,
	[nPersona] [smallint] NULL ,
	[lActivo] [bit] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tEstadoMesa] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lReplica] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TMODULO] (
	[tCodigoModulo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tSecuencia] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tFormulario] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMenu1] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMenu2] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMenu3] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMenu4] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMenu5] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMenu6] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tObjeto] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tObjetoDescripcion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TPARAMETRO] (
	[tIdentificacionTributaria] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tRazonSocial] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tRazonComercial] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDireccion] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTelefono] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEmail] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tWebPage] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMonedaN] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMonN] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMonedaE] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMonE] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTiempo] [smallint] NULL ,
	[nChkTiempo] [smallint] NULL ,
	[Impuesto1] [float] NULL ,
	[Impuesto2] [float] NULL ,
	[Impuesto3] [float] NULL ,
	[tImpuesto1] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tImpuesto2] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tImpuesto3] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[nCorrelativo] [int] NULL ,
	[nDelivery] [float] NULL ,
	[nLlevar] [float] NULL ,
	[nCanal4] [float] NULL ,
	[nCanal5] [float] NULL ,
	[tPie] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL ,
	[lBotonTrans] [bit] NULL ,
	[tElimina] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tPassword] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[nItem] [smallint] NULL ,
	[lLongitud] [bit] NULL ,
	[nLongitud] [int] NULL ,
	[lPrinter] [bit] NULL ,
	[lAlmacen] [bit] NULL ,
	[lRapido] [bit] NULL ,
	[tBoton1] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBoton2] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBoton3] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL , 
	[tBoton4] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBoton5] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tPiePreCuenta] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL ,
	[lInfhotel] [bit] NULL , 
	[tClub] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPunto] [float] NULL ,
	[lCierre] [bit] NULL ,
	[nDecimal] [int] NULL ,
	[nDias] [int] NULL ,
	[lEquivalencia] [bit] NULL ,
	[nCabecera] [int] NULL ,
	[nDetalle] [int] NULL ,
	[tVersion] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[lComboGeneral] [bit] NULL ,
	[nDiasDelivery] [int] NULL ,
	[nTiempoMinutoCD] [int] NULL , 
	[lMultilocal]	[bit] null ,
	[lKDS]	[bit] NULL ,
	[tOrderInfo]   [nvarchar] (600)  COLLATE Modern_Spanish_CI_AS NULL,
	[tOrderStatus] [nvarchar] (600)  COLLATE Modern_Spanish_CI_AS NULL,
	[tBump]	       [nvarchar] (600)  COLLATE Modern_Spanish_CI_AS NULL,
	[lDiaContableAutomatico] [bit] NULL ,
	[tHoraCierreDiaContable] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[lDiaContableManual] [bit] NULL ,
	[lClub] [bit] NULL ,
	[lImprimeDiaContable] [bit] NULL ,	
	[nItemGuia] [smallint] NULL ,	
	[nCabeceraGuia] [int] NULL ,
	[nDetalleGuia] [int] NULL 	,
	[nAsignacionMotorizado] [float] NULL ,
	[tTarifaActualMotorizado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuarioTarifa] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistroTarifa] [datetime] NULL ,
	[lEnvioChef] [bit] NULL	,
	[tContribuyenteEspecial] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,	
	[fContribuyenteEspecial] [smalldatetime] NULL ,
	[tDireccion2] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL ,
	[lMobileUnidadNegocio] [bit] NULL ,
	[lMobilePasswordCCaja] [bit] NULL ,
	[lActivaConsultaDescargo] [bit] NULL,
	[nCabeceraV] [int] NULL , 
	[nItemV] [int] NULL ,
	[nPieV] [int] NULL ,
	[lFacturacionE] [bit] NULL,
	[lControlUsuario] [bit] NULL   
) ON [PRIMARY]
GO

 

CREATE TABLE [dbo].[TPRODUCTO] (
	[tCodigoProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tGrupo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tSubGrupo] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoProducto] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (24) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMoneda] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lImpuesto1] [bit] NULL ,
	[lImpuesto2] [bit] NULL ,
	[lImpuesto3] [bit] NULL ,
	[lImpuesto4] [bit] NULL ,
	[lImpuesto5] [bit] NULL ,
	[lImpuesto6] [bit] NULL ,
	[lImpuesto7] [bit] NULL ,
	[lImpuesto8] [bit] NULL ,
	[lImpuesto9] [bit] NULL ,
	[lImpuesto10] [bit] NULL ,
	[lImpuesto11] [bit] NULL ,
	[lImpuesto12] [bit] NULL ,
	[lImpuesto13] [bit] NULL ,
	[lImpuesto14] [bit] NULL ,
	[lImpuesto15] [bit] NULL ,
	[tDescargo] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPrecioVenta] [float] NULL ,
	[nPrecioDelivery] [float] NULL ,
	[nPrecioLlevar] [float] NULL ,
	[nPrecioCanal4] [float] NULL ,
	[nPrecioCanal5] [float] NULL ,
	[tCortesia] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lModificable] [bit] NULL ,
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[lImprimeArea] [bit] NULL ,
	[lEspecial] [bit] NULL ,
	[lFijo] [bit] NULL ,
	[lActivo] [bit] NULL ,
	[lCombinacion] [bit] NULL ,
	[nCombinacion] [smallint] NULL ,
	[nBoton] [int] NULL ,
	[tIcono] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[oFoto] [image] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tEnlace] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[tCajaRapida] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nBotonRapido] [int] NULL ,
	[nInsumo] [float] NULL ,
	[nInsumo2] [float] NULL ,
	[nInsumo3] [float] NULL ,
	[nInsumo4] [float] NULL ,
	[nInsumo5] [float] NULL ,
	[nGasto] [float] NULL ,
	[nGasto2] [float] NULL ,
	[nGasto3] [float] NULL ,
	[nGasto4] [float] NULL ,
	[nGasto5] [float] NULL ,
	[nManoObra] [float] NULL ,
	[nManoObra2] [float] NULL ,
	[nManoObra3] [float] NULL ,
	[nManoObra4] [float] NULL ,
	[nManoObra5] [float] NULL ,
	[tBarra] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL  ,
	[lPropiedad] [bit] NULL ,
	[tInfhotel] [nvarchar] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[lDescuento] [bit] NULL ,
	[lLocal] [bit] NULL ,
	[lDelivery] [bit] NULL ,
	[lLlevar] [bit] NULL ,
	[lCanal4] [bit] NULL ,
	[lCanal5] [bit] NULL ,
	[tBotonPalm] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUnidadNegocio] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lReplica] [bit] NULL ,
	[lMultiArea] [bit] NULL ,
	[tAlternativa] [nvarchar] (24) COLLATE Modern_Spanish_CI_AS NULL ,
	[lControlInsumoCritico] [bit] NULL ,
	[tCodigoInsumo] [varchar](8) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nTiempo] [int] NULL ,
	[lBalanza] [bit] NULL 
) ON [PRIMARY] 
GO

CREATE TABLE [dbo].[TPROPIEDAD] (
	[tCodigoPropiedad] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[tOperador] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPrecio] [float] NULL ,
	[tEnlace] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[nInsumo] [float] NULL ,
	[nGasto] [float] NULL ,
	[nManoObra] [float] NULL ,
	[lActivo] [bit] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lReplica] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TRESERVA] (
	[tReserva] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fFecha] [smalldatetime] NULL ,
	[fHora] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[tApellido] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNombre] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTelefono] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPax] [int] NULL ,
	[tEstadoReserva] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tObservacion] [nvarchar] (250) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tPrioridad] [int] NULL , 
	[tPrioridad2] [bit] NULL ,
	[fFechaModificacion] [smalldatetime] NULL , 
	[tMesa] [nvarchar] (15),
	[fDiaContable] [smalldatetime] NULL 
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[TSUBGRUPO] (
	[tCodigoGrupo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoSubgrupo] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (24) COLLATE Modern_Spanish_CI_AS NULL ,
	[nBoton] [int] NULL ,
	[tIcono] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[lActivo] [bit] NULL ,
	[lImprimeArea] [bit] NULL ,
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[lImpuesto1] [bit] NULL ,
	[lImpuesto2] [bit] NULL ,
	[lImpuesto3] [bit] NULL ,
	[nOrden] [int] NULL ,
	[tAgrupacion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lReplica] [bit] NULL ,
	[tCuentaContable] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL,
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TTABLA] (
	[TTABLA] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[TCODIGO] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nTamano] [smallint] NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (24) COLLATE Modern_Spanish_CI_AS NULL ,
	[tIcono] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[nBoton] [smallint] NULL ,
	[tValor] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[nValor] [float] NULL ,
	[lActivo] [bit] NULL ,
	[lReplica] [bit] NULL,
	[tValor2] [nvarchar] (4000) COLLATE Modern_Spanish_CI_AS NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TTARJETACREDITO] (
	[tCodigoTarjeta] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[nFactorRetencion] [float] NULL ,
	[tRepresentante] [nvarchar] (40) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTelefono1] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTelefono2] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[nBoton] [smallint] NULL ,
	[tCuentaContable] [nvarchar] (18) COLLATE Modern_Spanish_CI_AS NULL ,
	[lPinPad] [bit] NULL ,
	[lActivo] [bit] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lReplica] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TTIENDA](
	[tCodigoDelivery] [nvarchar](7) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tCodigoTienda] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tNombre] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[tDireccion] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[tTelefono] [nvarchar](15) COLLATE Modern_Spanish_CI_AS NULL,
	[tEmail] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[tContacto] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[lActivo] [bit] NULL,
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TTIPOCAMBIO] (
	[fFecha] [smalldatetime] NOT NULL ,
	[nCompra] [float] NULL ,
	[nVenta] [float] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lReplica] [bit] NULL ,
	[nOficial] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TTIPODOCUMENTO](
	[tCodigoTipoDocumento] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tDescripcion] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[tPrefijo] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL,
	[tCodigoSunat] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL,
	[lPideCliente] [bit] NULL,
	[nMonto] [float] NULL,
	[lTransporte] [bit] NULL,
	[lRegistroVenta] [bit] NULL,
	[lActivo] [bit] NULL,
	[lCanjearNotaCredito] [bit] NULL,
	[lValidaRuc] [bit] NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TTIPODOCUMENTOIMPRESORA] (
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tImpresora] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tTipoEmision] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDescripcion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFormulario] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tSerie] [nvarchar] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUltimoNumero] [nvarchar] (9) COLLATE Modern_Spanish_CI_AS NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[lResumen] [bit] NULL ,
	[lImpuesto1] [bit] NULL ,
	[lImpuesto2] [bit] NULL ,
	[lImpuesto3] [bit] NULL ,
	[lEquivaDolares] [bit] NULL ,
	[tNumeroAutorizacion] [nvarchar](20) COLLATE Modern_Spanish_CI_AS NULL ,
	[fInicio] [smalldatetime] NULL ,
	[fCaducidad] [smalldatetime] NULL	
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TUSUARIO] (
	[tCodigoUsuario] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tGrupoUsuario] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tPassword] [nvarchar] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[lActivo] [bit] NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tBandaMagnetica] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL,
	[tHuella] [nvarchar] (4000) COLLATE Modern_Spanish_CI_AS NULL,
	[tUsuarioModifica] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL , 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TCAJA] (
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDescripcion] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tPrecuenta] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[lComanda] [bit] NULL ,
	[vComanda] [bit] NULL ,
	[lMotivoEliminaC] [bit] NULL ,
	[lMotivoElimina] [bit] NULL ,
	[lActivo] [bit] NULL ,
	[lRefresca] [bit] NULL , 
	[lPasswordC] [bit] NULL , 
	[lPassword] [bit] NULL , 
	[tGrupo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lConsumo1] [bit] NULL ,
	[lConsumo2] [bit] NULL ,
	[lConsumo3] [bit] NULL ,
	[lPrecuenta] [bit] NULL ,
	[lAdicion] [bit] NULL ,
	[lPrecuentaAgrupada] [bit] NULL ,
	[tTipoPedido] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lObliga] [bit] NULL ,
	[lMozo] [bit] NULL ,
	[lObligaPrinter] [bit] NULL ,
	[lPax] [bit] NULL ,
	[lObligaCierre] [bit] NULL ,
	[lFiltroTipoPedido] [bit] NULL ,
	[nPuerto] [int] NULL ,
	[tMensaje1] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMensaje2] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[lCancelacion] [bit] NULL ,
	[lDirecto] [bit] NULL ,
	[lObligaPrecuenta] [bit] NULL ,
	[lComboPrecuenta] [bit] NULL ,
	[lComboDocumento] [bit] NULL ,
	[lCambioMesa] [bit] NULL ,
	[lVisaNet] [bit] NULL ,
	[lImpuestoPrecuenta] [bit] NULL ,
	[lDocumentoAgrupado] [bit] NULL ,
	[lOrden] [bit] NULL ,
	[lValorCortesia] [bit] NULL ,
	[lObservacion] [bit] NULL ,
	[lCajaRapida] [bit] NULL ,
	[lPropiedadPrecuenta] [bit] NULL ,
	[lPropiedadDocumento] [bit] NULL ,
	[lPrecioNetoPrecuenta] [bit] NULL ,
	[nLimitePrecuenta] [smallint] NULL ,
	[tUnidadNegocio] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nLimiteReimpresion] [smallint] NULL ,
	[lPasswordTransferencia] [bit] NULL ,
	[lCD] [bit] NULL ,
	[lFechaEntregaDelivery] [bit] NULL ,
	[lMultiCajero] [bit] NULL ,
	[lMCPV] [bit] NULL ,
	[lCCVOX] [bit] NULL ,
	[lMotorizado] [bit] NULL ,
	[lEquivaDolaPrecuenta] [bit] NULL ,
	[tSubAlmacen] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[lObservacionPrecuenta] [bit] NULL ,
	[lObservacionDocumento] [bit] NULL ,
	[lPasswordImportarPedido] [bit] NULL ,
	[lActivaImpDscAlternativa] [bit] NULL ,
	[lCompatibilidadTVS] [bit] NULL ,
	[nLongitudBarra] [int] NULL ,
	[lPagoRapido] [bit] NULL,
	[lDisgrega] [bit] NULL,
	[lPasswordPorCobrar] [bit] NULL ,
	[lModificaTipoPedido] [bit] NULL,
	[tSucursal] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NULL ,
	[nBalanzaPuerto] [int] NULL ,
	[lCapturaPeso] [bit] NULL ,
	[lPagoRapidoPV] [bit] NULL ,
	[tTextoConsumo] [nvarchar](200) COLLATE Modern_Spanish_CI_AS NULL ,
	[lSiab] [bit] NULL,
	[tSectorVenta] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[lCajaMobile] [bit] NULL ,
	[lBloqueaPrecuenta] [bit] NULL,
	[lRotulado] [bit] NULL ,
	[lMultiAreaSubGrupo] [bit] NULL ,
	[lMultiAreaCaja] [bit] NULL,
	[lHuella] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MENVIO] (
	[fInicio] [smalldatetime] NOT NULL ,
	[fFinal] [smalldatetime] NOT NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[nMonto] [float] NULL ,
	[lCopia] [bit] NULL ,
	[lCierre] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TPRODUCTOAREA] (
	[tCodigoProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TOFERTA] (
	[tOferta] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tNombre] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tResumido] [nvarchar] (24) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tFrecuencia] [nvarchar] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[fFecha] [smalldatetime] NULL ,
	[tHoraInicial] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[tHoraFinal] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,	
	[lAcumulable] [bit] NULL ,
	[nRatio] [float] NULL ,
	[nMonto] [float] NULL ,
	[nPrecio] [float] NULL ,
	[lPermanente] [bit] NULL ,
	[fFechaInicial] [smalldatetime] NULL ,
	[fFechaFinal] [smalldatetime] NULL ,
	[lLocal] [bit] NULL ,
	[lDelivery] [bit] NULL ,
	[lLlevar] [bit] NULL ,
	[lCanal4] [bit] NULL ,
	[lCanal5] [bit] NULL ,
	[lExcluyente] [bit] NULL ,
	[lAutomatica] [bit] NULL ,
	[lActivo] [bit] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MPROPINA] (
	[tCodigopedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tMoneda] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[nMonto] [float] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tComanda] [nvarchar] (12) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TGRUPOUSUARIO] (
	[tGrupoUsuario] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[lModulo01] [bit] NULL ,
	[lModulo02] [bit] NULL ,
	[lModulo03] [bit] NULL ,
	[lOpcion01] [bit] NULL ,
	[lOpcion02] [bit] NULL ,
	[lOpcion03] [bit] NULL ,
	[lOpcion04] [bit] NULL ,
	[lOpcion05] [bit] NULL ,
	[lOpcion06] [bit] NULL ,
	[lOpcion07] [bit] NULL ,
	[lOpcion08] [bit] NULL ,
	[lOpcion09] [bit] NULL ,
	[lOpcion10] [bit] NULL ,
	[lOpcion11] [bit] NULL ,
	[lOpcion12] [bit] NULL ,
	[lOpcion13] [bit] NULL ,
	[lOpcion14] [bit] NULL ,
	[lOpcion15] [bit] NULL ,
	[lOpcion16] [bit] NULL ,
	[lOpcion17] [bit] NULL ,
	[lOpcion18] [bit] NULL ,
	[lActivo] [bit] NULL ,
	[lOpcion19] [bit] NULL ,
	[lOpcion20] [bit] NULL ,
	[lOpcion21] [bit] NULL ,
	[lModulo04] [bit] NULL ,
	[lModulo05] [bit] NULL ,
	[tNivel] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL,
	[lControlNivel] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TMOTIVODESCUENTO] (
	[tDescuento] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tResumido] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nRatio] [float] NOT NULL ,
	[nTope] [float] NOT NULL ,
	[lRatio] [bit] NULL ,
	[lActivo] [bit] NULL ,
	[lReplica] [bit] NULL ,
	[lTopePedido] [bit] NULL ,
	[lBloqueo] [bit] NULL ,
	[lAplicablePedido] [bit] NULL 
	
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TPRODUCTOPROPIEDAD] (
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItem] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoPropiedad] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tEnlace] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[nInsumo] [float] NULL ,
	[nGasto] [float] NULL ,
	[nManoObra] [float] NULL ,
	[lReplica] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TCOMBOPROPIEDAD] (
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItem] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tItemCombo] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoPropiedad] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL , 
	[tEnlace] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[nInsumo] [float] NULL ,
	[nGasto] [float] NULL ,
	[nManoObra] [float] NULL ,
	[lReplica] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TMENSAJE] (
	[tF1] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF2] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF3] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF4] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF5] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF6] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF7] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF8] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF9] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF10] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF11] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tF12] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DPAGOTARJETA] (
	[tDocumento] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tReferencia] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tTarjeta] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNumero] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFechaVencimiento] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[nMonto] [float] NULL ,
	[nPropina] [float] NULL ,
	[tEstadoDocumento] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TLOG] (
	[tCorrelativo] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nProducto] [int] NULL ,
	[nPropiedad] [int] NULL ,
	[nOferta] [int] NULL ,
	[nMesa] [int] NULL ,
	[nMozo] [int] NULL ,
	[nOtro] [int] NULL ,
	[nCliente] [int] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TOPERADOR] (
	[tOperador] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[lValor] [bit] NULL ,
	[lStockMas] [bit] NULL ,
	[lStockMenos] [bit] NULL ,
	[lObligaPropiedad] [bit] NULL ,
	[nControl] [int] NULL ,
	[lImprime] [bit] NULL ,
	[nBoton] [int] NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[lActivo] [bit] NULL ,
	[lReplica] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TACCESO] (
	[tCodigoAcceso] [nvarchar] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tModulo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDescripcion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFormulario] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTipoObjeto] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNombreObjeto] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTabla] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[nOrden] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TGRUPOACCESO] (
	[tGrupoUsuario] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoAcceso] [nvarchar] (8) COLLATE Modern_Spanish_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TPEDIDO] (
	[nCorrelativo] [bigint] NOT NULL ,	
	[tPedidoIni] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tItemIni] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tPedidoFin] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[tItemFin] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NULL ,
	[nCantidad] [float] NULL ,
	[nImpuesto1] [float] NULL ,
	[nImpuesto2] [float] NULL ,
	[nImpuesto3] [float] NULL ,
	[nVenta] [float] NULL ,
	[tTurno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[fDiaContable] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TPRODUCTOXPRODUCTO] (
	[tCodigoProducto] [char] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tSubProducto] [char] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nCantidad] [float] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TPEDIDOMESA] (
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tMesa] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TLOCAL] (
	[tCodigoLocal] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDetallado] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tResumido] [nvarchar] (24) COLLATE Modern_Spanish_CI_AS NULL ,
	[tcodigoSector] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tIP] [nvarchar] (40) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBaseDatosINF] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBaseDatosALM] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[nMeta] [float] NULL ,
	[lActivo] [bit] NULL ,
	[lReplica] [bit] NULL ,
	[ultConLocal]	[bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TACCESOENVIA] (
	[tCodigoAcceso] [nvarchar] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[lEnvia] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DPEDIDOKDS](
	[tCodigoPedido] [nvarchar](10) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tItem] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[fSalida] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TORIGENCODIGOCONTROL] (
	[tCaja] [nvarchar](3) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[nCorrelativo] [bigint] NOT NULL,
	[fInicio] [datetime] NOT NULL,
	[fFin] [datetime] NOT NULL,
	[tAutorizacion] [nvarchar](25) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tDosificacion] [nvarchar](100) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tSFC] [nvarchar](10) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[lActivo] [bit] NOT NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL 

) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TMENSAJECOCINA](
	[Codigo] [varchar](8) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tUsuarioReg] [varchar](15) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[Mensaje] [varchar](100) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[fregistro] [datetime] NOT NULL,
	[fFinal] [datetime] NOT NULL,
	[tUsuarioFinal] [varchar](15) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tCaja] [varchar](3) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[lActivo] [bit] NOT NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TINSUMO](
	[tcodigo] [varchar](8) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[Descripcion] [varchar](50) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[nStock] [float] NULL,
	[tUsuarioReg] [varchar](15) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[fregistro] [datetime] NOT NULL,
	[tcajaReg] [varchar](3) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tUsuarioModificacion] [varchar](15) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[fmodificacion] [datetime] NOT NULL,
	[tcajaModificacion] [varchar](3) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[lactivo] [bit] NOT NULL,
	[liNSUMO] [bit] NOT NULL
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TDIACONTABLE](
	[fDiaContable] [smalldatetime] NOT NULL,
	[lApertura] [bit] NOT NULL,
	[lCierre] [bit] NOT NULL,
	[tUsuario] [varchar](15) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[fregistro] [datetime] NOT NULL,
	[tUsuarioCierre] [varchar](15) COLLATE Modern_Spanish_CI_AS NULL,
	[fregistroCierre] [datetime] NULL
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TAREAPANTALLA] (
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nColumna] [int] NULL ,
	[lMuestra] [bit] NULL ,
	[nOrden] [int] NULL ,
	[nAncho] [float] NULL ,
	[tUsuario] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tDescripcionMostrar] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDescripcionInterna] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL 
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TAREAPANTALLA1] (
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nColumna] [int] NULL ,
	[lMuestra] [bit] NULL ,
	[nOrden] [int] NULL ,
	[nAncho] [float] NULL ,
	[tUsuario] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tDescripcionMostrar] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDescripcionInterna] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL 
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TAREAPANTALLADESPACHO] (
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nColumna] [int] NULL ,
	[lMuestra] [bit] NULL ,
	[nOrden] [int] NULL ,
	[nAncho] [float] NULL ,
	[tUsuario] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL ,
	[tDescripcionMostrar] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDescripcionInterna] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL 
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TTIPOPEDIDODETALLE] (
	[tcodigoTipoPedido] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[lActivaMozo] [bit] NULL ,
	[lActivaMotorizado] [bit] NULL ,
	[lObligaMesa] [bit] NULL ,
	[lObligaPax] [bit] NULL ,
	[lObligaMozo] [bit] NULL ,
	[lObligaMotorizado] [bit] NULL ,
	[lCanalCentralPedidos] [bit] NULL,
	[lCanalDelivery] [bit] NULL,
	[lObligaIngresoFechaEntrega] [bit] NULL ,
	[lObligaClienteFrecuente] [bit] NULL
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TCLIENTEPRODUCTO] (
	[tcodigoDelivery] [nvarchar] (12) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoProducto] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nPrecio] [float] NULL ,
	[lPermiteDescuentos] [bit] NULL	,
	[tUsuario] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL 
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TCANALVENTA](
	[tCodigoCanalVenta] [nvarchar](2) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tDetallado] [nvarchar](50) COLLATE Modern_Spanish_CI_AS NULL,
	[lActivaMozo] [bit] NULL,
	[lActivaMotorizado] [bit] NULL,
	[lObligaMesa] [bit] NULL,
	[lObligaPax] [bit] NULL,
	[lObligaMozo] [bit] NULL,
	[lObligaMotorizado] [bit] NULL,
	[lCanalCentralPedidos] [bit] NULL,
	[lCanalDelivery] [bit] NULL,
	[lObligaIngresoFechaEntrega] [bit] NULL,
	[lObligaClienteFrecuente] [bit] NULL,
	[lActivo] [bit] NULL
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TMOTORIZADODATOS] (
	[tCodigo] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDocumentoIdentidad] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nTarifaLV] [float] NULL ,
	[nTarifaSD] [float] NULL ,
	[nTarifaES] [float] NULL 
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TIMPORTACION] (
	[tCodigo] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nOrden] [int] NOT NULL  ,
	[tPadre] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS  NULL ,
	[tAgrupacion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS  NULL ,
	[tTablaMostrar] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS  NULL ,
	[tTablaInterna] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS  NULL ,
	[tCampoMostrar] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS  NULL ,
	[tCampoInterno] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS  NULL ,
	[tAgrupacionTabla] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS  NULL ,
	[lOculto] [bit] NULL ,
	[lImportar] [bit] NULL ,
	[fRegistro] [datetime] NULL ,
	[tUsuario] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TIMPORTACIONLOG] (
	[tModulo] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS  NULL ,
	[fRegistro] [datetime] NULL ,
	[tUsuario] [nvarchar] (55) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEstado] [nvarchar] (55) COLLATE Modern_Spanish_CI_AS NULL ,
	[tObservacion] [nvarchar] (4000) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TTRAMITE] (
	[tCodigoTramite] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tDescripcion] [nvarchar] (200) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[lSolicitaNAnteriorAutorizacion] [bit] NULL ,
	[lActivo] [bit] NULL  
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TSOLICITUD] (
	[tCodigoSolicitud] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoTramite] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tEstadoSolicitud] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNumeroAutorizacion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tNumeroAutorizacionAnterior] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[fFechaEmision] [datetime] NULL ,
	[fFechaCaducidad] [datetime] NULL ,
	[fRegistro] [datetime] NULL ,
	[tUsuario] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistroModifica] [datetime] NULL ,
	[tUsuarioModifica] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TSOLICITUDDETALLE] (
	[tCodigoSolicitud] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nCorrelativo] [int] NOT NULL ,
	[tTipoDocumento] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tSerieEstablecimiento] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tSerieCaja] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCaja] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tFolioInicial] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tFolioFinal] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEstadoDetalle] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [datetime] NULL ,
	[tUsuario] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistroModifica] [datetime] NULL ,
	[tUsuarioModifica] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TDELIVERYINVITADO] (
	[tCodigoInvitado] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoDelivery] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tNombre] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tApellido] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fRegistro] [datetime] NULL ,
	[tUsuario] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TPARIENTE] (
	[tCodigoPariente] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoDelivery] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tNombre] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tApellido] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[lConyugue] [bit] NULL,
	[lHijo] [bit] NULL,
	[fRegistro] [datetime] NULL ,
	[tUsuario] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NULL  ,
	[iFoto] [image] NULL
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[TDELIVERYCLIENTE] (
	[nCorrelativo] [bigint] NOT NULL ,
	[tCodigoDelivery] [nvarchar] (7) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tCodigoCliente] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL  
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TAREASUBGRUPO] (
	[tCaja] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tSubGrupo] [nvarchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tArea] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tUsuario] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[fRegistro] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TAREACHEF](
	[tCaja] [nvarchar](3) NOT NULL,
	[tArea] [nvarchar](3) NOT NULL,
	[lArea] [bit] NOT NULL,
	[tUsuario] [nvarchar](15) NULL,
	[fRegistro] [smalldatetime] NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TMENSAJEUSUARIO](
	[tCodigoPedido] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tUsuario] [nvarchar](15) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tMensaje] [nvarchar](4000) COLLATE Modern_Spanish_CI_AS NOT NULL,	
	[tCaja] [varchar](3) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tMozo] [nvarchar](200) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tMesa] [nvarchar](200) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[tTipoPedido] [nvarchar](200) COLLATE Modern_Spanish_CI_AS NOT NULL,
	[fregistro] [datetime] NOT NULL
) ON [PRIMARY]
GO