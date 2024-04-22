if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vArea]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vArea]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vCortesia]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vCortesia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadoDocumento]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadoDocumento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadoMesa]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadoMesa]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadoPedido]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadoPedido]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadoGuia]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadoGuia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDistrito]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDistrito]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadoReserva]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadoReserva]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vLocal]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vLocal]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMoneda]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vMoneda]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMotorizado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vMotorizado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMozo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vMozo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSalon]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vSalon]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoAtencion]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoAtencion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoDescargo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoDescargo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoDocumento]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoDocumento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoIdentidad]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoIdentidad]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoPago]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoPago]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoPedido]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoPedido]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoProducto]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoProducto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipodocumentoImpresora]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipodocumentoImpresora]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vZona]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vZona]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAreaImpresora]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vAreaImpresora]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPropiedad]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPropiedad]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vFrecuencia]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vFrecuencia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoClienteFrecuente]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoClienteFrecuente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vCliente]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vCompania]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vCompania]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vCtaCte]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vCtaCte]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDelivery]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDelivery]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vFormulario]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vFormulario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vGrupo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vGrupo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vGrupoUsuario]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vGrupoUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubGrupo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vSubGrupo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vProducto]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vProducto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vProductoArea]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vProductoArea]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vCombo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vCombo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vComboDetalle]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vComboDetalle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEgreso]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEgreso]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vFacturacionDetalle]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vFacturacionDetalle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vIngreso]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vIngreso]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vLiquidacion]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vLiquidacion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vNotaCredito]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vNotaCredito]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPreCuenta]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPreCuenta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPreCuentaDelivery]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPreCuentaDelivery]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPreCuentaDetallada]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPreCuentaDetallada]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedido]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedido]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoAgrupado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoAgrupado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoCabecera]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoCabecera]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoCombo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoCombo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoResultado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoResultado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoCorrelativo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoCorrelativo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoDetalle]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoDetalle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoGrilla]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoGrilla]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumento]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoImpresora]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoImpresora]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoAgrupado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoAgrupado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoConsolidado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoConsolidado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoGrilla]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoGrilla]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoPago]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoPago]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoResultado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoResultado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vOperador]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vOperador]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vOferta]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vOferta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMotivoEliminacion]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vMotivoEliminacion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMotivoDescuento]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vMotivoDescuento]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoCancelacion]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoCancelacion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMotivoTraslado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vMotivoTraslado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChofer]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vChofer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vVehiculo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vVehiculo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vGuiaTransporte]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vGuiaTransporte]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vGuia]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vGuia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPrecuentaAgrupada]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPrecuentaAgrupada]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoCliente]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoImpresoraAgrupado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoImpresoraAgrupado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEmpacador]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEmpacador]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDespachador]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDespachador]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubTipoCtaCte]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vSubTipoCtaCte]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTipoCtaCte]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTipoCtaCte]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPRODUCTOXPRODUCTO]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vProductoXProducto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vUnidadNegocio]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vUnidadNegocio]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[returnAnoMes]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[returnAnoMes]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTablasCentralizada]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTablasCentralizada]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[VSector]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[VSector]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPaloteoProduccionPropiedades]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPaloteoProduccionPropiedades]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPaloteoProduccionPropiedadesCombos]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPaloteoProduccionPropiedadesCombos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoCorrelativoDetalle]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoCorrelativoDetalle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TraePropiedad]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[TraePropiedad]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoImpresoraAgrupadoAlternativa]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoImpresoraAgrupadoAlternativa]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoImpresoraAlternativa]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoImpresoraAlternativa]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPaisOrigen]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPaisOrigen]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vCajaCodigoControl]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vCajaCodigoControl]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSucursal]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vSucursal]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMaitre]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vMaitre]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadoFrecuente]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadoFrecuente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTarjetaCredito]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTarjetaCredito]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTienda]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTienda]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[VTIPOEGRESO]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[VTIPOEGRESO]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadoSolicitud]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadoSolicitud]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadoSolicitudDetalle]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadoSolicitudDetalle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vInvitado]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vInvitado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPariente]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPariente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSectorVenta]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vSectorVenta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSectorVentaCajaR]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vSectorVentaCajaR]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAreaSubGrupo]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vAreaSubGrupo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAreaChef]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vAreaChef]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vInvitado
AS

select tcodigoinvitado as codigo, tcodigodelivery as codigoDelivery,
tNombre, tApellido,Tnombre + ' ' + tapellido as Invitado
from tdeliveryinvitado
		
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.VTIPOEGRESO
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = N'TIPOEGRESO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION TraePropiedad (@Pedido NVARCHAR(20), @Item NVARCHAR(6))
RETURNS NVARCHAR(400)
AS
BEGIN
	DECLARE @Cant	FLOAT
	DECLARE @Flag	CHAR(1)
	DECLARE @Cadena		NVARCHAR(400)
	DECLARE @Resumido	VARCHAR(50)

	SET @Cadena = ''
    DECLARE CURSORITO CURSOR FOR
	SELECT dbo.TOPERADOR.tResumido + ' ' + dbo.TPROPIEDAD.tResumido FROM dbo.TPRODUCTOPROPIEDAD
	INNER JOIN dbo.TPROPIEDAD ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto 
	INNER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador
	WHERE tCodigoPedido = @Pedido AND TOPERADOR.lImprime=1 AND tItem = @Item
	ORDER BY dbo.TOPERADOR.tResumido + ' ' + dbo.TPROPIEDAD.tResumido
    OPEN CURSORITO 
	FETCH NEXT FROM CURSORITO INTO @Resumido
	WHILE @@FETCH_STATUS = 0
    BEGIN
		SET @Cadena = @Cadena + @Resumido + ','
	    FETCH NEXT FROM CURSORITO INTO @Resumido
    END
    CLOSE CURSORITO
	DEALLOCATE CURSORITO

    RETURN(@Cadena)
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vOperador
AS
SELECT     tOperador AS Codigo, tDetallado AS Descripcion, tResumido, nBoton, lValor, lStockMas, lStockMenos, lObligaPropiedad, nControl, lImprime, lActivo, tUsuario, fRegistro
FROM       dbo.TOPERADOR

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPropiedad
AS
SELECT DISTINCT 
                      dbo.TPROPIEDAD.tCodigoPropiedad AS Codigo, dbo.TPROPIEDAD.tDetallado AS Descripcion, 
                      dbo.vOperador.Descripcion AS Operador, dbo.TPROPIEDAD.tResumido, dbo.TPROPIEDAD.lActivo, dbo.TPROPIEDAD.tOperador, dbo.vOperador.lValor, 
                      dbo.vOperador.lStockMas, dbo.vOperador.lStockMenos, dbo.vOperador.nControl, dbo.TPROPIEDAD.nPrecio, dbo.TPROPIEDAD.tEnlace, dbo.TPROPIEDAD.tArea
FROM  		      dbo.TPROPIEDAD LEFT OUTER JOIN
                      dbo.vOperador ON dbo.TPROPIEDAD.tOperador = dbo.vOperador.Codigo


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vPaloteoProduccionPropiedades]
AS
SELECT     dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tProducto, 
                      dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.vPropiedad.Operador + ' ' + dbo.vPropiedad.tResumido AS Propiedad, 
                      dbo.TPRODUCTOPROPIEDAD.nInsumo, dbo.TPRODUCTOPROPIEDAD.nGasto, dbo.TPRODUCTOPROPIEDAD.nManoObra, 
                      dbo.TPRODUCTOPROPIEDAD.nInsumo + dbo.TPRODUCTOPROPIEDAD.nGasto + dbo.TPRODUCTOPROPIEDAD.nManoObra AS Costo, 
                      dbo.vPropiedad.lStockMas, dbo.vPropiedad.lStockMenos
FROM         dbo.TPRODUCTOPROPIEDAD INNER JOIN
                      dbo.vPropiedad ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.vPropiedad.Codigo INNER JOIN
                      dbo.MPEDIDO ON dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido
WHERE     (dbo.vPropiedad.lStockMenos = 1) OR
                      (dbo.vPropiedad.lStockMas = 1)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
 

CREATE VIEW [dbo].[vPaloteoProduccionPropiedadesCombos]
AS
SELECT     dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, dbo.TCOMBOPROPIEDAD.tProducto, 
                      dbo.TCOMBOPROPIEDAD.tCodigoPropiedad, dbo.vPropiedad.Operador + ' ' + dbo.vPropiedad.tResumido AS Propiedad, 
                      dbo.TCOMBOPROPIEDAD.nInsumo, dbo.TCOMBOPROPIEDAD.nGasto, dbo.TCOMBOPROPIEDAD.nManoObra, 
                      dbo.TCOMBOPROPIEDAD.nInsumo + dbo.TCOMBOPROPIEDAD.nGasto + dbo.TCOMBOPROPIEDAD.nManoObra AS Costo, dbo.vPropiedad.lStockMas, 
                      dbo.vPropiedad.lStockMenos
FROM         dbo.TCOMBOPROPIEDAD INNER JOIN
                      dbo.vPropiedad ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.vPropiedad.Codigo INNER JOIN
                      dbo.MPEDIDO ON dbo.TCOMBOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido
WHERE     (dbo.vPropiedad.lStockMenos = 1) OR
                      (dbo.vPropiedad.lStockMas = 1)

GO
 
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

 
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.VSector
AS
SELECT     TCODIGO AS CODIGO, tDetallado AS DESCRIPCION, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = N'SECTOR')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.vFrecuencia
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, nBoton, tValor, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'FRECUENCIA')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.vTablasCentralizada
AS 
SELECT TACCESO.tCodigoAcceso, TACCESO.tDescripcion, 
TACCESO.tModulo, 
ISNULL(TACCESOENVIA.lEnvia, 0) AS lenvia, 
TACCESO.tFormulario , 
TACCESO.tNombreObjeto, TACCESO.tTipoObjeto , isnull(tacceso.ttabla,'0') as ttabla
FROM     TACCESO LEFT OUTER JOIN TACCESOENVIA ON TACCESO.tCodigoAcceso = TACCESOENVIA.tCodigoAcceso 
WHERE  (TACCESO.tTipoObjeto = 'MN') AND (TACCESO.tModulo = '03') AND (SUBSTRING(TACCESO.tCodigoAcceso, 1, 3) <> '104') AND (SUBSTRING(TACCESO.tCodigoAcceso, 1, 3) <> '105')
 and (isnull(tacceso.ttabla,'0')<>'0' )
union all
SELECT TACCESO.tCodigoAcceso, TACCESO.tDescripcion, 
TACCESO.tModulo, 
ISNULL(TACCESOENVIA.lEnvia, 0) AS lenvia, 
TACCESO.tFormulario , 
TACCESO.tNombreObjeto, TACCESO.tTipoObjeto , isnull(tacceso.ttabla,'0') as ttabla
FROM     TACCESO LEFT OUTER JOIN TACCESOENVIA ON TACCESO.tCodigoAcceso = TACCESOENVIA.tCodigoAcceso 
WHERE  (TACCESO.tTipoObjeto = 'MN') AND (TACCESO.tModulo = '03') AND (SUBSTRING(TACCESO.tCodigoAcceso, 1, 3) <> '104') and (SUBSTRING(TACCESO.tCodigoAcceso, 1, 3) <> '105') and substring(tacceso.tcodigoacceso,4,5)='00000'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE FUNCTION returnAnoMes (@now datetime)
RETURNS nvarchar(6)
AS
 BEGIN
  DECLARE @returnDate nvarchar(6)
  SET @returnDate=RIGHT('0'+LTRIM(MONTH(@now)),2)+LTRIM((YEAR(@now)))
  RETURN(@returnDate)
 END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTipoCliente
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, tValor, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'TIPOCLIENTE')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vEstadoFrecuente]
AS

SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido,nboton,nValor ,lActivo, TTABLA  
FROM         dbo.TTABLA  
WHERE     (TTABLA = 'ESTADOFRECUENTE') 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vArea
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido, tValor, lActivo, tIcono,     nValor,nBoton as 'KDS' ,isnull(ntamano,0) as lCheffControl 
FROM         dbo.TTABLA
WHERE     (TTABLA = N'AREA')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vSubTipoCtaCte
AS
SELECT     SUBSTRING(TCODIGO, 1, 4) AS Codigo, tDetallado AS Descripcion, tResumido, tValor AS tTipoCtaCte, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'SUBTIPOCTACTE')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTipoCtaCte
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'TIPOCTACTE')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vCortesia
AS
SELECT     SUBSTRING(TCODIGO, 1, 4) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, isnull(nvalor,0) as  tope
FROM         dbo.TTABLA
WHERE     (TTABLA = 'CORTESIA')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vEstadoDocumento
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'ESTADODOCUMENTO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vEstadoMesa
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'ESTADOMESA')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vEstadoPedido
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, TTABLA
FROM         dbo.TTABLA
WHERE     (TTABLA = 'ESTADOPEDIDO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vEstadoGuia
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, TTABLA
FROM         dbo.TTABLA
WHERE     (TTABLA = 'ESTADOGUIA')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vEstadoReserva
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'ESTADORESERVA')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vLocal
AS
SELECT      dbo.TLOCAL.tCodigoLocal AS Codigo, ISNULL(dbo.VSector.tResumido, '') AS Sector, dbo.TLOCAL.tDetallado AS Descripcion, 
                      dbo.TLOCAL.tResumido, dbo.TLOCAL.tIP AS IP, dbo.TLOCAL.tBaseDatosINF AS BDINF, dbo.TLOCAL.tBaseDatosALM AS BDALM, 
                      dbo.TLOCAL.nMeta AS Meta, dbo.TLOCAL.lActivo, dbo.TLOCAL.ultConLocal, ISNULL(dbo.VSector.CODIGO, '') AS tcodigosector
FROM         dbo.TLOCAL LEFT OUTER JOIN
                      dbo.VSector ON dbo.TLOCAL.tCodigoSector = dbo.VSector.CODIGO

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vMoneda
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, TTABLA
FROM         dbo.TTABLA
WHERE     (TTABLA = 'MONEDA') AND (tResumido <> '')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vMotorizado
AS
SELECT     SUBSTRING(TCODIGO, 1, 4) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, nBoton, nValor, TTABLA
FROM         dbo.TTABLA
WHERE     (TTABLA = 'MOTORIZADO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vMozo
AS
SELECT     SUBSTRING(TCODIGO, 1, 4) AS Codigo, tDetallado AS Descripcion, tResumido, nBoton, tValor, lActivo, nValor, tIcono AS tBandaMagnetica, nTamano, tValor2 As tHuella 
FROM         dbo.TTABLA
WHERE     (TTABLA = 'MOZO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vSalon
AS
SELECT     SUBSTRING(dbo.TTABLA.TCODIGO, 1, 2) AS Codigo, dbo.TTABLA.tDetallado AS Descripcion, dbo.TTABLA.tResumido, dbo.TTABLA.lActivo, 
                      dbo.TTABLA.tValor AS tLocal, dbo.TTABLA.nValor, dbo.vLocal.Descripcion AS [Local], dbo.TTABLA.tIcono
FROM         dbo.TTABLA LEFT OUTER JOIN
                      dbo.vLocal ON dbo.TTABLA.tValor = dbo.vLocal.Codigo
WHERE     (dbo.TTABLA.TTABLA = 'SALON')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTipoAtencion
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = N'TIPOATENCION')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTipoDescargo
AS
SELECT     SUBSTRING(TCODIGO, 1, 1) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'TIPODESCARGO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vTarjetaCredito]
AS
SELECT     tCodigoTarjeta AS Codigo, tDetallado AS Descripcion
FROM         dbo.TTARJETACREDITO
WHERE     (lActivo = 1)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vTipoDocumento]
AS
SELECT     tCodigoTipoDocumento AS Codigo, tDescripcion AS Descripcion, tPrefijo AS Prefijo, lPideCliente AS Cliente, lActivo, tCodigoSunat AS Sunat, 
           lRegistroVenta AS RegistroVenta, nMonto AS Monto, lTransporte AS Transporte, lCanjearNotaCredito as Canjear, lValidaRuc
FROM         dbo.TTIPODOCUMENTO

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTipoIdentidad
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, nValor
FROM         dbo.TTABLA
WHERE     (TTABLA = 'TIPOIDENTIDAD')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW dbo.vTipoPago
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, tValor AS CTAMN, tIcono AS CTAME
FROM         dbo.TTABLA
WHERE     (TTABLA = N'TIPOPAGO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTipoPedido
AS
SELECT     tCodigoCanalVenta AS Codigo, tDetallado AS Descripcion, lActivaMozo, lActivaMotorizado, lObligaMesa, lObligaPax, lObligaMozo, lObligaMotorizado, 
                      lCanalCentralPedidos, lCanalDelivery, lObligaIngresoFechaEntrega, lObligaClienteFrecuente, lActivo
FROM         dbo.TCANALVENTA
WHERE     (lActivo = 1)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vUnidadNegocio
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, tValor, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'UNIDADNEGOCIO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDistrito
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido, tValor, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'DISTRITO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vTipoClienteFrecuente]
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, TTABLA, nValor, tValor
FROM         dbo.TTABLA
WHERE     (TTABLA = 'TIPOCLIENTEFRECUENTE')
GO

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vMotivoDescuento
AS
SELECT     tDescuento AS Codigo, tDetallado AS Descripcion, tResumido, nRatio, lRatio, nTope, lTopePedido, lBloqueo, lActivo, lAplicablePedido
FROM         dbo.TMOTIVODESCUENTO

GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTipoProducto
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, tValor, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'TIPOPRODUCTO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW dbo.vTipodocumentoImpresora
AS
SELECT     dbo.TTIPODOCUMENTOIMPRESORA.tCaja, dbo.TTIPODOCUMENTOIMPRESORA.tTipoEmision, dbo.TTIPODOCUMENTOIMPRESORA.tImpresora, 
                      dbo.TTIPODOCUMENTOIMPRESORA.tDescripcion, dbo.TTIPODOCUMENTOIMPRESORA.tFormulario, dbo.TTIPODOCUMENTOIMPRESORA.tSerie, 
                      dbo.TTIPODOCUMENTOIMPRESORA.tUltimoNumero, dbo.TTIPODOCUMENTOIMPRESORA.tUsuario, dbo.TTIPODOCUMENTOIMPRESORA.fRegistro, 
                      dbo.vTipoDocumento.Descripcion, dbo.vTipoDocumento.Prefijo, dbo.vTipoDocumento.Cliente, dbo.TIMPRESORA.tRuta, 
                      dbo.TTIPODOCUMENTOIMPRESORA.lResumen, dbo.vTipoDocumento.Monto, dbo.TTIPODOCUMENTOIMPRESORA.lImpuesto1, 
                      dbo.TTIPODOCUMENTOIMPRESORA.lImpuesto2, dbo.TTIPODOCUMENTOIMPRESORA.lImpuesto3, dbo.vTipoDocumento.RegistroVenta, 
                      dbo.TIMPRESORA.tDescripcion AS Impresora, dbo.TTIPODOCUMENTOIMPRESORA.lEquivaDolares, dbo.vTipoDocumento.Transporte, 
                      dbo.TTIPODOCUMENTOIMPRESORA.tNumeroAutorizacion, dbo.TTIPODOCUMENTOIMPRESORA.fInicio, dbo.TTIPODOCUMENTOIMPRESORA.fCaducidad
FROM         dbo.TIMPRESORA RIGHT OUTER JOIN
                      dbo.TTIPODOCUMENTOIMPRESORA ON dbo.TIMPRESORA.tImpresora = dbo.TTIPODOCUMENTOIMPRESORA.tImpresora AND 
                      dbo.TIMPRESORA.tCaja = dbo.TTIPODOCUMENTOIMPRESORA.tCaja LEFT OUTER JOIN
                      dbo.vTipoDocumento ON dbo.TTIPODOCUMENTOIMPRESORA.tTipoEmision = dbo.vTipoDocumento.Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW dbo.vZona
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido, nValor, lActivo, TTABLA
FROM         dbo.TTABLA
WHERE     (TTABLA = 'ZONA')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vAreaImpresora
AS
SELECT     dbo.TAREAIMPRESORA.tCaja, dbo.TAREAIMPRESORA.tArea, dbo.TAREAIMPRESORA.tImpresora, dbo.TIMPRESORA.tRuta, dbo.TIMPRESORA.tFont, 
                      dbo.vArea.Descripcion AS Area, dbo.TIMPRESORA.tDescripcion AS Impresora, dbo.vArea.tIcono, dbo.vArea.nValor
FROM         dbo.TAREAIMPRESORA LEFT OUTER JOIN
                      dbo.TIMPRESORA ON dbo.TAREAIMPRESORA.tImpresora = dbo.TIMPRESORA.tImpresora AND 
                      dbo.TAREAIMPRESORA.tCaja = dbo.TIMPRESORA.tCaja LEFT OUTER JOIN
                      dbo.vArea ON dbo.TAREAIMPRESORA.tArea = dbo.vArea.Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vCliente
AS
SELECT     tCodigoCliente AS Codigo, tEmpresa AS Descripcion, tIdentidad, tDireccion, tUsuario, fRegistro, lActivo, tCorreo
FROM         dbo.TCLIENTE

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW [dbo].[vCompania]
AS
SELECT     dbo.TDELIVERY.tCodigoDelivery AS codigo, dbo.TDELIVERY.tApellido + ' ' + dbo.TDELIVERY.tNombre AS tnomSoc, 
                      dbo.TDELIVERY.tApellido + ' ' + dbo.TDELIVERY.tNombre AS Descripcion, dbo.TDELIVERY.tDireccion, dbo.TDELIVERY.tTelefono, dbo.TDELIVERY.lActivo, 
                      dbo.TDELIVERY.tEMail, ISNULL(dbo.TDELIVERY.nConsumo, 0) AS nConsumo, ISNULL(dbo.TDELIVERY.nLinea, 0) AS nLinea, ISNULL(dbo.TDELIVERY.nLinea, 0) 
                      - ISNULL(dbo.TDELIVERY.nConsumo, 0) AS nSaldo, dbo.TDELIVERY.fRegistro, dbo.TDELIVERY.tUsuario, dbo.TDELIVERY.tTipoCtaCte, 
                      dbo.TDELIVERY.tSubTipoCtaCte, ISNULL(dbo.TDELIVERY.lClienteCtaCte, 0) AS lclientrectacte, ISNULL(dbo.TDELIVERY.tCodigoCliente, N'') AS TCODIGOCLIENTE, 
                      dbo.TDELIVERY.tIdentidad AS Identidad, dbo.vEstadoFrecuente.Descripcion AS EstadoFrecuente, ISNULL(dbo.TDELIVERY.nLineaPorCobrar, 0) AS nLineaPorCobrar, 
                      ISNULL(dbo.TDELIVERY.nConsumoPorCobrar, 0) AS nConsumoPorCobrar, ISNULL(dbo.vTipoIdentidad.tResumido, N'') + '  - ' + ISNULL(dbo.TDELIVERY.tIdentidad, N'') 
                      AS Identi
FROM         dbo.TDELIVERY LEFT OUTER JOIN
                      dbo.vTipoIdentidad ON dbo.TDELIVERY.tTipoIdentidad = dbo.vTipoIdentidad.Codigo LEFT OUTER JOIN
                      dbo.vEstadoFrecuente ON dbo.TDELIVERY.tEstadoFrecuente = dbo.vEstadoFrecuente.Codigo
WHERE     (dbo.TDELIVERY.lActivo = 1) AND (dbo.TDELIVERY.lClienteCtaCte = 1)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vCtaCte
AS
SELECT     dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tCaja, 
                      dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tResumido AS Producto, dbo.DPEDIDO.nPrecioOficial, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.nPrecioNeto, 
                      dbo.DPEDIDO.nImpuesto1, dbo.DPEDIDO.nImpuesto2, dbo.DPEDIDO.nImpuesto3, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tMesa, 
                      dbo.vMozo.Descripcion AS Mozo, dbo.DPEDIDO.tFacturado, dbo.vCompania.tnomSoc AS Cliente, dbo.vCompania.nLinea - dbo.vCompania.nConsumo AS nSaldo, 
                      dbo.vCompania.nLinea, ISNULL(dbo.MPEDIDO.tUsuario, N'') AS tusuario
FROM         dbo.vSalon RIGHT OUTER JOIN
                      dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon RIGHT OUTER JOIN
                      dbo.vMozo RIGHT OUTER JOIN
                      dbo.vCompania RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vCompania.codigo = dbo.MPEDIDO.tClienteCtaCte ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON 
                      dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa RIGHT OUTER JOIN
                      dbo.TPRODUCTO RIGHT OUTER JOIN
                      dbo.DPEDIDO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vDelivery]
AS
SELECT     dbo.TDELIVERY.tCodigoDelivery AS Codigo, dbo.TDELIVERY.tTipoCliente, dbo.TDELIVERY.tApellido, dbo.TDELIVERY.tNombre, 
                      dbo.TDELIVERY.tDireccion, dbo.TDELIVERY.tTelefono, dbo.TDELIVERY.tReferencia, dbo.TDELIVERY.tCodigoCliente, dbo.TDELIVERY.lActivo, 
                      dbo.TDELIVERY.tUsuario, dbo.TDELIVERY.fRegistro, dbo.TDELIVERY.tZona, dbo.TDELIVERY.tObservacion, dbo.TDELIVERY.nDescuento, 
                      dbo.TDELIVERY.tEmail, dbo.TDELIVERY.fNacimiento, dbo.vZona.Descripcion AS Zona, LTRIM(dbo.TDELIVERY.tNombre) 
                      + ' ' + LTRIM(dbo.TDELIVERY.tApellido) AS Cliente, ISNULL(dbo.TDELIVERY.lPuntos, 0) AS lpuntos, dbo.TDELIVERY.nAcumulado, 
                      dbo.TDELIVERY.nUtilizado, dbo.TDELIVERY.nDisponible, dbo.vTipoClienteFrecuente.Descripcion AS TipoCliente, dbo.vZona.tResumido, 
                      dbo.TDELIVERY.tCodigoTarjeta, dbo.TDELIVERY.tNumeroTarjeta, dbo.TDELIVERY.tFechaTarjeta, dbo.TDELIVERY.tDistrito, 
                      dbo.vDistrito.Descripcion AS Distrito, dbo.TDELIVERY.tEstadoFrecuente, dbo.TDELIVERY.lExcluyeProductos, dbo.TDELIVERY.lClienteCtaCte, 
                      ISNULL(dbo.TDELIVERY.nConsumo, 0) AS nConsumo, ISNULL(dbo.TDELIVERY.nLinea, 0) AS nLinea, ISNULL(dbo.TDELIVERY.tTipoCtaCte, N'') 
                      AS tTipoCtaCte, ISNULL(dbo.TDELIVERY.tSubTipoCtaCte, N'') AS tsubtipoctacte, dbo.vTarjetaCredito.Descripcion AS TarjetaCredito, 
                      dbo.VESTADOFRECUENTE.Descripcion AS EstadoFrecuente, dbo.TDELIVERY.tTipoIdentidad, dbo.TDELIVERY.tIdentidad, 
                      dbo.vTipoIdentidad.Descripcion AS TipoIdentidad, dbo.vTipoCtaCte.Descripcion AS TipoCtaCte, 
                      dbo.vSubTipoCtaCte.Descripcion AS SubTipoCtaCte, dbo.tdelivery.iFoto as Foto, isnull(dbo.tdelivery.taccionsocio,'') taccionsocio, isnull(tdelivery.nlineaporcobrar,0) nLineaPorCobrar,  ISNULL(dbo.TDELIVERY.nConsumoPorCobrar, 0) AS nConsumoPorCobrar
FROM         dbo.TDELIVERY LEFT OUTER JOIN
                      dbo.vTarjetaCredito ON dbo.TDELIVERY.tCodigoTarjeta = dbo.vTarjetaCredito.Codigo LEFT OUTER JOIN
                      dbo.vTipoIdentidad ON dbo.TDELIVERY.tTipoIdentidad = dbo.vTipoIdentidad.Codigo LEFT OUTER JOIN
                      dbo.vSubTipoCtaCte ON dbo.TDELIVERY.tSubTipoCtaCte = dbo.vSubTipoCtaCte.Codigo LEFT OUTER JOIN
                      dbo.vTipoCtaCte ON dbo.TDELIVERY.tTipoCtaCte = dbo.vTipoCtaCte.Codigo LEFT OUTER JOIN
                      dbo.VESTADOFRECUENTE ON dbo.TDELIVERY.tEstadoFrecuente = dbo.VESTADOFRECUENTE.Codigo LEFT OUTER JOIN
                      dbo.vDistrito ON dbo.TDELIVERY.tDistrito = dbo.vDistrito.Codigo LEFT OUTER JOIN
                      dbo.vTipoClienteFrecuente ON dbo.TDELIVERY.tTipoCliente = dbo.vTipoClienteFrecuente.Codigo LEFT OUTER JOIN
                      dbo.vZona ON dbo.TDELIVERY.tZona = dbo.vZona.Codigo


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vFormulario
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = N'FORMULARIO')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vGrupo
AS
SELECT     tCodigoGrupo AS Codigo, tDetallado AS Descripcion, tResumido, nBoton, lActivo, tCaja
FROM         dbo.TGRUPO

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vSubGrupo
AS
SELECT     dbo.TSUBGRUPO.tCodigoSubgrupo AS Codigo, dbo.TSUBGRUPO.tDetallado AS Descripcion, dbo.TSUBGRUPO.tResumido, dbo.TSUBGRUPO.lActivo, 
           dbo.TSUBGRUPO.nBoton, dbo.TSUBGRUPO.tIcono, dbo.TSUBGRUPO.tCodigoGrupo AS tGrupo, dbo.TSUBGRUPO.tUsuario, dbo.TSUBGRUPO.fRegistro,
           dbo.vArea.Descripcion AS Area, dbo.TSUBGRUPO.tArea, dbo.TSUBGRUPO.lImprimeArea, dbo.TSUBGRUPO.lImpuesto1, dbo.TSUBGRUPO.lImpuesto2, 
           dbo.TSUBGRUPO.lImpuesto3, dbo.vGrupo.Descripcion AS Grupo, dbo.TSUBGRUPO.nOrden, dbo.TSUBGRUPO.tAgrupacion,dbo.TSUBGRUPO.tcuentacontable
FROM       dbo.TSUBGRUPO LEFT OUTER JOIN
           dbo.vArea ON dbo.TSUBGRUPO.tArea = dbo.vArea.Codigo LEFT OUTER JOIN
           dbo.vGrupo ON dbo.TSUBGRUPO.tCodigoGrupo = dbo.vGrupo.Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vProducto
AS
SELECT     dbo.TPRODUCTO.tCodigoProducto AS Codigo, dbo.vTipoProducto.Descripcion AS TipoProducto, dbo.TPRODUCTO.tDetallado AS Descripcion, 
                      dbo.TPRODUCTO.tResumido, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TSUBGRUPO.tDetallado AS SubGrupo, 
                      dbo.TPRODUCTO.lActivo, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, 
                      dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, 
                      dbo.TPRODUCTO.lImpuesto10, dbo.TPRODUCTO.lImpuesto11, dbo.TPRODUCTO.lImpuesto12, dbo.TPRODUCTO.lImpuesto13, 
                      dbo.TPRODUCTO.lImpuesto14, dbo.TPRODUCTO.lImpuesto15, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.tCortesia, dbo.TPRODUCTO.tArea, 
                      dbo.TPRODUCTO.lImprimeArea, dbo.TGRUPO.tDetallado AS Grupo, dbo.TPRODUCTO.tUsuario, dbo.TPRODUCTO.fRegistro, 
                      dbo.TPRODUCTO.tTipoProducto, dbo.TPRODUCTO.lModificable, dbo.TPRODUCTO.tDescargo, dbo.vArea.Descripcion AS Area, dbo.TPRODUCTO.nBoton, 
                      dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.nPrecioCanal4, dbo.TPRODUCTO.nPrecioCanal5, 
                      dbo.TPRODUCTO.lCombinacion, dbo.TPRODUCTO.nCombinacion, dbo.TPRODUCTO.tEnlace, dbo.TsubGRUPO.tCuentaContable, 
                      dbo.TPRODUCTO.nBotonRapido, dbo.TPRODUCTO.tCajaRapida, dbo.TPRODUCTO.tMoneda, dbo.vMoneda.tResumido AS Moneda, dbo.TPRODUCTO.nInsumo,
                      dbo.TPRODUCTO.nInsumo2, dbo.TPRODUCTO.nInsumo3, dbo.TPRODUCTO.nInsumo4, dbo.TPRODUCTO.nInsumo5, dbo.TPRODUCTO.nGasto, 
                      dbo.TPRODUCTO.nGasto2, dbo.TPRODUCTO.nGasto3, dbo.TPRODUCTO.nGasto4, dbo.TPRODUCTO.nGasto5, dbo.TPRODUCTO.nManoObra, 
                      dbo.TPRODUCTO.nManoObra2, dbo.TPRODUCTO.nManoObra3, dbo.TPRODUCTO.nManoObra4, dbo.TPRODUCTO.nManoObra5, 
                      dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.tBarra, dbo.TPRODUCTO.lPropiedad, dbo.TPRODUCTO.tInfhotel, dbo.TPRODUCTO.lDescuento, 
                      dbo.TPRODUCTO.lLocal, dbo.TPRODUCTO.lDelivery, dbo.TPRODUCTO.lLlevar, dbo.TPRODUCTO.lCanal4, dbo.TPRODUCTO.lCanal5, dbo.TPRODUCTO.tUnidadNegocio, dbo.TPRODUCTO.lMultiArea,dbo.tproducto.talternativa as Talternativa, lControlInsumoCritico,tcodigoInsumo, dbo.TPRODUCTO.nTiempo,ISNULL(DBO.TPRODUCTO.LBALANZA,0) LBALANZA
FROM         dbo.vMoneda INNER JOIN
                      dbo.TPRODUCTO ON dbo.vMoneda.Codigo = dbo.TPRODUCTO.tMoneda LEFT OUTER JOIN
                      dbo.TSUBGRUPO ON dbo.TPRODUCTO.tSubGrupo = dbo.TSUBGRUPO.tCodigoSubgrupo LEFT OUTER JOIN
                      dbo.vTipoProducto ON dbo.TPRODUCTO.tTipoProducto = dbo.vTipoProducto.Codigo LEFT OUTER JOIN
                      dbo.vArea ON dbo.TPRODUCTO.tArea = dbo.vArea.Codigo LEFT OUTER JOIN
                      dbo.TGRUPO ON dbo.TPRODUCTO.tGrupo = dbo.TGRUPO.tCodigoGrupo



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vProductoArea
AS
SELECT     dbo.TPRODUCTOAREA.tCodigoProducto, dbo.TPRODUCTOAREA.tArea, dbo.vArea.Descripcion AS Area
FROM         dbo.TPRODUCTOAREA INNER JOIN
                      dbo.vArea ON dbo.TPRODUCTOAREA.tArea = dbo.vArea.Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE view dbo.vCombo as
select TCODIGOPEDIDO, TCODIGOGRUPO, TCODIGOSUBGRUPO, TCODIGOPRODUCTO, NCANTIDAD, NVENTA  
from DPEDIDO 
where lCombinacion = 0 and tEstadoItem ='N'
UNION ALL 
select TCODIGOPEDIDO, TCODIGOGRUPO, TCODIGOSUBGRUPO, TPRODUCTOCOMBO, NCANTIDAD, 0 AS NVENTA   
from CPEDIDO

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vComboDetalle
AS
SELECT     dbo.vCombo.TCODIGOPEDIDO, dbo.vCombo.TCODIGOGRUPO, dbo.vCombo.TCODIGOSUBGRUPO, dbo.vCombo.TCODIGOPRODUCTO, 
                      dbo.vCombo.NCANTIDAD, dbo.vCombo.NVENTA, dbo.TPRODUCTO.tDetallado AS Producto, dbo.TPRODUCTO.nPrecioVenta
FROM         dbo.TPRODUCTO RIGHT OUTER JOIN
                      dbo.vCombo ON dbo.TPRODUCTO.tCodigoProducto = dbo.vCombo.TCODIGOPRODUCTO

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vEgreso
AS
SELECT     dbo.MEGRESO.tRecibo, dbo.MEGRESO.tCaja, dbo.MEGRESO.tTurno, dbo.MEGRESO.fFecha, dbo.MEGRESO.tMoneda, dbo.MEGRESO.nTipoCambio, 
                      dbo.MEGRESO.nMonto, dbo.MEGRESO.tDescripcion, dbo.MEGRESO.tAutoriza, dbo.MEGRESO.tEstadoDocumento, dbo.MEGRESO.tUsuario, 
                      dbo.MEGRESO.fRegistro, dbo.MEGRESO.lReplica, dbo.MEGRESO.fDiaContable, ISNULL(dbo.MEGRESO.tTipoEgreso,'') TTIPOEGRESO, 
                      dbo.vEstadoDocumento.Descripcion AS Estado, dbo.vMoneda.tResumido AS Moneda, ISNULL(dbo.vTipoEgreso.Descripcion,'') AS TipoEgreso
FROM         dbo.vEstadoDocumento RIGHT OUTER JOIN
                      dbo.vTipoEgreso RIGHT OUTER JOIN
                      dbo.MEGRESO ON dbo.vTipoEgreso.Codigo = dbo.MEGRESO.tTipoEgreso ON 
                      dbo.vEstadoDocumento.Codigo = dbo.MEGRESO.tEstadoDocumento LEFT OUTER JOIN
                      dbo.vMoneda ON dbo.MEGRESO.tMoneda = dbo.vMoneda.Codigo
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vFacturacionDetalle
AS
SELECT     dbo.DPEDIDO.tCodigoProducto AS Codigo, dbo.vProducto.Descripcion AS Descripcion, dbo.DPEDIDO.nPrecioVenta AS nPrecio, 
                      SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(dbo.DPEDIDO.nVenta) AS nVenta, SUM(dbo.DPEDIDO.nImpuesto1) AS nImpuesto1, 
                      SUM(dbo.DPEDIDO.nImpuesto2) AS nImpuesto2, SUM(dbo.DPEDIDO.nImpuesto3) AS nImpuesto3, MIN(dbo.DPEDIDO.tItem) AS tItem
FROM         dbo.DPEDIDO LEFT OUTER JOIN
                      dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo
GROUP BY dbo.DPEDIDO.tCodigoProducto, dbo.DPEDIDO.nPrecioVenta, dbo.vProducto.Descripcion

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW dbo.vIngreso
AS
SELECT     dbo.MINGRESO.*, dbo.vMoneda.tResumido AS Moneda, dbo.vEstadoDocumento.Descripcion AS EstadoDocumento, 
                      dbo.vTipoPago.Descripcion AS TipoPago
FROM         dbo.MINGRESO LEFT OUTER JOIN
                      dbo.vTipoPago ON dbo.MINGRESO.tTipoPago = dbo.vTipoPago.Codigo LEFT OUTER JOIN
                      dbo.vMoneda ON dbo.MINGRESO.tMoneda = dbo.vMoneda.Codigo LEFT OUTER JOIN
                      dbo.vEstadoDocumento ON dbo.MINGRESO.tEstadoDocumento = dbo.vEstadoDocumento.Codigo


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vLiquidacion
AS
SELECT                dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tTipoDocumento, dbo.MDOCUMENTO.nVenta, 
                      dbo.MDOCUMENTO.tUsuarioAnulado, dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.fPago, dbo.MDOCUMENTO.tCodigoCliente, 
                      dbo.MDOCUMENTO.tObservacion AS Motivo, dbo.MDOCUMENTO.tTurno AS Turno, dbo.MDOCUMENTO.fRegistro, dbo.DPAGODOCUMENTO.tTurno, 
                      dbo.DPAGODOCUMENTO.tUsuario, dbo.DPAGODOCUMENTO.tTipoPago, dbo.DPAGODOCUMENTO.tMoneda, dbo.DPAGODOCUMENTO.nTipoCambio, 
                      dbo.DPAGODOCUMENTO.nMonto, dbo.DPAGODOCUMENTO.nPropina, dbo.DPAGODOCUMENTO.tTarjeta, dbo.DPAGODOCUMENTO.tBanco, 
                      dbo.DPAGODOCUMENTO.tNumero, dbo.vCliente.Descripcion AS Cliente, dbo.vCompania.Descripcion AS ClientePago, 
                      dbo.TTARJETACREDITO.tDetallado AS Tarjeta, dbo.vCortesia.Descripcion AS Cortesia, dbo.MINGRESO.tTurno AS Rturno, 
                      dbo.DPAGODOCUMENTO.tCorrelativo, dbo.DPAGODOCUMENTO.nDolar
FROM         dbo.TTARJETACREDITO RIGHT OUTER JOIN
                      dbo.MINGRESO RIGHT OUTER JOIN
                      dbo.DPAGODOCUMENTO ON dbo.MINGRESO.tRecibo = dbo.DPAGODOCUMENTO.tBanco ON 
                      dbo.TTARJETACREDITO.tCodigoTarjeta = dbo.DPAGODOCUMENTO.tTarjeta RIGHT OUTER JOIN
                      dbo.vCortesia RIGHT OUTER JOIN
                      dbo.MDOCUMENTO ON dbo.vCortesia.Codigo = dbo.MDOCUMENTO.tCortesia ON 
                      dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN
                      dbo.vCompania ON dbo.MDOCUMENTO.tClientePago = dbo.vCompania.Codigo LEFT OUTER JOIN
                      dbo.vCliente ON dbo.MDOCUMENTO.tCodigoCliente = dbo.vCliente.Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vNotaCredito
AS
SELECT     dbo.MNOTACREDITO.tNotaCredito, dbo.MNOTACREDITO.fFecha, dbo.MNOTACREDITO.tDocumento, dbo.MNOTACREDITO.nNeto, 
                      dbo.MNOTACREDITO.nImpuesto1, dbo.MNOTACREDITO.nImpuesto2, dbo.MNOTACREDITO.nImpuesto3, round(dbo.MNOTACREDITO.nVenta,2) as nVenta, 
                      dbo.MNOTACREDITO.tEstadoDocumento, dbo.MNOTACREDITO.tTurno, dbo.MNOTACREDITO.tCaja, 
                      dbo.MNOTACREDITO.tUsuario, dbo.MNOTACREDITO.fRegistro, dbo.MNOTACREDITO.tUsuarioAnulado, dbo.MNOTACREDITO.fRegistroAnulado, 
                      dbo.MNOTACREDITO.tObservacion, dbo.vEstadoDocumento.Descripcion AS Estadodocumento, dbo.MDOCUMENTO.nNeto AS nDocNeto, 
                      dbo.MDOCUMENTO.nPrecioImpuesto1 AS nDocImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2 AS nDocImpuesto2, 
                      dbo.MDOCUMENTO.nPrecioImpuesto3 AS nDocImpuesto3, round(dbo.MDOCUMENTO.nVenta,2) AS nDocVenta, dbo.vCliente.Descripcion AS Cliente, 
                      dbo.vCliente.tIdentidad AS Identidad,dbo.mnotacredito.fdiacontable
FROM         dbo.vEstadoDocumento RIGHT OUTER JOIN
                      dbo.MNOTACREDITO ON dbo.vEstadoDocumento.Codigo = dbo.MNOTACREDITO.tEstadoDocumento LEFT OUTER JOIN
                      dbo.vCliente RIGHT OUTER JOIN
                      dbo.MDOCUMENTO ON dbo.vCliente.Codigo = dbo.MDOCUMENTO.tCodigoCliente ON 
                      dbo.MNOTACREDITO.tDocumento = dbo.MDOCUMENTO.tDocumento

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPreCuenta
AS
SELECT     dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tCaja, 
                      dbo.DPEDIDO.tItem, TPRODUCTO_2.tResumido AS Producto, dbo.DPEDIDO.nPrecioOficial, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.nImpuesto1, 
                      dbo.DPEDIDO.nImpuesto2, dbo.DPEDIDO.nImpuesto3, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tComanda, dbo.MPEDIDO.tObservacion, 
                      dbo.MPEDIDO.tHabitacion, dbo.MPEDIDO.tPasajero, dbo.MPEDIDO.tReserva, dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, 
                      dbo.vTipoPedido.Descripcion AS TipoPedido, TPRODUCTO_1.tDetallado AS Combo, dbo.vMozo.Descripcion AS Mozo, 
                      dbo.MPEDIDO.nDescuento AS xDescuento, dbo.DPEDIDO.nDescuento, dbo.DPEDIDO.nRecargo, dbo.MPEDIDO.tTipoPedido, dbo.vDelivery.lpuntos, 
                      dbo.vDelivery.nDisponible, dbo.MPEDIDO.nAdulto, dbo.CPEDIDO.nCantidad AS nCantidadCombo, dbo.CPEDIDO.tItemCombo, dbo.vDelivery.Cliente, 
                      dbo.DPEDIDO.tObservacion AS tObservacionPedido, dbo.CPEDIDO.tObservacion AS tObservacionCombo, dbo.TOFERTA.tResumido AS tOferta, 
                      dbo.vMotivoDescuento.Descripcion AS Descuento
FROM         dbo.vSalon RIGHT OUTER JOIN
                      dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon RIGHT OUTER JOIN
                      dbo.vDelivery RIGHT OUTER JOIN
                      dbo.vMotivoDescuento RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vMotivoDescuento.Codigo = dbo.MPEDIDO.tDescuento LEFT OUTER JOIN
                      dbo.vMozo ON dbo.MPEDIDO.tMozo = dbo.vMozo.Codigo ON dbo.vDelivery.Codigo = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN
                      dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo ON dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa RIGHT OUTER JOIN
                      dbo.TPRODUCTO AS TPRODUCTO_2 RIGHT OUTER JOIN
                      dbo.TOFERTA RIGHT OUTER JOIN
                      dbo.DPEDIDO ON dbo.TOFERTA.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto AND dbo.TOFERTA.tOferta = dbo.DPEDIDO.tOferta LEFT OUTER JOIN
                      dbo.CPEDIDO LEFT OUTER JOIN
                      dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto ON 
                      dbo.DPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.CPEDIDO.tItem ON 
                      TPRODUCTO_2.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido
WHERE     (dbo.DPEDIDO.tEstadoItem = 'N') AND (ISNULL(dbo.DPEDIDO.tFacturado, '') = '')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPreCuentaDelivery
AS
SELECT     dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tCaja, 
                      dbo.DPEDIDO.tItem, TPRODUCTO_2.tResumido AS Producto, dbo.DPEDIDO.nPrecioOficial, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.nImpuesto1, 
                      dbo.DPEDIDO.nImpuesto2, dbo.DPEDIDO.nImpuesto3, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tMesa, 
                      dbo.DPEDIDO.tFacturado, TPRODUCTO_1.tDetallado AS Combo, dbo.vMozo.Descripcion AS Mozo, dbo.MPEDIDO.nDescuento AS xDescuento, 
                      dbo.DPEDIDO.nDescuento, dbo.DPEDIDO.nRecargo, dbo.CPEDIDO.tItemCombo, dbo.CPEDIDO.nCantidad AS nCantidadCombo, dbo.vDelivery.tApellido, 
                      dbo.vDelivery.tNombre, dbo.vDelivery.tDireccion, dbo.vDelivery.tTelefono, dbo.vDelivery.tReferencia, dbo.vDelivery.Zona, dbo.vDelivery.lpuntos, 
                      dbo.vDelivery.nDisponible, dbo.vDelivery.tResumido, dbo.DPEDIDO.tObservacion AS ObservacionDetalle, 
                      dbo.vDelivery.tObservacion AS ObservacionCliente, dbo.CPEDIDO.tObservacion AS ObservacionCombo, dbo.TOFERTA.tResumido AS tOferta, 
                      dbo.vMotivoDescuento.Descripcion AS Descuento
FROM         dbo.vSalon RIGHT OUTER JOIN
                      dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon RIGHT OUTER JOIN
                      dbo.vDelivery RIGHT OUTER JOIN
                      dbo.vMotivoDescuento RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vMotivoDescuento.Codigo = dbo.MPEDIDO.tDescuento LEFT OUTER JOIN
                      dbo.vMozo ON dbo.MPEDIDO.tMozo = dbo.vMozo.Codigo ON dbo.vDelivery.Codigo = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN
                      dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo ON dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa RIGHT OUTER JOIN
                      dbo.TPRODUCTO AS TPRODUCTO_2 RIGHT OUTER JOIN
                      dbo.TOFERTA RIGHT OUTER JOIN
                      dbo.DPEDIDO ON dbo.TOFERTA.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto AND dbo.TOFERTA.tOferta = dbo.DPEDIDO.tOferta LEFT OUTER JOIN
                      dbo.CPEDIDO LEFT OUTER JOIN
                      dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto ON 
                      dbo.DPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.CPEDIDO.tItem ON 
                      TPRODUCTO_2.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPreCuentaDetallada
AS
SELECT  TOP 100 PERCENT dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, 
                      dbo.MPEDIDO.tCaja, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tResumido AS Producto, dbo.DPEDIDO.nPrecioOficial, dbo.DPEDIDO.nCantidad, 
                      dbo.DPEDIDO.nImpuesto1, dbo.DPEDIDO.nImpuesto2, dbo.DPEDIDO.nImpuesto3, dbo.DPEDIDO.nVenta, dbo.MPEDIDO.tObservacion, 
                      dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, dbo.vTipoPedido.Descripcion AS TipoPedido, CONVERT(nvarchar, dbo.DPEDIDO.fRegistro, 103) 
                      AS Fecha, dbo.DPEDIDO.tUsuarioD
FROM         dbo.TPRODUCTO RIGHT OUTER JOIN
                      dbo.DPEDIDO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto LEFT OUTER JOIN
                      dbo.vTipoPedido RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN
                      dbo.vSalon RIGHT OUTER JOIN
                      dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON 
                      dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido
WHERE     (dbo.DPEDIDO.tEstadoItem = 'N') AND (ISNULL(dbo.DPEDIDO.tFacturado, '') = '')
ORDER BY dbo.DPEDIDO.fRegistro, dbo.DPEDIDO.tUsuarioD

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumento
AS
SELECT     dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, 
                      dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, 
                      dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.TCLIENTE.tEmpresa AS Cliente, dbo.TCLIENTE.tIdentidad AS RUC, 
                      dbo.TCLIENTE.tDireccion AS Direccion, dbo.DDOCUMENTO.tItem, dbo.DDOCUMENTO.nPrecioNeto, dbo.DDOCUMENTO.nPrecioVenta, 
                      dbo.DDOCUMENTO.nCantidad, dbo.TPRODUCTO.tResumido AS Producto, dbo.TPRODUCTO.lModificable, dbo.MDOCUMENTO.nPrecioOficial, 
                      dbo.DDOCUMENTO.nVenta AS Venta, dbo.DDOCUMENTO.tCodigoProducto, dbo.TPRODUCTO.tDetallado AS ProductoDetallado, 
                      dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, dbo.DDOCUMENTO.tCodigoPedido
FROM         dbo.MDOCUMENTO LEFT OUTER JOIN
                      dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente LEFT OUTER JOIN
                      dbo.TPRODUCTO RIGHT OUTER JOIN
                      dbo.DDOCUMENTO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DDOCUMENTO.tCodigoProducto ON 
                      dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumentoImpresora
AS
SELECT     dbo.MDOCUMENTO.tDocumento, dbo.MPEDIDO.tCodigoPedido, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, 
                      dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, 
                      dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.vCortesia.Descripcion AS Cortesia, dbo.TCLIENTE.tEmpresa AS Cliente, 
                      dbo.TCLIENTE.tIdentidad AS RUC, dbo.TCLIENTE.tDireccion AS Direccion, dbo.DDOCUMENTO.tItem, dbo.DDOCUMENTO.nPrecioNeto, 
                      dbo.DDOCUMENTO.nPrecioVenta, dbo.DDOCUMENTO.nCantidad, dbo.vMozo.Descripcion AS Mozo, 
                      dbo.vSalon.Descripcion + N' - ' + dbo.TMESA.tDetallado AS Mesa, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tMesa, 
                      TPRODUCTO_1.tResumido AS Producto, TPRODUCTO_1.lModificable, dbo.DDOCUMENTO.nPrecioOficial, dbo.DDOCUMENTO.nVenta AS Venta, 
                      dbo.MPEDIDO.nCorrelativo AS Orden, dbo.MPEDIDO.tTipoPedido, dbo.vTipoPedido.Descripcion AS TipoPedido, 
                      dbo.vTipodocumentoImpresora.lImpuesto1, dbo.vTipodocumentoImpresora.lImpuesto2, dbo.vTipodocumentoImpresora.lImpuesto3, 
                      dbo.DDOCUMENTO.tCodigoProducto, TPRODUCTO_1.tDetallado AS ProductoDetallado, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, 
                      TPRODUCTO_2.tResumido AS Combo, dbo.CPEDIDO.nCantidad AS nCombo, dbo.DDOCUMENTO.nPrecioImpuesto1 AS nImpuesto1, 
                      dbo.DDOCUMENTO.nPrecioImpuesto2 AS nImpuesto2, dbo.DDOCUMENTO.nPrecioImpuesto3 AS nImpuesto3, 
                      dbo.DPEDIDO.tObservacion AS tObservacionPedido, dbo.CPEDIDO.tItemCombo, dbo.CPEDIDO.tObservacion AS tObservacionCombo, 
                      dbo.TOFERTA.tResumido AS tOferta, dbo.vMotivoDescuento.Descripcion AS Descuento
FROM         dbo.TCLIENTE RIGHT OUTER JOIN
                      dbo.MDOCUMENTO LEFT OUTER JOIN
                      dbo.vMotivoDescuento ON dbo.MDOCUMENTO.tDescuento = dbo.vMotivoDescuento.Codigo LEFT OUTER JOIN
                      dbo.vTipodocumentoImpresora ON dbo.MDOCUMENTO.tCaja = dbo.vTipodocumentoImpresora.tCaja AND 
                      dbo.MDOCUMENTO.tTipoDocumento = dbo.vTipodocumentoImpresora.tTipoEmision ON 
                      dbo.TCLIENTE.tCodigoCliente = dbo.MDOCUMENTO.tCodigoCliente LEFT OUTER JOIN
                      dbo.CPEDIDO LEFT OUTER JOIN
                      dbo.TPRODUCTO AS TPRODUCTO_2 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_2.tCodigoProducto RIGHT OUTER JOIN
                      dbo.DPEDIDO LEFT OUTER JOIN
                      dbo.TOFERTA ON dbo.DPEDIDO.tCodigoProducto = dbo.TOFERTA.tCodigoProducto AND dbo.DPEDIDO.tOferta = dbo.TOFERTA.tOferta RIGHT OUTER JOIN
                      dbo.DDOCUMENTO ON dbo.DPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.DDOCUMENTO.tItem ON 
                      dbo.CPEDIDO.tItem = dbo.DDOCUMENTO.tItem AND dbo.CPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido LEFT OUTER JOIN
                      dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.DDOCUMENTO.tCodigoProducto = TPRODUCTO_1.tCodigoProducto ON 
                      dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento LEFT OUTER JOIN
                      dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN
                      dbo.vMozo RIGHT OUTER JOIN
                      dbo.vTipoPedido RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN
                      dbo.TMESA LEFT OUTER JOIN
                      dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON 
                      dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON dbo.DDOCUMENTO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumentoAgrupado
AS
SELECT DISTINCT 
                      dbo.MDOCUMENTO.tDocumento, dbo.DDOCUMENTO.tCodigoPedido, dbo.MDOCUMENTO.nVenta, dbo.vEstadoDocumento.Descripcion AS Estado, 
                      dbo.MDOCUMENTO.tUsuarioAnulado, dbo.MDOCUMENTO.fRegistroAnulado, dbo.MDOCUMENTO.tObservacion, dbo.MDOCUMENTO.tTurno
FROM         dbo.DDOCUMENTO RIGHT OUTER JOIN
                      dbo.MDOCUMENTO ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN
                      dbo.vEstadoDocumento ON dbo.MDOCUMENTO.tEstadoDocumento = dbo.vEstadoDocumento.Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumentoConsolidado
AS
SELECT     dbo.DDOCUMENTO.tCodigoProducto, dbo.DDOCUMENTO.nPrecioVenta, MAX(dbo.MDOCUMENTO.tDocumento) AS tDocumento, 
                      MAX(dbo.MPEDIDO.tCodigoPedido) AS tCodigoPedido, MAX(dbo.MDOCUMENTO.fRegistro) AS fRegistro, 
                      MAX(dbo.MDOCUMENTO.nNeto) AS nNeto, MAX(dbo.MDOCUMENTO.nRecargo) AS nRecargo, MAX(dbo.MDOCUMENTO.nDescuento) AS nDescuento, 
                      MAX(dbo.MDOCUMENTO.nPrecioImpuesto1) AS nPrecioImpuesto1, MAX(dbo.MDOCUMENTO.nPrecioImpuesto2) AS nPrecioImpuesto2, 
                      MAX(dbo.MDOCUMENTO.nPrecioImpuesto3) AS nPrecioImpuesto3, MAX(dbo.MDOCUMENTO.nVenta) AS nVenta, MAX(dbo.vCortesia.Descripcion) 
                      AS Cortesia, MAX(dbo.TCLIENTE.tEmpresa) AS Cliente, MAX(dbo.TCLIENTE.tIdentidad) AS RUC, MAX(dbo.TCLIENTE.tDireccion) AS Direccion, 
                      MAX(dbo.DDOCUMENTO.tItem) AS tItem, MAX(dbo.vMozo.Descripcion) AS Mozo, MAX(dbo.vSalon.Descripcion + N' - ' + dbo.TMESA.tDetallado) 
                      AS Mesa, MAX(dbo.MPEDIDO.tObservacion) AS tObservacion, MAX(dbo.MPEDIDO.tMesa) AS tMesa, MAX(dbo.TPRODUCTO.tResumido) AS Producto, 
                      MAX(dbo.MDOCUMENTO.nPrecioOficial) AS nPrecioOficial, SUM(dbo.DDOCUMENTO.nCantidad) AS nCantidad, SUM(dbo.DDOCUMENTO.nVenta) 
                      AS Venta
FROM         dbo.vMozo RIGHT OUTER JOIN
                      dbo.TMESA LEFT OUTER JOIN
                      dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo RIGHT OUTER JOIN
                      dbo.MDOCUMENTO LEFT OUTER JOIN
                      dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente LEFT OUTER JOIN
                      dbo.TPRODUCTO RIGHT OUTER JOIN
                      dbo.DDOCUMENTO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DDOCUMENTO.tCodigoProducto ON 
                      dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento LEFT OUTER JOIN
                      dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido
GROUP BY dbo.DDOCUMENTO.nPrecioVenta, dbo.DDOCUMENTO.tCodigoProducto

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumentoGrilla
AS
SELECT     dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa AS Cliente, 
                      dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tTipoDocumento, MAX(dbo.MDOCUMENTO.fRegistro) 
                      AS fFecha, dbo.MDOCUMENTO.tTurno, dbo.MDOCUMENTO.tUsuario, dbo.vCortesia.Descripcion AS Cortesia, MAX(dbo.TMESA.tResumido) AS Mesa, 
                      MAX(dbo.DDOCUMENTO.tCodigoPedido) AS tCodigoPedido, dbo.vEstadoDocumento.Descripcion AS EstadoDocumento, 
                      dbo.MDOCUMENTO.tClientePago, dbo.vCompania.Descripcion AS ClientePago, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, 
                      dbo.MDOCUMENTO.nPrecioImpuesto2, dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, 
                      MAX(dbo.MPEDIDO.tTipoPedido) AS tTipoPedido, MAX(dbo.MPEDIDO.tObservacion) AS tObservacion, 
                      dbo.vMotorizado.Descripcion AS Motorizado
FROM         dbo.MPEDIDO LEFT OUTER JOIN
                      dbo.vMotorizado ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo LEFT OUTER JOIN
                      dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN
                      dbo.DDOCUMENTO ON dbo.MPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido RIGHT OUTER JOIN
                      dbo.vCompania RIGHT OUTER JOIN
                      dbo.MDOCUMENTO ON dbo.vCompania.Codigo = dbo.MDOCUMENTO.tClientePago LEFT OUTER JOIN
                      dbo.vEstadoDocumento ON dbo.MDOCUMENTO.tEstadoDocumento = dbo.vEstadoDocumento.Codigo LEFT OUTER JOIN
                      dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN
                      dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON 
                      dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento
GROUP BY dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tCodigoCliente, dbo.TCLIENTE.tEmpresa, 
                      dbo.MDOCUMENTO.tEstadoDocumento, dbo.MDOCUMENTO.nVenta, dbo.MDOCUMENTO.tTipoDocumento, dbo.MDOCUMENTO.tTurno, 
                      dbo.MDOCUMENTO.tUsuario, dbo.vCortesia.Descripcion, dbo.vEstadoDocumento.Descripcion, dbo.MDOCUMENTO.tClientePago, 
                      dbo.vCompania.Descripcion, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, 
                      dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.vMotorizado.Descripcion

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumentoPago
AS
SELECT     dbo.DPAGODOCUMENTO.tDocumento, dbo.vTipoPago.Descripcion AS TipoPago, dbo.vMoneda.tResumido AS Moneda, 
                      dbo.DPAGODOCUMENTO.nMonto, dbo.DPAGODOCUMENTO.nPropina, dbo.DPAGODOCUMENTO.nTipoCambio, 
                      dbo.TTARJETACREDITO.tDetallado AS Tarjeta, dbo.DPAGODOCUMENTO.tBanco, dbo.DPAGODOCUMENTO.tNumero, 
                      dbo.DPAGODOCUMENTO.tReferencia, dbo.DPAGODOCUMENTO.tTurno, dbo.DPAGODOCUMENTO.fRegistro
FROM         dbo.DPAGODOCUMENTO LEFT OUTER JOIN
                      dbo.vMoneda ON dbo.DPAGODOCUMENTO.tMoneda = dbo.vMoneda.Codigo LEFT OUTER JOIN
                      dbo.TTARJETACREDITO ON dbo.DPAGODOCUMENTO.tTarjeta = dbo.TTARJETACREDITO.tCodigoTarjeta LEFT OUTER JOIN
                      dbo.vTipoPago ON dbo.DPAGODOCUMENTO.tTipoPago = dbo.vTipoPago.Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumentoResultado
AS
SELECT     tCodigoPedido, COUNT(tDocumento) AS tDocumento
FROM         dbo.vDocumentoAgrupado
GROUP BY tCodigoPedido

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedido
AS
SELECT     TPRODUCTO_1.tDetallado AS Producto, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tObservacion, dbo.DPEDIDO.tEstadoItem, 
                      dbo.DPEDIDO.tCodigoPedido AS Codigo, dbo.DPEDIDO.lImprime, dbo.DPEDIDO.lImprimeArea, TPRODUCTO_2.tResumido AS Combo, 
                      dbo.CPEDIDO.nCantidad AS nCombo, dbo.DPEDIDO.tItem, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, 
                      dbo.MPEDIDO.tObservacion AS Observacion, dbo.MPEDIDO.nCorrelativo AS Orden, dbo.MPEDIDO.tTipoPedido AS TipoPedido, 
                      dbo.MPEDIDO.lPrioridad AS Prioridad, dbo.MPEDIDO.nAdulto, dbo.vMozo.tResumido AS Mozo, dbo.DPEDIDO.nOrden, 
                      dbo.vDelivery.Cliente, dbo.MPEDIDO.tMotivoAnulacion AS tMotivoEliminacion, dbo.MPEDIDO.tobservacionAnulado, 
                      dbo.CPEDIDO.lImprimeArea AS lImprimeAreaCombo, dbo.CPEDIDO.lImprime AS lImprimeCombo, dbo.CPEDIDO.nOrden AS nOrdenCombo, dbo.CPEDIDO.tItemCombo AS tItemCombo, 
                      dbo.CPEDIDO.tObservacion AS tObservacionCombo, case TPRODUCTO_1.lCombinacion when 1 then TPRODUCTOAREA_1.tArea else dbo.TPRODUCTOAREA.tArea end as tArea,   dbo.DPEDIDO.lCorte AS lCorte,  dbo.CPEDIDO.lCorte AS lCorteCombo,  dbo.TCOMBO.tEtiqueta
FROM         dbo.TPRODUCTO AS TPRODUCTO_2 RIGHT OUTER JOIN
                      dbo.TCOMBO RIGHT OUTER JOIN
                      dbo.CPEDIDO ON dbo.TCOMBO.tCombo = dbo.CPEDIDO.tProducto AND 
                      dbo.TCOMBO.tCodigoProducto = dbo.CPEDIDO.tProductoCombo LEFT OUTER JOIN
                      dbo.TPRODUCTOAREA AS TPRODUCTOAREA_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTOAREA_1.tCodigoProducto ON 
                      TPRODUCTO_2.tCodigoProducto = dbo.CPEDIDO.tProductoCombo RIGHT OUTER JOIN
                      dbo.TPRODUCTOAREA RIGHT OUTER JOIN
                      dbo.DPEDIDO ON dbo.TPRODUCTOAREA.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto LEFT OUTER JOIN
                      dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.DPEDIDO.tCodigoProducto = TPRODUCTO_1.tCodigoProducto ON 
                      dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem AND dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido RIGHT OUTER JOIN
                      dbo.vMozo RIGHT OUTER JOIN
                      dbo.vDelivery RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vDelivery.Codigo = dbo.MPEDIDO.tClienteDelivery ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON 
                      dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido LEFT OUTER JOIN
                      dbo.TMESA LEFT OUTER JOIN
                      dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoAgrupado
AS
SELECT     tCodigoPedido, tDocumento, SUM(nVenta) AS nVenta
FROM         dbo.DPEDIDO
GROUP BY tCodigoPedido, tDocumento

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoCabecera
AS

SELECT     TOP (100) PERCENT dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.tCaja, dbo.MPEDIDO.tClienteDelivery, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.nCorrelativo, 
                      dbo.MPEDIDO.fProgramacion, dbo.MPEDIDO.nTiempo, dbo.MPEDIDO.tTipoPedido, dbo.MPEDIDO.tTipoAtencion, dbo.MPEDIDO.tEstadoPedido, dbo.MPEDIDO.tMesa, 
                      dbo.MPEDIDO.tMotorizado, dbo.MPEDIDO.lPrioridad, dbo.MPEDIDO.tObservacion, dbo.vTipoPedido.Descripcion AS TipoPedido, 
                      dbo.vTipoAtencion.Descripcion AS TipoAtencion, RTRIM(dbo.TDELIVERY.tApellido) + ', ' + RTRIM(dbo.TDELIVERY.tNombre) AS Cliente, 
                      dbo.TDELIVERY.tTelefono AS Telefono, dbo.vMozo.Descripcion AS Mozo, dbo.TMESA.tResumido AS Mesa, dbo.vMotorizado.Descripcion AS Motorizado, 
                      dbo.MPEDIDO.fProgramacion AS FecProg, dbo.MPEDIDO.nAdulto, dbo.MPEDIDO.nNino, dbo.MPEDIDO.tUsuario, dbo.MPEDIDO.tClienteCorp, 
                      dbo.TTIENDA.tNombre AS Tienda, dbo.MPEDIDO.tTienda, dbo.MPEDIDO.tTurno, dbo.TMESA.tSalon, dbo.MPEDIDO.tComanda, dbo.MPEDIDO.tPuntoVenta, 
                      dbo.MPEDIDO.tHabitacion, dbo.MPEDIDO.tReserva, dbo.MPEDIDO.tPasajero, dbo.MPEDIDO.nDescuento, dbo.MPEDIDO.tDescuento, 
                      dbo.MPEDIDO.tObservacionDescuento, dbo.MPEDIDO.tCompania, dbo.MPEDIDO.tContacto, dbo.MPEDIDO.tMozo, dbo.MPEDIDO.nMesa, 
                      dbo.MPEDIDO.tCodigoPedidoCD, dbo.MPEDIDO.tUsuarioDescuento, CASE isnull(tTienda, '') 
                      WHEN '' THEN dbo.TDELIVERY.tDireccion ELSE dbo.TTIENDA.tDireccion END AS Direccion, CASE ISNULL(mpedido.tcodigoinvitado, '') 
                      WHEN '' THEN '' ELSE TDELIVERYINVITADO.tNombre + ' ' + TDELIVERYINVITADO.tApellido END AS Invitado, dbo.MPEDIDO.tCodigoInvitado, dbo.mpedido.tcodigopariente
FROM         dbo.TMESA RIGHT OUTER JOIN
                      dbo.vMozo RIGHT OUTER JOIN
                      dbo.TDELIVERY RIGHT OUTER JOIN
                      dbo.TTIENDA RIGHT OUTER JOIN
                      dbo.TDELIVERYINVITADO RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.TDELIVERYINVITADO.tCodigoInvitado = dbo.MPEDIDO.tCodigoInvitado ON dbo.TTIENDA.tCodigoDelivery = dbo.MPEDIDO.tClienteDelivery AND 
                      dbo.TTIENDA.tCodigoTienda = dbo.MPEDIDO.tTienda ON dbo.TDELIVERY.tCodigoDelivery = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN
                      dbo.vMotorizado ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo LEFT OUTER JOIN
                      dbo.vTipoAtencion ON dbo.MPEDIDO.tTipoAtencion = dbo.vTipoAtencion.Codigo ON dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa LEFT OUTER JOIN
                      dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo
ORDER BY Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoCombo
AS
SELECT     dbo.CPEDIDO.tCodigoPedido, dbo.CPEDIDO.tProducto, dbo.CPEDIDO.tItem, dbo.CPEDIDO.tItemCombo, dbo.CPEDIDO.tProductoCombo, 
           dbo.CPEDIDO.nCantidad, dbo.CPEDIDO.tCodigoGrupo, dbo.CPEDIDO.tCodigoSubGrupo, dbo.TPRODUCTO.tDetallado AS Producto, 
           dbo.MPEDIDO.tEstadoPedido, dbo.MPEDIDO.tCaja, dbo.CPEDIDO.lImprimeArea, dbo.CPEDIDO.lImprime, dbo.CPEDIDO.nOrden, CONVERT(bit, 
           ISNULL(DATALENGTH(dbo.CPEDIDO.tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) AS lPropiedad, dbo.CPEDIDO.tObservacion, dbo.CPEDIDO.lCorte
FROM       dbo.CPEDIDO LEFT OUTER JOIN
           (SELECT     tCodigoPedido, tItem, tItemCombo, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END AS nPropiedad
           FROM          dbo.TCOMBOPROPIEDAD
           GROUP BY tCodigoPedido, tItem, tItemCombo) AS T1 ON dbo.CPEDIDO.tItemCombo = T1.tItemCombo AND dbo.CPEDIDO.tItem = T1.tItem AND 
           dbo.CPEDIDO.tCodigoPedido = T1.tCodigoPedido LEFT OUTER JOIN
           dbo.TPRODUCTO ON dbo.CPEDIDO.tProductoCombo = dbo.TPRODUCTO.tCodigoProducto LEFT OUTER JOIN
           dbo.MPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoResultado
AS
SELECT     tCodigoPedido, COUNT(tDocumento) AS Total, MAX(tDocumento) AS tDocumento, SUM(nVenta) AS nVenta
FROM         dbo.vPedidoAgrupado
GROUP BY tCodigoPedido

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoCorrelativo
AS
SELECT     dbo.MPEDIDO.tCodigoPedido, dbo.vEstadoPedido.Descripcion AS Estado, dbo.MPEDIDO.fFecha, dbo.vPedidoResultado.nVenta, 
                      'Documento' = CASE WHEN dbo.vPedidoResultado.Total > 1 THEN 'Varios' ELSE dbo.vPedidoResultado.tDocumento END,   
                      dbo.MPEDIDO.tUsuario, dbo.MPEDIDO.tCaja, dbo.MPEDIDO.tusuarioAnulado, dbo.MPEDIDO.tobservacionAnulado, dbo.MPEDIDO.tTipoPedido, 
                      dbo.TMESA.tDetallado AS Mesa, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tTurno, dbo.MPEDIDO.tEstadoPedido, 
                      dbo.vCompania.Descripcion AS Cliente, dbo.TMESA.tSalon, dbo.MPEDIDO.tComanda, dbo.MPEDIDO.tReserva, dbo.MPEDIDO.tHabitacion, 
                      dbo.MPEDIDO.tPasajero, dbo.MPROPINA.nmonto AS nPropina, dbo.vMoneda.tResumido AS tMonedaPropina, dbo.MPEDIDO.tPuntoVenta, 
                      dbo.MPEDIDO.tMozo, dbo.MPEDIDO.nAdulto, dbo.vTipoPedido.Descripcion AS TipoPedido, dbo.vMozo.Descripcion AS Mozo, dbo.MPEDIDO.tTipoComanda, dbo.MPEDIDO.tFichaPasajero
FROM         dbo.MPEDIDO LEFT OUTER JOIN
                      dbo.vMozo ON dbo.MPEDIDO.tMozo = dbo.vMozo.Codigo LEFT OUTER JOIN
                      dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo LEFT OUTER JOIN
                      dbo.vCompania ON dbo.MPEDIDO.tClienteCorp = dbo.vCompania.Codigo LEFT OUTER JOIN
                      dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa LEFT OUTER JOIN
                      dbo.vEstadoPedido ON dbo.MPEDIDO.tEstadoPedido = dbo.vEstadoPedido.Codigo LEFT OUTER JOIN
                      dbo.vPedidoResultado ON dbo.MPEDIDO.tCodigoPedido = dbo.vPedidoResultado.tCodigoPedido LEFT OUTER JOIN
                      dbo.vDocumentoResultado ON dbo.MPEDIDO.tCodigoPedido = dbo.vDocumentoResultado.tCodigoPedido LEFT OUTER JOIN
                      dbo.MPROPINA ON dbo.MPEDIDO.tCodigoPedido = dbo.MPROPINA.tcodigopedido LEFT OUTER JOIN
                      dbo.vMoneda ON dbo.MPROPINA.tmoneda = dbo.vMoneda.Codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoDetalle
AS
SELECT     TOP 100 PERCENT dbo.DPEDIDO.*, dbo.TPRODUCTO.tDetallado AS Producto, dbo.vCortesia.Descripcion AS Cortesia, 
           dbo.vMozo.Descripcion AS MozoD, CASE dbo.DPEDIDO.nDescuento WHEN 0 THEN 0 ELSE dbo.DPEDIDO.nDescuento * 100 / dbo.DPEDIDO.nPrecioOficial END AS Descuento, 
           dbo.TPRODUCTO.lDescuento AS lDescuento, dbo.TPRODUCTO.lModificable AS lModificable, 
           CONVERT(bit, ISNULL(DATALENGTH(dbo.DPEDIDO.tObservacion), 0)) AS lObservacion, 
           ISNULL(T1.nPropiedad,0) as lPropiedad
FROM       dbo.DPEDIDO LEFT OUTER JOIN
           (SELECT tCodigoPedido, tItem, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END as nPropiedad FROM dbo.TPRODUCTOPROPIEDAD GROUP BY tCodigoPedido, tItem ) as T1              
           ON dbo.DPEDIDO.tItem = T1.tItem AND dbo.DPEDIDO.tCodigoPedido = T1.tCodigoPedido LEFT OUTER JOIN
           dbo.vMozo ON dbo.DPEDIDO.tMozoD = dbo.vMozo.Codigo LEFT OUTER JOIN
           dbo.vCortesia ON dbo.DPEDIDO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN
           dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto LEFT OUTER JOIN
           dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido
ORDER BY dbo.DPEDIDO.tCodigoPedido, dbo.DPEDIDO.tItem

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoGrilla
AS
SELECT     dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.tCodigoPedido AS Descripcion, dbo.MPEDIDO.tCaja, dbo.MPEDIDO.tObservacion, 
                      dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tUsuario, dbo.vTipoPedido.Descripcion AS TipoPedido, dbo.MPEDIDO.tEstadoPedido, 
                      dbo.MPEDIDO.tTipoPedido, SUM(dbo.DPEDIDO.nVenta) AS Suma, dbo.MPEDIDO.tClienteCorp, dbo.MPEDIDO.tMozo, dbo.MPEDIDO.tTurno, 
                      dbo.vMozo.Descripcion AS Mozo
FROM         dbo.vMozo RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo LEFT OUTER JOIN
                      dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa LEFT OUTER JOIN
                      dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo LEFT OUTER JOIN
                      dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido
GROUP BY dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.tCaja, dbo.MPEDIDO.tObservacion, dbo.TMESA.tResumido, 
                      dbo.MPEDIDO.tUsuario, dbo.vTipoPedido.Descripcion, dbo.MPEDIDO.tEstadoPedido, dbo.MPEDIDO.tEstadoPedido, dbo.MPEDIDO.tTipoPedido, 
                      dbo.MPEDIDO.tClienteCorp, dbo.MPEDIDO.tMozo, dbo.MPEDIDO.tTurno, dbo.vMozo.Descripcion

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vOferta
AS
SELECT DISTINCT 
                      dbo.TOFERTA.tOferta, dbo.TOFERTA.tNombre, dbo.TOFERTA.tResumido, dbo.TOFERTA.tFrecuencia, dbo.TOFERTA.fFecha, 
                      dbo.TOFERTA.tHoraInicial, dbo.TOFERTA.tHoraFinal, dbo.TOFERTA.lAcumulable, dbo.TOFERTA.nRatio, dbo.TOFERTA.nMonto, dbo.TOFERTA.nPrecio,
                      dbo.TOFERTA.lPermanente, dbo.TOFERTA.fFechaInicial, dbo.TOFERTA.fFechaFinal, dbo.TOFERTA.lLocal, dbo.TOFERTA.lDelivery, 
                      dbo.TOFERTA.lLlevar, dbo.TOFERTA.lCanal4, dbo.TOFERTA.lCanal5, dbo.TOFERTA.lExcluyente, dbo.TOFERTA.lAutomatica, dbo.TOFERTA.lActivo
FROM         dbo.TOFERTA 
                      

GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vMotivoEliminacion
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, nValor
FROM         dbo.TTABLA
WHERE     (TTABLA = N'MOTIVOELIMINACION')

GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTipoCancelacion
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, nValor, tValor
FROM         dbo.TTABLA
WHERE     (TTABLA = N'TIPOCANCELACION')

GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.vChofer
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'CHOFER')


GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vVehiculo
AS
SELECT     SUBSTRING(TCODIGO, 1, 3) AS Codigo, tDetallado AS Marca, tResumido AS Placa, tValor AS Licencia, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'VEHICULO')

GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.vMotivoTraslado
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'MOTIVOTRASLADO')


GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW dbo.vPrecuentaAgrupada
AS
SELECT     dbo.MPEDIDO.tCodigoPedido AS Codigo, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, 
                      dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tCaja, dbo.MPEDIDO.nAdulto, TPRODUCTO_2.tResumido AS Producto, 
                      dbo.DPEDIDO.nPrecioOficial, dbo.vDelivery.lpuntos, dbo.vDelivery.nDisponible, MAX(dbo.DPEDIDO.tItem) AS tItem, SUM(dbo.DPEDIDO.nCantidad) 
                      AS nCantidad, SUM(dbo.DPEDIDO.nImpuesto1) AS nImpuesto1, SUM(dbo.DPEDIDO.nImpuesto2) AS nImpuesto2, SUM(dbo.DPEDIDO.nImpuesto3) 
                      AS nImpuesto3, SUM(dbo.DPEDIDO.nVenta) AS nVenta, dbo.MPEDIDO.tComanda, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tHabitacion, 
                      dbo.MPEDIDO.tPasajero, dbo.MPEDIDO.tReserva, dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, dbo.vTipoPedido.Descripcion AS TipoPedido, 
                      dbo.vMozo.Descripcion AS Mozo, dbo.MPEDIDO.nDescuento AS xDescuento, dbo.vDelivery.Cliente, AVG(dbo.DPEDIDO.nDescuento) AS nDescuento, 
                      SUM(dbo.DPEDIDO.nRecargo) AS nRecargo, '' AS Combo, dbo.vMotivoDescuento.Descripcion AS Descuento
FROM         dbo.vDelivery RIGHT OUTER JOIN
                      dbo.vMotivoDescuento RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vMotivoDescuento.Codigo = dbo.MPEDIDO.tDescuento LEFT OUTER JOIN
                      dbo.vMozo ON dbo.MPEDIDO.tMozo = dbo.vMozo.Codigo ON dbo.vDelivery.Codigo = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN
                      dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo LEFT OUTER JOIN
                      dbo.vSalon RIGHT OUTER JOIN
                      dbo.TMESA ON dbo.vSalon.Codigo = dbo.TMESA.tSalon ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa RIGHT OUTER JOIN
                      dbo.DPEDIDO LEFT OUTER JOIN
                      dbo.TPRODUCTO AS TPRODUCTO_2 ON dbo.DPEDIDO.tCodigoProducto = TPRODUCTO_2.tCodigoProducto ON 
                      dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido
WHERE     (dbo.DPEDIDO.tEstadoItem = N'N')
GROUP BY dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tTipoPedido, dbo.vSalon.tResumido + N' - ' + dbo.TMESA.tResumido, 
                      dbo.MPEDIDO.tCaja, dbo.MPEDIDO.nAdulto, TPRODUCTO_2.tResumido, dbo.MPEDIDO.tComanda, dbo.MPEDIDO.tObservacion, 
                      dbo.MPEDIDO.tHabitacion, dbo.MPEDIDO.tPasajero, dbo.MPEDIDO.tReserva, dbo.MPEDIDO.tMesa, dbo.DPEDIDO.tFacturado, 
                      dbo.vTipoPedido.Descripcion, dbo.vMozo.Descripcion, dbo.MPEDIDO.nDescuento, dbo.DPEDIDO.nPrecioOficial, dbo.vDelivery.nDisponible, 
                      dbo.vDelivery.lpuntos, dbo.MPEDIDO.tTipoPedido, dbo.vDelivery.Cliente, dbo.vMotivoDescuento.Descripcion
HAVING      (ISNULL(dbo.DPEDIDO.tFacturado, '') = '')

GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW dbo.vGrupoUsuario
AS

SELECT     dbo.TUSUARIO.tCodigoUsuario, dbo.TUSUARIO.tDetallado, dbo.TUSUARIO.tResumido, dbo.TUSUARIO.tPassword, dbo.TUSUARIO. tBandaMagnetica, dbo.TUSUARIO.lActivo, 
                      dbo.TGRUPOUSUARIO.lModulo01, dbo.TGRUPOUSUARIO.lModulo02, dbo.TGRUPOUSUARIO.lModulo03, dbo.TGRUPOUSUARIO.lOpcion01, 
                      dbo.TGRUPOUSUARIO.lOpcion02, dbo.TGRUPOUSUARIO.lOpcion03, dbo.TGRUPOUSUARIO.lOpcion04, dbo.TGRUPOUSUARIO.lOpcion05, 
                      dbo.TGRUPOUSUARIO.lOpcion06, dbo.TGRUPOUSUARIO.lOpcion07, dbo.TGRUPOUSUARIO.lOpcion08, dbo.TGRUPOUSUARIO.lOpcion09, 
                      dbo.TGRUPOUSUARIO.lOpcion10, dbo.TGRUPOUSUARIO.lOpcion11, dbo.TGRUPOUSUARIO.lOpcion12, dbo.TGRUPOUSUARIO.lOpcion13, 
                      dbo.TGRUPOUSUARIO.lOpcion14, dbo.TGRUPOUSUARIO.lOpcion15, dbo.TGRUPOUSUARIO.lOpcion16, dbo.TGRUPOUSUARIO.lOpcion17,
                      dbo.TGRUPOUSUARIO.lOpcion18, dbo.TGRUPOUSUARIO.lOpcion19, dbo.TGRUPOUSUARIO.lOpcion20 , dbo.TGRUPOUSUARIO.lOpcion21,
                      dbo.TGRUPOUSUARIO.lModulo04, dbo.TGRUPOUSUARIO.lModulo05, dbo.TUSUARIO.tHuella, dbo.TGRUPOUSUARIO.lActivo AS ActivoGrupo
FROM         dbo.TUSUARIO INNER JOIN
                      dbo.TGRUPOUSUARIO ON dbo.TUSUARIO.tGrupoUsuario = dbo.TGRUPOUSUARIO.tGrupoUsuario

GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumentoImpresoraAgrupado
AS
SELECT     dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, 
                      dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, 
                      dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.DDOCUMENTO.nPrecioOficial, dbo.MDOCUMENTO.tCaja, 
                      dbo.MDOCUMENTO.tUsuario, dbo.vTipodocumentoImpresora.lImpuesto1, dbo.vTipodocumentoImpresora.lImpuesto2, 
                      dbo.vTipodocumentoImpresora.lImpuesto3, dbo.vCortesia.Descripcion AS Cortesia, dbo.TCLIENTE.tEmpresa AS Cliente, 
                      dbo.TCLIENTE.tIdentidad AS RUC, dbo.TCLIENTE.tDireccion AS Direccion, dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.tObservacion, 
                      dbo.MPEDIDO.tMesa, dbo.MPEDIDO.nCorrelativo AS Orden, dbo.MPEDIDO.tTipoPedido, dbo.vMozo.Descripcion AS Mozo, 
                      dbo.vTipoPedido.Descripcion AS TipoPedido, dbo.DDOCUMENTO.tItem, dbo.DDOCUMENTO.nPrecioNeto, dbo.DDOCUMENTO.nPrecioVenta, 
                      dbo.DDOCUMENTO.nCantidad, dbo.DDOCUMENTO.nVenta AS Venta, dbo.DDOCUMENTO.tCodigoProducto, 
                      dbo.vSalon.Descripcion + N' - ' + dbo.TMESA.tDetallado AS Mesa, TPRODUCTO_2.tResumido AS Producto, '' AS Combo, 
                      dbo.TraePropiedad(dbo.MPEDIDO.tCodigoPedido, dbo.DDOCUMENTO.tItem) AS Propiedad, vMotivoDescuento_1.Descripcion AS Descuento,
                      dbo.DDOCUMENTO.nPrecioImpuesto1 AS nImpuesto1, dbo.DDOCUMENTO.nPrecioImpuesto2 AS nImpuesto2, dbo.DDOCUMENTO.nPrecioImpuesto3 AS nImpuesto3
FROM         dbo.TPRODUCTO AS TPRODUCTO_2 RIGHT OUTER JOIN
                      dbo.DDOCUMENTO ON TPRODUCTO_2.tCodigoProducto = dbo.DDOCUMENTO.tCodigoProducto RIGHT OUTER JOIN
                      dbo.TCLIENTE RIGHT OUTER JOIN
                      dbo.MDOCUMENTO LEFT OUTER JOIN
                      dbo.vMotivoDescuento AS vMotivoDescuento_1 ON dbo.MDOCUMENTO.tDescuento = vMotivoDescuento_1.Codigo LEFT OUTER JOIN
                      dbo.vTipodocumentoImpresora ON dbo.MDOCUMENTO.tCaja = dbo.vTipodocumentoImpresora.tCaja AND 
                      dbo.MDOCUMENTO.tTipoDocumento = dbo.vTipodocumentoImpresora.tTipoEmision ON 
                      dbo.TCLIENTE.tCodigoCliente = dbo.MDOCUMENTO.tCodigoCliente ON 
                      dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN
                      dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN
                      dbo.vMozo RIGHT OUTER JOIN
                      dbo.vTipoPedido RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN
                      dbo.TMESA LEFT OUTER JOIN
                      dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON 
                      dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON dbo.DDOCUMENTO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vEmpacador
AS
SELECT     SUBSTRING(TCODIGO, 1, 4) AS Codigo, tDetallado AS Descripcion, tResumido, tValor, lActivo
FROM         dbo.TTABLA
WHERE     (TTABLA = 'EMPACADOR')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDespachador
AS
SELECT     dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.nCorrelativo, dbo.MPEDIDO.tClienteDelivery, dbo.MPEDIDO.fFecha, dbo.MPEDIDO.tCaja, 
                      dbo.MPEDIDO.tMotorizado, dbo.MPEDIDO.tUsuario, dbo.MPEDIDO.tTipoPedido, dbo.MPEDIDO.tEstadoPedido, LTRIM(dbo.TDELIVERY.tApellido) 
                      + ' ' + LTRIM(dbo.TDELIVERY.tNombre) AS Cliente, dbo.vMotorizado.Descripcion AS Motorizado, dbo.vEmpacador.Descripcion AS Empacador, 
                      dbo.MPEDIDO.fAsignacion, dbo.MPEDIDO.fSalida, dbo.MPEDIDO.fLlegada, dbo.TDELIVERY.tTelefono, dbo.TDELIVERY.tDireccion, 
                      dbo.TDELIVERY.tReferencia, dbo.vZona.Descripcion AS Zona, dbo.vZona.tResumido AS Referencia, ISNULL(DATEDIFF(mi, GETDATE(), DATEADD(mi, 
                      dbo.vZona.nValor, dbo.MPEDIDO.fSalida)), 0) AS Restante
FROM         dbo.vMotorizado RIGHT OUTER JOIN
                      dbo.TDELIVERY LEFT OUTER JOIN
                      dbo.vZona ON dbo.TDELIVERY.tZona = dbo.vZona.Codigo RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.TDELIVERY.tCodigoDelivery = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN
                      dbo.vEmpacador ON dbo.MPEDIDO.tEmpacador = dbo.vEmpacador.Codigo ON dbo.vMotorizado.Codigo = dbo.MPEDIDO.tMotorizado

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPRODUCTOXPRODUCTO
AS
SELECT     dbo.TPRODUCTOXPRODUCTO.tCodigoProducto, A.tDetallado, dbo.TPRODUCTOXPRODUCTO.tSubProducto, B.tDetallado AS Producto, 
           dbo.TPRODUCTOXPRODUCTO.nCantidad
FROM       dbo.TPRODUCTOXPRODUCTO LEFT OUTER JOIN
           dbo.TPRODUCTO A ON A.tCodigoProducto = dbo.TPRODUCTOXPRODUCTO.tCodigoProducto LEFT OUTER JOIN
           dbo.TPRODUCTO B ON B.tCodigoProducto = dbo.TPRODUCTOXPRODUCTO.tSubProducto

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW dbo.vDocumentoCorrelativoDetalle
AS
SELECT     dbo.MDOCUMENTO.tDocumento, dbo.MPEDIDO.tCodigoPedido, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, 
                      dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, 
                      dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.vCortesia.Descripcion AS Cortesia, dbo.TCLIENTE.tEmpresa AS Cliente, 
                      dbo.TCLIENTE.tIdentidad AS RUC, dbo.TCLIENTE.tDireccion AS Direccion, dbo.DDOCUMENTO.tItem, dbo.DDOCUMENTO.nPrecioNeto, 
                      dbo.DDOCUMENTO.nPrecioVenta, dbo.DDOCUMENTO.nCantidad, dbo.vMozo.Descripcion AS Mozo, 
                      dbo.vSalon.Descripcion + N' - ' + dbo.TMESA.tDetallado AS Mesa, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tMesa, 
                      TPRODUCTO_1.tResumido AS Producto, TPRODUCTO_1.lModificable, dbo.MDOCUMENTO.nPrecioOficial, dbo.DDOCUMENTO.nVenta AS Venta, 
                      dbo.MPEDIDO.nCorrelativo AS Orden, dbo.MPEDIDO.tTipoPedido, dbo.vTipoPedido.Descripcion AS TipoPedido, 
                      dbo.vTipodocumentoImpresora.lImpuesto1, dbo.vTipodocumentoImpresora.lImpuesto2, dbo.vTipodocumentoImpresora.lImpuesto3, 
                      dbo.DDOCUMENTO.tCodigoProducto, TPRODUCTO_1.tDetallado AS ProductoDetallado, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, 
                      dbo.DDOCUMENTO.nPrecioImpuesto1 AS nImpuesto1, dbo.DDOCUMENTO.nPrecioImpuesto2 AS nImpuesto2, 
                      dbo.DDOCUMENTO.nPrecioImpuesto3 AS nImpuesto3, dbo.DPEDIDO.tObservacion AS tObservacionPedido, dbo.TOFERTA.tResumido AS tOferta
FROM         dbo.TPRODUCTO AS TPRODUCTO_1 RIGHT OUTER JOIN
                      dbo.DPEDIDO LEFT OUTER JOIN
                      dbo.TOFERTA ON dbo.DPEDIDO.tCodigoProducto = dbo.TOFERTA.tCodigoProducto AND dbo.DPEDIDO.tOferta = dbo.TOFERTA.tOferta RIGHT OUTER JOIN
                      dbo.DDOCUMENTO ON dbo.DPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.DDOCUMENTO.tItem ON 
                      TPRODUCTO_1.tCodigoProducto = dbo.DDOCUMENTO.tCodigoProducto RIGHT OUTER JOIN
                      dbo.MDOCUMENTO LEFT OUTER JOIN
                      dbo.vTipodocumentoImpresora ON dbo.MDOCUMENTO.tCaja = dbo.vTipodocumentoImpresora.tCaja AND 
                      dbo.MDOCUMENTO.tTipoDocumento = dbo.vTipodocumentoImpresora.tTipoEmision LEFT OUTER JOIN
                      dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON 
                      dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN
                      dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN
                      dbo.vMozo RIGHT OUTER JOIN
                      dbo.vTipoPedido RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN
                      dbo.TMESA LEFT OUTER JOIN
                      dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON 
                      dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON dbo.DDOCUMENTO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW dbo.vDocumentoImpresoraAgrupadoAlternativa
AS
SELECT     dbo.MDOCUMENTO.tDocumento, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, dbo.MDOCUMENTO.nRecargo, 
                      dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, 
                      dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.DDOCUMENTO.nPrecioOficial, dbo.MDOCUMENTO.tCaja, 
                      dbo.MDOCUMENTO.tUsuario, dbo.vTipodocumentoImpresora.lImpuesto1, dbo.vTipodocumentoImpresora.lImpuesto2, 
                      dbo.vTipodocumentoImpresora.lImpuesto3, dbo.vCortesia.Descripcion AS Cortesia, dbo.TCLIENTE.tEmpresa AS Cliente, 
                      dbo.TCLIENTE.tIdentidad AS RUC, dbo.TCLIENTE.tDireccion AS Direccion, dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.tObservacion, 
                      dbo.MPEDIDO.tMesa, dbo.MPEDIDO.nCorrelativo AS Orden, dbo.MPEDIDO.tTipoPedido, dbo.vMozo.Descripcion AS Mozo, 
                      dbo.vTipoPedido.Descripcion AS TipoPedido, dbo.DDOCUMENTO.tItem, dbo.DDOCUMENTO.nPrecioNeto, dbo.DDOCUMENTO.nPrecioVenta, 
                      dbo.DDOCUMENTO.nCantidad, dbo.DDOCUMENTO.nVenta AS Venta, dbo.DDOCUMENTO.tCodigoProducto, 
                      dbo.vSalon.Descripcion + N' - ' + dbo.TMESA.tDetallado AS Mesa, CASE WHEN TPRODUCTO_2.talternativa IS NULL 
                      THEN TPRODUCTO_2.tresumido ELSE (CASE WHEN len(TPRODUCTO_2.talternativa) 
                      = 0 THEN TPRODUCTO_2.tresumido ELSE TPRODUCTO_2.talternativa END) END AS Producto, '' AS Combo, 
                      dbo.TraePropiedad(dbo.MPEDIDO.tCodigoPedido, dbo.DDOCUMENTO.tItem) AS Propiedad, dbo.vMotivoDescuento.Descripcion AS Descuento,
                      dbo.DDOCUMENTO.nPrecioImpuesto1 AS nImpuesto1, dbo.DDOCUMENTO.nPrecioImpuesto2 AS nImpuesto2, dbo.DDOCUMENTO.nPrecioImpuesto3 AS nImpuesto3
FROM         dbo.TPRODUCTO AS TPRODUCTO_2 RIGHT OUTER JOIN
                      dbo.DDOCUMENTO ON TPRODUCTO_2.tCodigoProducto = dbo.DDOCUMENTO.tCodigoProducto RIGHT OUTER JOIN
                      dbo.TCLIENTE RIGHT OUTER JOIN
                      dbo.MDOCUMENTO LEFT OUTER JOIN
                      dbo.vMotivoDescuento ON dbo.MDOCUMENTO.tDescuento = dbo.vMotivoDescuento.Codigo LEFT OUTER JOIN
                      dbo.vTipodocumentoImpresora ON dbo.MDOCUMENTO.tCaja = dbo.vTipodocumentoImpresora.tCaja AND 
                      dbo.MDOCUMENTO.tTipoDocumento = dbo.vTipodocumentoImpresora.tTipoEmision ON 
                      dbo.TCLIENTE.tCodigoCliente = dbo.MDOCUMENTO.tCodigoCliente ON 
                      dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN
                      dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN
                      dbo.vMozo RIGHT OUTER JOIN
                      dbo.vTipoPedido RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN
                      dbo.TMESA LEFT OUTER JOIN
                      dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON 
                      dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON dbo.DDOCUMENTO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vDocumentoImpresoraAlternativa
AS

SELECT     dbo.MDOCUMENTO.tDocumento, dbo.MPEDIDO.tCodigoPedido, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nNeto, 
                      dbo.MDOCUMENTO.nRecargo, dbo.MDOCUMENTO.nDescuento, dbo.MDOCUMENTO.nPrecioImpuesto1, dbo.MDOCUMENTO.nPrecioImpuesto2, 
                      dbo.MDOCUMENTO.nPrecioImpuesto3, dbo.MDOCUMENTO.nVenta, dbo.vCortesia.Descripcion AS Cortesia, dbo.TCLIENTE.tEmpresa AS Cliente, 
                      dbo.TCLIENTE.tIdentidad AS RUC, dbo.TCLIENTE.tDireccion AS Direccion, dbo.DDOCUMENTO.tItem, dbo.DDOCUMENTO.nPrecioNeto, 
                      dbo.DDOCUMENTO.nPrecioVenta, dbo.DDOCUMENTO.nCantidad, dbo.vMozo.Descripcion AS Mozo, 
                      dbo.vSalon.Descripcion + N' - ' + dbo.TMESA.tDetallado AS Mesa, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tMesa, 
					  case when TPRODUCTO_1.talternativa is null then TPRODUCTO_1.tresumido else (case when len(TPRODUCTO_1.talternativa)=0 then TPRODUCTO_1.tresumido else TPRODUCTO_1.talternativa end) end 
                      AS Producto,
					  TPRODUCTO_1.lModificable, dbo.DDOCUMENTO.nPrecioOficial, dbo.DDOCUMENTO.nVenta AS Venta, 
                      dbo.MPEDIDO.nCorrelativo AS Orden, dbo.MPEDIDO.tTipoPedido, dbo.vTipoPedido.Descripcion AS TipoPedido, 
                      dbo.vTipodocumentoImpresora.lImpuesto1, dbo.vTipodocumentoImpresora.lImpuesto2, dbo.vTipodocumentoImpresora.lImpuesto3, 
                      dbo.DDOCUMENTO.tCodigoProducto, TPRODUCTO_1.tDetallado AS ProductoDetallado, dbo.MDOCUMENTO.tCaja, dbo.MDOCUMENTO.tUsuario, 
					  case when TPRODUCTO_2.talternativa is null then TPRODUCTO_2.tresumido else (case when len(TPRODUCTO_2.talternativa)=0 then TPRODUCTO_2.tresumido else TPRODUCTO_2.talternativa end) end 
   					  AS Combo, 
					  dbo.CPEDIDO.nCantidad AS nCombo, dbo.DDOCUMENTO.nPrecioImpuesto1 AS nImpuesto1, 
                      dbo.DDOCUMENTO.nPrecioImpuesto2 AS nImpuesto2, dbo.DDOCUMENTO.nPrecioImpuesto3 AS nImpuesto3, 
                      dbo.DPEDIDO.tObservacion AS tObservacionPedido, dbo.CPEDIDO.tItemCombo, dbo.CPEDIDO.tObservacion AS tObservacionCombo, 
                      dbo.TOFERTA.tResumido AS tOferta
FROM         dbo.CPEDIDO LEFT OUTER JOIN
                      dbo.TPRODUCTO TPRODUCTO_2 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_2.tCodigoProducto RIGHT OUTER JOIN
                      dbo.DPEDIDO LEFT OUTER JOIN
                      dbo.TOFERTA ON dbo.DPEDIDO.tCodigoProducto = dbo.TOFERTA.tCodigoProducto AND dbo.DPEDIDO.tOferta = dbo.TOFERTA.tOferta RIGHT OUTER JOIN
                      dbo.DDOCUMENTO ON dbo.DPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.DDOCUMENTO.tItem ON 
                      dbo.CPEDIDO.tItem = dbo.DDOCUMENTO.tItem AND dbo.CPEDIDO.tCodigoPedido = dbo.DDOCUMENTO.tCodigoPedido LEFT OUTER JOIN
                      dbo.TPRODUCTO TPRODUCTO_1 ON dbo.DDOCUMENTO.tCodigoProducto = TPRODUCTO_1.tCodigoProducto RIGHT OUTER JOIN
                      dbo.MDOCUMENTO LEFT OUTER JOIN
                      dbo.vTipodocumentoImpresora ON dbo.MDOCUMENTO.tCaja = dbo.vTipodocumentoImpresora.tCaja AND 
                      dbo.MDOCUMENTO.tTipoDocumento = dbo.vTipodocumentoImpresora.tTipoEmision LEFT OUTER JOIN
                      dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente ON 
                      dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN
                      dbo.vCortesia ON dbo.MDOCUMENTO.tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN
                      dbo.vMozo RIGHT OUTER JOIN
                      dbo.vTipoPedido RIGHT OUTER JOIN
                      dbo.MPEDIDO ON dbo.vTipoPedido.Codigo = dbo.MPEDIDO.tTipoPedido LEFT OUTER JOIN
                      dbo.TMESA LEFT OUTER JOIN
                      dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON 
                      dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON dbo.DDOCUMENTO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido

GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create view vPaisOrigen
as

select tcodigo,ntamano,tdetallado,tresumido,lactivo
from ttabla where ttabla='PAISORIGEN'
AND  NVALOR=1


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create view vCajaCodigoControl
as

SELECT     dbo.TORIGENCODIGOCONTROL.tCaja, dbo.TCAJA.tDescripcion, dbo.TORIGENCODIGOCONTROL.fInicio, 
                      dbo.TORIGENCODIGOCONTROL.fFin, dbo.TORIGENCODIGOCONTROL.tAutorizacion,dbo.TORIGENCODIGOCONTROL.tSFC, dbo.TORIGENCODIGOCONTROL.lActivo, dbo.TORIGENCODIGOCONTROL.tDosificacion, 
                      dbo.TORIGENCODIGOCONTROL.fRegistro, dbo.TORIGENCODIGOCONTROL.tUsuario, dbo.TORIGENCODIGOCONTROL.ncorrelativo
FROM         dbo.TORIGENCODIGOCONTROL INNER JOIN
                      dbo.TCAJA ON dbo.TORIGENCODIGOCONTROL.tCaja = dbo.TCAJA.tCaja

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW vSucursal
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, tValor, lActivo  
FROM         dbo.TTABLA  
WHERE     (TTABLA = 'SUCURSAL')  

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




 
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vMaitre]
AS
SELECT     SUBSTRING(TCODIGO, 1, 4) AS Codigo, tDetallado AS Descripcion, tResumido, nBoton, tValor, lActivo, nValor, tIcono AS tBandaMagnetica, nTamano 
FROM         dbo.TTABLA
WHERE     (TTABLA = 'MAITRE')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vTienda]
AS
SELECT     tCodigoDelivery, tCodigoTienda as Codigo, tNombre as Descripcion, tDireccion, tTelefono, tEmail, tContacto, lActivo
FROM         dbo.TTIENDA
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[vGuiaTransporte]
AS
SELECT     dbo.MGUIATRANSPORTE.tGuiaTransporte, dbo.MGUIATRANSPORTE.fFecha, dbo.MGUIATRANSPORTE.nTotal, dbo.MGUIATRANSPORTE.tUsuario, 
                      dbo.MGUIATRANSPORTE.fRegistro, dbo.vCliente.Descripcion AS Destinatario, dbo.vCliente.tIdentidad AS Identidad, dbo.vCliente.tDireccion AS DireccionDestinatario, 
                      dbo.vDelivery.tDireccion AS Direccion, dbo.vTienda.Descripcion AS Tienda, dbo.vTienda.tDireccion AS DireccionTienda, vCliente_1.Descripcion AS Transportista, 
                      vCliente_1.tIdentidad AS IdentidadTransportista, vCliente_1.tDireccion AS DireccionTransportista, dbo.vVehiculo.Marca, dbo.vVehiculo.Placa, dbo.vVehiculo.Licencia, 
                      dbo.DGUIATRANSPORTE.tItem, dbo.vProducto.tResumido, dbo.DGUIATRANSPORTE.nPrecioVenta, dbo.DGUIATRANSPORTE.nCantidad, 
                      dbo.DGUIATRANSPORTE.nVenta, dbo.MGUIATRANSPORTE.tCaja, dbo.MGUIATRANSPORTE.tTurno, dbo.vEstadoGuia.Descripcion AS Estado
FROM         dbo.DGUIATRANSPORTE INNER JOIN
                      dbo.MGUIATRANSPORTE ON dbo.DGUIATRANSPORTE.tGuiaTransporte = dbo.MGUIATRANSPORTE.tGuiaTransporte INNER JOIN
                      dbo.vProducto ON dbo.DGUIATRANSPORTE.tCodigoProducto = dbo.vProducto.Codigo LEFT OUTER JOIN
                      dbo.vEstadoGuia ON dbo.MGUIATRANSPORTE.tEstadoGuia = dbo.vEstadoGuia.Codigo LEFT OUTER JOIN
                      dbo.vDelivery ON dbo.MGUIATRANSPORTE.tCodigoDelivery = dbo.vDelivery.Codigo LEFT OUTER JOIN
                      dbo.vVehiculo ON dbo.MGUIATRANSPORTE.tUnidadTransporte = dbo.vVehiculo.Codigo LEFT OUTER JOIN
                      dbo.vCliente ON dbo.MGUIATRANSPORTE.tDestinatario = dbo.vCliente.Codigo LEFT OUTER JOIN
                      dbo.vCliente AS vCliente_1 ON dbo.MGUIATRANSPORTE.tTransportista = vCliente_1.Codigo LEFT OUTER JOIN
                      dbo.vTienda ON dbo.MGUIATRANSPORTE.tCodigoDelivery = dbo.vTienda.tCodigoDelivery AND dbo.MGUIATRANSPORTE.tTienda = dbo.vTienda.Codigo CROSS JOIN
                      dbo.vEstadoPedido
GROUP BY dbo.MGUIATRANSPORTE.tGuiaTransporte, dbo.MGUIATRANSPORTE.fFecha, dbo.MGUIATRANSPORTE.nTotal, dbo.MGUIATRANSPORTE.tUsuario, 
                      dbo.MGUIATRANSPORTE.fRegistro, dbo.vCliente.Descripcion, dbo.vCliente.tIdentidad, dbo.vDelivery.tDireccion, dbo.vTienda.Descripcion, dbo.vTienda.tDireccion, 
                      vCliente_1.Descripcion, vCliente_1.tIdentidad, dbo.vVehiculo.Marca, dbo.vVehiculo.Placa, dbo.vVehiculo.Licencia, dbo.DGUIATRANSPORTE.tItem, 
                      dbo.MGUIATRANSPORTE.tCaja, dbo.MGUIATRANSPORTE.tTurno, dbo.vEstadoGuia.Descripcion, dbo.vProducto.tResumido, vCliente_1.tDireccion, 
                      dbo.vCliente.tDireccion, dbo.DGUIATRANSPORTE.nPrecioVenta, dbo.DGUIATRANSPORTE.nCantidad, dbo.DGUIATRANSPORTE.nVenta
                                           
                      
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vEstadoSolicitud]
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, TTABLA
FROM         dbo.TTABLA
WHERE     (TTABLA = 'ESTADOSOLICITUD')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[vEstadoSolicitudDetalle]
AS
SELECT     SUBSTRING(TCODIGO, 1, 2) AS Codigo, tDetallado AS Descripcion, tResumido, lActivo, TTABLA
FROM         dbo.TTABLA
WHERE     (TTABLA = 'ESTADODETSOLICITUD')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPariente  
AS  
SELECT     dbo.TPARIENTE.tCodigoPariente, dbo.TPARIENTE.tCodigoDelivery,    
 rtrim(ISNULL(dbo.TPARIENTE.tApellido, N'')) + ' ' + rtrim(ISNULL(dbo.TPARIENTE.tNombre, N'')) AS Pariente,     
                      ISNULL(dbo.vDelivery.tApellido, N'') + '  ' + ISNULL(dbo.vDelivery.tNombre, N'') AS Frecuente, ISNULL(dbo.TPARIENTE.lConyugue, 0) AS lconyugue,     
                      ISNULL(dbo.TPARIENTE.lHijo, 0) AS lHijo, dbo.vDelivery.taccionsocio  , substring(dbo.TPARIENTE.tCodigoPariente,1,5) + '-'+substring(dbo.TPARIENTE.tCodigoPariente,6,7) AS deliveryTelefono  
FROM         dbo.TPARIENTE INNER JOIN    
                      dbo.vDelivery ON dbo.TPARIENTE.tCodigoDelivery = dbo.vDelivery.Codigo    
  
go
 
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vSectorVenta
AS
SELECT        TCODIGO AS Codigo, tDetallado AS Detallado, tResumido AS Resumido, tValor AS CuentaContable, lActivo AS Activo
FROM          dbo.TTABLA
WHERE        (TTABLA = N'SECTORVENTA')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create view vSectorVentaCajaR
as
 
select tcaja.tcaja, tcaja.tdescripcion, tcaja.lactivo, vsectorventa.codigo, vsectorventa.resumido from tcaja
inner join vsectorventa
on tcaja.tsectorventa=vsectorventa.codigo
 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vAreaSubGrupo
AS

SELECT     dbo.TAREASUBGRUPO.tCaja, dbo.TAREASUBGRUPO.tArea AS tarea, dbo.TAREASUBGRUPO.tSubGrupo, dbo.TCAJA.tDescripcion AS Caja, 
                      dbo.vArea.tResumido AS Area, dbo.TSUBGRUPO.tResumido AS SubGrupo, dbo.TAREASUBGRUPO.tUsuario, dbo.TAREASUBGRUPO.fRegistro
FROM         dbo.TAREASUBGRUPO LEFT OUTER JOIN
                      dbo.TCAJA ON dbo.TAREASUBGRUPO.tCaja = dbo.TCAJA.tCaja LEFT OUTER JOIN
                      dbo.vArea ON dbo.TAREASUBGRUPO.tArea = dbo.vArea.Codigo LEFT OUTER JOIN
                      dbo.TSUBGRUPO ON dbo.TAREASUBGRUPO.tSubGrupo = dbo.TSUBGRUPO.tCodigoSubgrupo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW vAreaChef
AS
SELECT        dbo.TAREACHEF.tCaja, dbo.TAREACHEF.tArea, dbo.TCAJA.tDescripcion AS Caja, dbo.vArea.Descripcion AS Area, ISNULL(dbo.TAREACHEF.lArea, 0) AS AreaChef
FROM            dbo.TAREACHEF INNER JOIN
                         dbo.TCAJA ON dbo.TAREACHEF.tCaja = dbo.TCAJA.tCaja INNER JOIN
                         dbo.vArea ON dbo.TAREACHEF.tArea = dbo.vArea.Codigo
 
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO