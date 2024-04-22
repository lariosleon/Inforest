if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Inforest_Impresion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Inforest_Impresion]
GO

Create PROCEDURE [dbo].[usp_Inforest_Impresion]
@Codigo nvarchar(50),
@tipooper int
as
BEGIN
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
	DECLARE @val as integer
	declare @Mont as float
	select @val=isnull(ttipodocumentoimpresora.lImpDocMayorCero,0) from mdocumento inner join ttipodocumentoimpresora 
	on mdocumento.ttipodocumento=ttipodocumentoimpresora.ttipoemision and mdocumento.tcaja=ttipodocumentoimpresora.tcaja
	where mdocumento.tdocumento=@Codigo
----------------------------------------------------------------------------------------------------------------------------------------------------------------------
	DECLARE @valNC as integer
	declare @MontNC as float

	select @valNC=isnull(ttipodocumentoimpresora.lImpDocMayorCero,0) from mnotacredito inner join ttipodocumentoimpresora 
	on mnotacredito.ttipodocumento=ttipodocumentoimpresora.ttipoemision and mnotacredito.tcaja=ttipodocumentoimpresora.tcaja
	where mnotacredito.tnotacredito=@Codigo
----------------------------------------------------------------------------------------------------------------------------------------------------------------------
	declare @lFEBiz as Integer
	select @lFEBiz = isnull(lfebiz,0) from tparametro 
----------------------------------------------------------------------------
	if isnull(@val,0)=1
		BEGIN
			SET @Mont=0.01
		END
	ELSE
		BEGIN
			SET @Mont=0
		END

	if isnull(@valNC,0)=1
		BEGIN
			SET @MontNC=0.01
		END
	ELSE
		BEGIN
			SET @MontNC=0
		END
------------------------------------------------------- ESTO ESTA BAJO lImprimeAlternativa=false  EN EL INFOREST -----------------------------------------------------------------------
	if @tipooper=1 -- DATOS PARA LA IMPRESION DE DOCUMENTO AGRUPADO SIN FACTURACION ELECTRONICA
		BEGIN
			SELECT tDocumento, fRegistro, round(nVenta,2)-(round(nPrecioImpuesto1,2)+round(nPrecioImpuesto2,2)+round(nPrecioImpuesto3,2)) as nNeto, round(nRecargo,2) as nRecargo, round(nDescuento,2) as nDescuento , round(nPrecioImpuesto1,2) as nPrecioImpuesto1, 
			round(nPrecioImpuesto2,2) as nPrecioImpuesto2, round(nPrecioImpuesto3,2) as nPrecioImpuesto3, round(nVenta,2) as nVenta, round(nPrecioOficial,2)as nPrecioOficial, tCaja,
            tUsuario,  lImpuesto1,  lImpuesto2,  lImpuesto3, Cortesia, Cliente, RUC, Direccion, tCodigoPedido, tObservacion, tMesa, Orden, tTipoPedido, Mozo, 
            TipoPedido, MAX(tItem) As tItem, round(nPrecioNeto,2) as nPrecioNeto, round(nPrecioVenta,2)as nPrecioVenta, round(SUM(nCantidad),2) As nCantidad, round(SUM(Venta),2) AS Venta, tCodigoProducto , Mesa, Producto, 
			Combo, Descuento, Propiedad, nImpuesto1, nImpuesto2, nImpuesto3,ProductoDetallado,lOpGravInaf ,Vuelto, round(nDescuentoNeto,2) as nDescuentoNeto, tCortesia,lImpProdDesc,sum(ncantidad*DescUnitario) as DescUnitario,
			isnull(Descuento,'') as MotivoDescuento,lAnticipo
            From vDocumentoImpresoraAgrupado  where tDocumento = @Codigo 
            Group by tDocumento, fRegistro, nNeto, nRecargo, nDescuento, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta, nPrecioOficial, tCaja,
            tUsuario, lImpuesto1, lImpuesto2, lImpuesto3, Cortesia, Cliente, RUC, Direccion, tCodigoPedido, tObservacion, tMesa, Orden, tTipoPedido, Mozo, 
            TipoPedido, nPrecioNeto, nPrecioVenta, tCodigoProducto, Mesa, Producto, Combo, Descuento, Propiedad, nImpuesto1, nImpuesto2, nImpuesto3,ProductoDetallado,lOpGravInaf,Vuelto,nDescuentoNeto,tCortesia,lImpProdDesc,lAnticipo  having   round(SUM(Venta),2)>= @Mont order by tItem
		END

	if @tipooper=2 -- DATOS PARA LA IMPRESION DE DOCUMENTO AGRUPADO CON FACTURACION ELECTRONICA
		BEGIN
			select tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,ISNULL(RUC,'') As RUC,SUM(ncantidad) As ncantidad,
            producto,nprecioventa,nPrecioOficial,(nPrecioOficial-nprecioventa) As descuento,SUM(venta) As venta, round(nVenta,2) - (round(nPrecioImpuesto1,2)+round(nPrecioImpuesto2,2)) as nNeto,round(nPrecioImpuesto1,2) as nPrecioImpuesto1,
            round(nPrecioImpuesto2,2) as nPrecioImpuesto2,round(nVenta,2) as nVenta,nDescuento,MAX(tItem) As tItem,Mesa,Mozo,
            (SELECT SUM(D.nPrecioNeto*nCantidad) FROM dDocumento D WHERE D.tDocumento = @Codigo AND ((D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2=0) OR (D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2>0))) As Gravada,
            (SELECT ISNULL(SUM(D.nPrecioNeto*nCantidad),0) FROM dDocumento D WHERE ((D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2>0) OR(D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2<=0)) AND D.tDocumento = @Codigo ) As Inafecta,ProductoDetallado,lOpGravInaf,Vuelto, round(nDescuentoNeto,2) as nDescuentoNeto, tCortesia ,lImpProdDesc,sum(ncantidad*DescUnitario) as DescUnitario, isnull(Descuento,'') as MotivoDescuento,lAnticipo
            from vDocumentoImpresoraAgrupado where tdocumento=@Codigo --AND round(SUM(Venta),2)>= @Mont
            Group By tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,RUC, producto,nprecioventa,nPrecioOficial,nNeto,nPrecioImpuesto1,nPrecioImpuesto2,nVenta,nDescuento,Mesa,Mozo,ProductoDetallado,lOpGravInaf, Vuelto, nDescuentoNeto, 
			tCortesia,lImpProdDesc,Descuento,lAnticipo having   round(SUM(Venta),2)>= @Mont  order by tItem
		END

	if @tipooper=3 -- DATOS PARA LA IMPRESION DE DOCUMENTO DESAGRUPADO SIN FACTURACION ELECTRONICA
		BEGIN
			select tDocumento,tCodigoPedido,fRegistro,round(nVenta,2)-(round(nPrecioImpuesto1,2)+round(nPrecioImpuesto2,2)+round(nPrecioImpuesto3,2)) as nNeto,nRecargo,nDescuento,round(nPrecioImpuesto1,2) as nPrecioImpuesto1,round(nPrecioImpuesto2,2) as nPrecioImpuesto2,round(nPrecioImpuesto3,2) as nPrecioImpuesto3,round(nVenta,2) as nVenta,Cortesia,Cliente,isnull(RUC,'') as RUC,
			Direccion,tItem,nPrecioNeto,nPrecioVenta,nCantidad,Mozo,Mesa,tObservacion,tMesa,Producto,lModificable,nPrecioOficial,Venta,Orden,tTipoPedido,
			TipoPedido,lImpuesto1,lImpuesto2,lImpuesto3, tCodigoProducto,tCaja,tUsuario,Combo,nCombo,nImpuesto1,nImpuesto2,nImpuesto3,tObservacionPedido,
			tItemCombo,tObservacionCombo,tOferta,Descuento,ProductoDetallado,lOpGravInaf, '' as propiedad,Vuelto, round(nDescuentoNeto,2) as nDescuentoNeto, tCortesia,lImpProdDesc ,ncantidad*DescUnitario as DescUnitario, isnull(Descuento,'') as MotivoDescuento,lAnticipo
			from vDocumentoImpresora where tDocumento = @Codigo and round(Venta,2)>= @Mont  order by tItem
		END

	if @tipooper=4 -- DATOS PARA LA IMPRESION DE DOCUMENTO DESAGRUPADO CON FACTURACION ELECTRONICA
		BEGIN
			select tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,ISNULL(RUC,'') As RUC,ncantidad,producto,nprecioventa,nPrecioOficial,(nPrecioOficial-nprecioventa) As descuento,round(venta,2) as venta, round(nVenta,2) -(round(nPrecioImpuesto1,2)+round(nPrecioImpuesto2,2))as nNeto,round(nPrecioImpuesto1,2) as nPrecioImpuesto1, 
            round(nPrecioImpuesto2,2) as nPrecioImpuesto2,round(nVenta,2) as nVenta,nDescuento,tItem,Mesa,Mozo,
            (SELECT SUM(D.nPrecioNeto*nCantidad) FROM dDocumento D WHERE D.tDocumento = @Codigo AND ((D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2=0) OR (D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2>0))) As Gravada,
            (SELECT ISNULL(SUM(D.nPrecioNeto*nCantidad),0) FROM dDocumento D WHERE ((D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2>0) OR(D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2<=0)) AND D.tDocumento = @Codigo) As Inafecta,ProductoDetallado,lOpGravInaf,Vuelto, round(nDescuentoNeto,2) as nDescuentoNeto, tCortesia,lImpProdDesc ,sum(ncantidad*DescUnitario) as DescUnitario, isnull(Descuento,'') as MotivoDescuento, lAnticipo
            from vDocumentoImpresora where tdocumento=@Codigo --AND round(SUM(Venta),2)>= @Mont
			group by tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,RUC,ncantidad,producto,nprecioventa,nPrecioOficial,Mesa,Mozo,venta,nNeto,nPrecioImpuesto1, nPrecioImpuesto2,nVenta,nDescuento,titem,ProductoDetallado,lOpGravInaf,Vuelto,nDescuentoNeto, tCortesia,lImpProdDesc,Descuento, lAnticipo  having   round(SUM(Venta),2)>= @Mont 
			order by tItem
		END

----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------- ESTO ESTA BAJO lImprimeAlternativa=true  EN EL INFOREST -----------------------------------------------------------------------

	if @tipooper=5 -- DATOS PARA LA IMPRESION DE DOCUMENTO AGRUPADO SIN FACTURACION ELECTRONICA
		BEGIN
			SELECT tDocumento, fRegistro, round(nVenta,2) - (round(nPrecioImpuesto1,2)+round(nPrecioImpuesto2,2)+round(nPrecioImpuesto3,2)) as nNeto, nRecargo, nDescuento, round(nPrecioImpuesto1,2) as nPrecioImpuesto1, round(nPrecioImpuesto2,2) as nPrecioImpuesto2, round(nPrecioImpuesto3,2) as nPrecioImpuesto3, round(nVenta,2) as nVenta, nPrecioOficial, tCaja,
			tUsuario, lImpuesto1, lImpuesto2, lImpuesto3, Cortesia, Cliente, RUC, Direccion, tCodigoPedido, tObservacion, tMesa, Orden, tTipoPedido, Mozo, 
			TipoPedido, MAX(tItem) As tItem, nPrecioNeto, nPrecioVenta, SUM(nCantidad) As nCantidad, SUM(Venta) AS Venta,  tCodigoProducto, Mesa, Producto, Combo, Descuento, Propiedad, nImpuesto1, nImpuesto2, nImpuesto3,ProductoDetallado,lOpGravInaf ,Vuelto, round(nDescuentoNeto,2) as nDescuentoNeto, tCortesia ,lImpProdDesc, sum(ncantidad*DescUnitario) as DescUnitario, isnull(Descuento,'') as MotivoDescuento, lAnticipo
			From vDocumentoImpresoraAgrupadoAlternativa where tDocumento =@Codigo --AND round(SUM(Venta),2)>= @Mont
			Group by tDocumento, fRegistro, nNeto, nRecargo, nDescuento, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta, nPrecioOficial, tCaja,
			tUsuario, lImpuesto1, lImpuesto2, lImpuesto3, Cortesia, Cliente, RUC, Direccion, tCodigoPedido, tObservacion, tMesa, Orden, tTipoPedido, Mozo, 
			TipoPedido, nPrecioNeto, nPrecioVenta, tCodigoProducto, Mesa, Producto, Combo, Descuento, Propiedad, nImpuesto1, nImpuesto2, nImpuesto3,ProductoDetallado,lOpGravInaf, Vuelto, nDescuentoNeto, tCortesia, lImpProdDesc, Descuento, lAnticipo having   round(SUM(Venta),2)>= @Mont  order by tItem
		END

	if @tipooper=6 -- DATOS PARA LA IMPRESION DE DOCUMENTO AGRUPADO CON FACTURACION ELECTRONICA
		BEGIN
			select tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,ISNULL(RUC,'') As RUC,SUM(ncantidad) As ncantidad,
			producto,nprecioventa,nPrecioOficial,(nPrecioOficial-nprecioventa) As descuento,SUM(venta) As venta,round(nVenta,2) -(round(nPrecioImpuesto1,2)+round(nPrecioImpuesto2,2)) as nNeto,round(nPrecioImpuesto1,2) as nPrecioImpuesto1, 
			round(nPrecioImpuesto2,2) as nPrecioImpuesto2,round(nVenta,2) as nVenta,nDescuento,MAX(tItem) As tItem,Mesa,Mozo,
			(SELECT SUM(D.nPrecioNeto*nCantidad) FROM dDocumento D WHERE D.tDocumento = @Codigo AND ((D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2=0) OR (D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2>0))) As Gravada,
			(SELECT ISNULL(SUM(D.nPrecioNeto*nCantidad),0) FROM dDocumento D WHERE ((D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2>0) OR(D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2<=0)) AND D.tDocumento = @Codigo) As Inafecta,ProductoDetallado,lOpGravInaf, Vuelto, round(nDescuentoNeto,2) as nDescuentoNeto, tCortesia ,lImpProdDesc, sum(ncantidad*DescUnitario) as DescUnitario, isnull(Descuento,'') as MotivoDescuento, lAnticipo
			from vDocumentoImpresoraAgrupadoAlternativa where tdocumento=@Codigo --AND round(SUM(Venta),2)>= @Mont
			group By tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,RUC,producto,nprecioventa,nPrecioOficial,nNeto,nPrecioImpuesto1,nPrecioImpuesto2,nVenta,nDescuento,Mesa,Mozo,ProductoDetallado,lOpGravInaf, Vuelto, nDescuentoNeto, tCortesia,lImpProdDesc, Descuento, lAnticipo having   round(SUM(Venta),2)>= @Mont  order by tItem,lOpGravInaf
		END

	if @tipooper=7 -- DATOS PARA LA IMPRESION DE DOCUMENTO DESAGRUPADO SIN FACTURACION ELECTRONICA
		BEGIN
			select tDocumento,tCodigoPedido,fRegistro,round(nVenta,2)-(round(nPrecioImpuesto1,2)+round(nPrecioImpuesto2,2)+round(nPrecioImpuesto3,2)) as nNeto,nRecargo,nDescuento,round(nPrecioImpuesto1,2) as nPrecioImpuesto1,round(nPrecioImpuesto2,2) as nPrecioImpuesto2,round(nPrecioImpuesto3,2) as nPrecioImpuesto3,round(nVenta,2) as nVenta,Cortesia,Cliente,RUC,Direccion,tItem,
			nPrecioNeto,nPrecioVenta,nCantidad,Mozo,Mesa,tObservacion,tMesa,Producto,lModificable,nPrecioOficial,Venta,Orden,tTipoPedido,TipoPedido,lImpuesto1,lImpuesto2,lImpuesto3,case @lFEBiz when 1 then CodigoProductoSunat else tCodigoProducto end as tCodigoProducto,
			tCaja,tUsuario,Combo,nCombo,nImpuesto1,nImpuesto2,nImpuesto3,tObservacionPedido,tItemCombo,tObservacionCombo,tOferta,ProductoDetallado,lOpGravInaf, '' as propiedad ,Vuelto, round(nDescuentoNeto,2) as nDescuentoNeto, tCortesia ,lImpProdDesc, ncantidad*DescUnitario as DescUnitario,'' as MotivoDescuento, lAnticipo
			from vDocumentoImpresoraAlternativa where tDocumento = @Codigo AND round(Venta,2)>= @Mont order by tItem
		END

	if @tipooper=8 -- DATOS PARA LA IMPRESION DE DOCUMENTO DESAGRUPADO CON FACTURACION ELECTRONICA
		BEGIN
			select tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,ISNULL(RUC,'') As RUC,ncantidad,producto,nprecioventa,nPrecioOficial,(nPrecioOficial-nprecioventa) As descuento,
			round(venta,2) as venta,round(nVenta,2)-(round(nPrecioImpuesto1,2)+round(nPrecioImpuesto2,2)) as nNeto,round(nPrecioImpuesto1,2) as nPrecioImpuesto1,round(nPrecioImpuesto2,2) as nPrecioImpuesto2,round(nVenta,2) as nVenta,nDescuento,tItem,Mesa,Mozo,
            (SELECT SUM(D.nPrecioNeto*nCantidad) FROM dDocumento D WHERE D.tDocumento = @Codigo AND ((D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2=0) OR (D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2>0))) As Gravada,
            (SELECT ISNULL(SUM(D.nPrecioNeto*nCantidad),0) FROM dDocumento D WHERE ((D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2>0) OR(D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2<=0)) AND D.tDocumento = @Codigo) As Inafecta,ProductoDetallado,lOpGravInaf,Vuelto, round(nDescuentoNeto,2) as nDescuentoNeto, tCortesia ,lImpProdDesc, ncantidad*DescUnitario as DescUnitario,'' as MotivoDescuento, lAnticipo
            from vDocumentoImpresoraAlternativa where tdocumento= @Codigo AND round(Venta,2)>= @Mont order by tItem
		END

	if @tipooper=9 -- Impresion de Forma de pago
		BEGIN
			select Distinct v.tResumido + ' ' + M.tResumido as pago ,  Sum(D.nMonto)  as monto , md.nvuelto as vuelto
			from DPAGODOCUMENTO D inner join vTipoPago v on d.tTipoPago = v.Codigo 
			inner join vMoneda M on M.Codigo = D.tMoneda inner join MDOCUMENTO md on d.tDocumento=md.tDocumento
			Where D.tDocumento = @Codigo 
			group by v.tResumido,  M.tResumido, md.nvuelto
		END
		
	if @tipooper=10 -- Impresion De Notas de Credito Parte 1
		BEGIN
			select tNotaCredito,tDocumento,tCodigoPedido,fRegistro,tCaja,Cliente,Direccion,ISNULL(RUC,'') As RUC,ncantidad,producto,nprecioventa,nPrecioOficial,
			(nPrecioOficial-nprecioventa) As descuento,venta,nNeto,nPrecioImpuesto1,nPrecioImpuesto2,nVenta,nDescuento,tItem,Mesa,Mozo, 
			(SELECT SUM(D.nPrecioNeto*nCantidad) FROM dDocumento D WHERE D.tDocumento = @Codigo AND ((D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2=0) OR (D.nPrecioImpuesto1>0 AND D.nPrecioImpuesto2>0))) As Gravada,
			(SELECT ISNULL(SUM(D.nPrecioNeto*nCantidad),0) FROM dDocumento D WHERE ((D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2>0) OR(D.nPrecioImpuesto1<=0 AND D.nPrecioImpuesto2<=0)) AND D.tDocumento = @Codigo) As Inafecta, 
			tObservacion,lImpProdDesc, isnull(ncantidad*DescUnitario,0) as DescUnitario,case @lFEBiz when 1 then CodigoProductoSunat else tCodigoProducto end as tCodigoProducto  from vNotaCreditoImpresora 
			where tNotaCredito=@Codigo AND round(Venta,2)>= @MontNC order by tItem
		END
	if @tipooper=11 -- Impresion De Notas de Credito Parte 2
		BEGIN
			select * from vnotacreditoimpresora where tNotaCredito=@Codigo AND round(Venta,2)>= @MontNC order by tItem
		END
END