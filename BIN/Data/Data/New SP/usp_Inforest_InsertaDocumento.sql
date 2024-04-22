if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Inforest_InsertaDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Inforest_InsertaDocumento]
GO

create PROCEDURE [dbo].[usp_Inforest_InsertaDocumento]
@Pedido nvarchar(50),
@Documento nvarchar(50),
@tTipoDocumento nvarchar(50),
@tCodigoCliente nvarchar(50),
@tEstadoDocumento nvarchar(50),
@tCaja nvarchar(50),
@tTurno nvarchar(50),
@tSalon nvarchar(50),
@tUsuario nvarchar(50),
@tUsuarioAutoriza nvarchar(50),
@fDiaContable date,
@tDescuento nvarchar(50),
@tConsumo nvarchar(300),
@lImpresionMonedaExtranjera integer,
@tautorizacion nvarchar(50),
@tcodigocontrol nvarchar(50),
@Cortesia nvarchar(50),
@fInicio date,
@fCaducidad date,
@tContribuyenteEspecial nvarchar(50),
@tipooper int
as
BEGIN
Declare @nPrecioNeto float
Declare @nimpuesto1 float
Declare @nimpuesto2 float
Declare @nimpuesto3 float
Declare @nventa float
Declare @nDescuento float
Declare @Recargo float
declare @tReservaInf as nvarchar(50)
declare @anticipo int

	if @tipooper=1 -- INSERTA LOS DATOS DEL DOCUMENTO DE INFOREST Tickets
		BEGIN
			select @nPrecioNeto=sum(nPrecioNeto*nCantidad) , @nimpuesto1=sum(nImpuesto1) , 
			@nimpuesto2=sum(nImpuesto2) , @nimpuesto3=sum(nImpuesto3) ,@nventa= sum(nVenta),
			@nDescuento=isnull(sum(nDescuento*nCantidad),0) , @Recargo=isnull(sum(nRecargo*nCantidad),0) 
            from DPEDIDO 
			where tCodigoPedido =@Pedido and ((isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0))
			group by tCodigoPedido
			--------------Detalle del documento------------
			set @tReservaInf=(select isnull(tReservaInf,'') from mpedido where tCodigoPedido=@Pedido)

			if @tReservaInf<>''
				begin
					set @anticipo=1
				end
			else
				begin
					set @anticipo=0
				end

			Insert into 
			DDOCUMENTO 
			( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, 
			nPrecioImpuesto2, nPrecioImpuesto3,nPrecioVenta, nRecargo, nDescuento, nCantidad,
			nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta )  
			select  @Documento as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1,
			nPrecioImpuesto2, nPrecioImpuesto3,nPrecioVenta, nRecargo, nDescuento, nCantidad, 
			nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta 
			From DPEDIDO  
			where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) 
			and tCodigoPedido =@Pedido
			------------- cabecera del documento---------------
			Insert into 
			MDOCUMENTO  
			( tDocumento, tTipoDocumento, tCodigoCliente,nNeto,nRecargo ,nDescuento, nPrecioOficial,
			nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta,tEstadoDocumento, tCaja, tTurno, 
			tSalon, tUsuario, tUsuarioAutoriza, fRegistro, fDiaContable, tDescuento, tConsumo,
			lImpresionMonedaExtranjera,tAutorizacion,tCodigoControl,lreplica,fInicio, fCaducidad, tContribuyenteEspecial, tCortesia,lImpresionAut,lAnticipo)  
			Values(@Documento,@tTipoDocumento, @tCodigoCliente, @nPrecioNeto, @Recargo, @nDescuento, 0,
			@nimpuesto1, @nimpuesto2, @nimpuesto3, @nventa, @tEstadoDocumento ,@tCaja ,@tTurno,
			@tSalon,@tUsuario, @tUsuarioAutoriza, getdate(),@fDiaContable,@tDescuento,@tConsumo,
			@lImpresionMonedaExtranjera,@tautorizacion,@tcodigocontrol,1,@fInicio,@fCaducidad,@tContribuyenteEspecial,@Cortesia,1,@anticipo)
		END
	if @tipooper=2 -- INSERTA LOS DATOS DEL DOCUMENTO DE INFOREST variable
		BEGIN
			select @nPrecioNeto=sum(nPrecioNeto*nCantidad) , @nimpuesto1=sum(nImpuesto1) , 
			@nimpuesto2=sum(nImpuesto2) , @nimpuesto3=sum(nImpuesto3) ,@nventa= sum(nVenta),
			@nDescuento=isnull(sum(nDescuento*nCantidad),0) 
            from DPEDIDO 
			where tDocumento =@Documento
			group by tDocumento
			--------------Detalle del documento------------
			Insert into 
			DDOCUMENTO 
			( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, 
			nPrecioImpuesto2, nPrecioImpuesto3,nPrecioVenta, nRecargo, nDescuento, nCantidad,
			nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta )  
			select  @Documento as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1,
			nPrecioImpuesto2, nPrecioImpuesto3,nPrecioVenta, nRecargo, nDescuento, nCantidad, 
			nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta 
			From DPEDIDO  
			where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) 
			and tDocumento =@Documento
			------------- cabecera del documento---------------
			Insert into 
			MDOCUMENTO  
			( tDocumento, tTipoDocumento, tCodigoCliente,nNeto,nRecargo ,nDescuento, nPrecioOficial,
			nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta,tEstadoDocumento, tCaja, tTurno, 
			tSalon, tUsuario, tUsuarioAutoriza, fRegistro, fDiaContable, tDescuento, tConsumo,
			lImpresionMonedaExtranjera,tAutorizacion,tCodigoControl,lreplica,fInicio, fCaducidad, tContribuyenteEspecial,tCortesia)  
			Values(@Documento,@tTipoDocumento, @tCodigoCliente, @nPrecioNeto, 0, @nDescuento, 0,
			@nimpuesto1, @nimpuesto2, @nimpuesto3, @nventa, @tEstadoDocumento ,@tCaja ,@tTurno,
			@tSalon,@tUsuario, @tUsuarioAutoriza, getdate(),@fDiaContable,@tDescuento,@tConsumo,
			@lImpresionMonedaExtranjera,@tautorizacion,@tcodigocontrol,1,@fInicio,@fCaducidad,@tContribuyenteEspecial,@Cortesia)
		END

END
