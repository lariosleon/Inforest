if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_FE_factObtieneCodigoBHQ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_FE_factObtieneCodigoBHQ]
GO
create proc [dbo].[usp_FE_factObtieneCodigoBHQ]	
	@Documento		varchar(50),
	@TipoImagen		int,
	@NotaCredito	int
as
begin

	Declare @Resp  Nvarchar(max)
	declare @FEpape int

	select @FEpape=isnull(lFEpape,0) from TPARAMETRO

	if @NotaCredito	= 1 
	begin
		if @TipoImagen = 1 
			select @Resp=isnull(tImprTermica,'')  from mnotacredito where tNotaCredito = @Documento
		if @TipoImagen = 2
			select @Resp= isnull(tRespFacturacion,'')  from mnotacredito where tNotaCredito = @Documento
		if @TipoImagen = 3 
			select @Resp=isnull(tImprTermica,'')  from mnotacredito where tNotaCredito = @Documento
	end
	else
	begin
		if @TipoImagen = 1 
			select @Resp=isnull(tImprTermica,'')  from mdocumento where tdocumento = @Documento
		if @TipoImagen = 2
			select @Resp=isnull(tRespFacturacion,'')  from mdocumento where tdocumento = @Documento
		if @TipoImagen = 3 
			select @Resp=isnull(tImprTermica,'') from mdocumento where tdocumento = @Documento
	end	

	If isnull(@Resp,'')='' or @FEpape=0
		Begin
			SET NOCOUNT ON 
			declare @cadena nvarchar(4000)
			declare @nitEmisor nvarchar(12)
			declare @TIPODOCU nvarchar(12)
			declare @IGV nvarchar(12)
			declare @numFactura nvarchar(12)
			declare @serie as nvarchar(12)
			declare @caja as nvarchar(10)
			declare @fechaEmision as nvarchar(10)
			declare @nTotal as  money
			declare @codDocCli as nvarchar(17)
			declare @nitReceptor nvarchar(12)
			declare @NCredito as integer
			set @NCredito =ISNULL((select lNotaCredito from ttipodocumento where tCodigoTipoDocumento=(select tTipoDocumento from MNOTACREDITO where tNotaCredito=@Documento)),0)
			set @nitEmisor=(select ISNULL(tIdentificacionTributaria,'') from TPARAMETRO)
			set @serie = RIGHT(@Documento,8)
			print @NCredito
			print @nitEmisor
			print @serie



			if @NCredito=0 -- no es nota de credito
					begin
						set @nitReceptor=(SELECT     ISNULL(dbo.TCLIENTE.tIdentidad, N'0') AS CLIENTE 
										FROM dbo.MDOCUMENTO LEFT OUTER JOIN  dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente and dbo.MDOCUMENTO.tDocumento = @Documento
										WHERE (dbo.MDOCUMENTO.tDocumento = @Documento))

						set @codDocCli=(SELECT   ISNULL(vt.tvalor, N'0') AS CLIENTE 
							FROM dbo.MDOCUMENTO LEFT OUTER JOIN dbo.TCLIENTE ON dbo.MDOCUMENTO.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente left outer join vTipoIdentidad  vt 
							on vt.codigo= dbo.TCLIENTE.ttipoidentidad and dbo.MDOCUMENTO.tDocumento = @Documento
							WHERE (dbo.MDOCUMENTO.tDocumento = @Documento))

						select @fechaEmision=replace(convert(nvarchar(10),fRegistro,104),'.','/'),
							   @TIPODOCU=(select isnull(tCodigoFacturacion,'0') from TTIPODOCUMENTO tt where tt.tcodigotipodocumento=ttipodocumento),
							   @IGV=ROUND(NPRECIOIMPUESTO1,2),
							   @ntotal=ROUND(isnull(nventa,0),2)--,
       						 from MDOCUMENTO where tDocumento=@Documento
							set @numFactura=left(@Documento,1) +RIGHT(left(@Documento,6) ,3)
					end 

				else
				
					begin 
						set @nitReceptor=(select ISNULL(t.tIdentidad, N'0') AS CLIENTE from MNOTACREDITO mn 
											left outer join MDOCUMENTO md on mn.tDocumento=md.tDocumento and mn.tNotaCredito = @Documento
											LEFT OUTER JOIN  dbo.TCLIENTE t ON md.tCodigoCliente = t.tCodigoCliente
											WHERE (mn.tNotaCredito = @Documento))

						set @codDocCli=(select  ISNULL(vt.tvalor, N'0') AS CLIENTE
											from MNOTACREDITO mn left outer join MDOCUMENTO md on mn.tDocumento=md.tDocumento and mn.tNotaCredito = @Documento
											LEFT OUTER JOIN  dbo.TCLIENTE t ON md.tCodigoCliente = t.tCodigoCliente
											left outer join vTipoIdentidad  vt on vt.codigo= t.ttipoidentidad
											WHERE (mn.tNotaCredito = @Documento))

						select	@fechaEmision=replace(convert(nvarchar(10),fRegistro,104),'.','/'),
								@TIPODOCU=(select isnull(tCodigoFacturacion,'0') from TTIPODOCUMENTO tt where tt.tcodigotipodocumento=ttipodocumento),
								@IGV=ROUND(nImpuesto1,2),
								@ntotal=ROUND(isnull(nventa,0),2)--,
						 from MNOTACREDITO where tNotaCredito=@Documento
						set @caja=(select isnull(tCaja,'') from MNOTACREDITO where tNotaCredito=@Documento)
						set @numFactura=(select tPrefijoEnlace from TTIPODOCUMENTOIMPRESORA where tCaja=@caja and tTipoEmision=(SELECT tTipoDocumento FROM MNOTACREDITO WHERE tNotaCredito=@documento) ) +RIGHT(left(@documento,6) ,3)
					end 


			select  SUBSTRING(@nitEmisor,1,12) + '|'+-- ruc del emisior del recibo
					@TIPODOCU + '|'+-- tipodocumentosunatFE
					@numfactura+'|' + -- serie de factura
					@serie+'|' + -- correlativo
					@IGV+'|' + --- igv
					CAST(@nTotal as nvarchar(12))+'|'+ --- totla
					@fechaEmision+'|' + --- fecha emisiomn
					SUBSTRING(@codDocCli,1,17) + '|'+ --- codigo cliente
					substring(@nitReceptor,1,12) --- ruc del cliente receptor
					as codigoQr

			end
		Else
			Begin
				Select @Resp 
			end
		

end
