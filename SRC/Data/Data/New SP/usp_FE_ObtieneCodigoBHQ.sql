if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_FE_ObtieneCodigoBHQ]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_FE_ObtieneCodigoBHQ]
GO
Create proc [dbo].[usp_FE_ObtieneCodigoBHQ]	
	@Documento		varchar(50),
	@TipoImagen		int,
	@NotaCredito	int
as
begin

	if @NotaCredito	= 1 
	begin
		if @TipoImagen = 1 
			select isnull(tImprTermica,'') as codigo from mnotacredito where tNotaCredito = @Documento
		if @TipoImagen = 2
			select isnull(tRespFacturacion,'') as codigo from mnotacredito where tNotaCredito = @Documento
		if @TipoImagen = 3 
			select isnull(tImprTermica,'') as codigo from mnotacredito where tNotaCredito = @Documento
	end
	else
	begin
		if @TipoImagen = 1 
			select isnull(tImprTermica,'') as codigo from mdocumento where tdocumento = @Documento
		if @TipoImagen = 2
			select isnull(tRespFacturacion,'') as codigo from mdocumento where tdocumento = @Documento
		if @TipoImagen = 3 
			select isnull(tImprTermica,'') as codigo from mdocumento where tdocumento = @Documento
	end	
end

