if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Inforest_InsertarLogErrores]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Inforest_InsertarLogErrores]
GO

create PROCEDURE [dbo].[usp_Inforest_InsertarLogErrores]
@Ttabla nvarchar(100),
@Proceso nvarchar(100),
@CodError nvarchar(100),
@ErrorProcedure nvarchar(100),
@ErrorLine nvarchar(100),
@ErrorMensaje nvarchar(500),
@DatoAlternativo nvarchar(100),
@Observaciones nvarchar(100),
@Usuario nvarchar(100),
@tipooper int
as
BEGIN

	if @tipooper=1 -- Insertar Log de errores de inforest
		BEGIN
			DECLARE @ID AS BIGINT
			SET @ID = (select isnull( MAX(ID),0) FROM LOG_INFOREST )+1

			insert into LOG_INFOREST (ID,Ttabla,Proceso,CodError,ErrorProcedure,ErrorLine,ErrorMensaje,DatoAlternativo,Observaciones,Usuario,fregistro,Estado) 
			values					(@ID,@Ttabla,@Proceso,@CodError,@ErrorProcedure,@ErrorLine,@ErrorMensaje,@DatoAlternativo,@Observaciones,@Usuario,GETDATE(),1)
		END

END

