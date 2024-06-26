if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ListDocumentosFE]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ListDocumentosFE]
GO
Create Procedure [dbo].[usp_ListDocumentosFE]
@caja nvarchar(50),
@FechaIni as date,
@FechaFin as date,
@tipooper as int
as
begin
	if @tipooper=1 --- Lista Documentos 
		BEGIN
			if (select isnull(lfebiz,0) from tparametro)=1 --- Bizlinks
				bEGIN 
					select left(tdocumento,1) as Pref ,tdocumento as Documento,
					isnull((select Descripcion from vcliente where codigo=mdocumento.tcodigocliente),'Sin Cliente Facturado') as Cliente,
					nVenta as Monto, tusuario as Usuario, tturno as Turno,ttipodocumento as TipoDoc,'D' as DOC
					from mdocumento 
					where ttipodocumento<>'00' and 
					--testadoDOcumento<>'04' and
					ttipodocumento in (select isnull(ttipoemision,0) from vtipodocumentoimpresora where tcaja=@caja and ISNULL(lfacturacionelectronica,0)=1) and 
					isnull(lEstadoFacturacion,0)=0 and 
					 tcaja=@caja and
					 MONTH(fregistro)>month(getdate()-30)
					-- order by 2
					UNION
					select left(tdocumento,1) as Pref ,tNotaCredito as Documento,
					isnull((select Descripcion from vcliente where codigo=(Select isnull(tcodigocliente,'') from mdocumento where tdocumento=dbo.mnotacredito.tdocumento)),'Sin Cliente Facturado') as Cliente,
					nVenta as Monto, tusuario as Usuario, tturno as Turno,ttipodocumento as TipoDoc,'N' as DOC
					from MNOTACREDITO 
					where ttipodocumento<>'00' and 
					--testadoDOcumento<>'04' and
					ttipodocumento in (select isnull(ttipoemision,0) from vtipodocumentoimpresora where tcaja=@caja and ISNULL(lfacturacionelectronica,0)=1) and 
					isnull(lEstadoFacturacion,0)=0 and 
					 tcaja=@caja and
					 MONTH(fregistro)>month(getdate()-30)
				END 
				ELSE
				BEGIN ----Lista Documentos no enviado a  Paperlees
					select left(tdocumento,1) as Pref ,tdocumento as Documento,
					isnull((select Descripcion from vcliente where codigo=mdocumento.tcodigocliente),'Sin Cliente Facturado') as Cliente,
					nVenta as Monto, tusuario as Usuario, tturno as Turno,ttipodocumento as TipoDoc--,*
					from mdocumento 
					where ttipodocumento<>'00' and 
					testadoDOcumento<>'04' and
					ttipodocumento in (select isnull(ttipoemision,0) from vtipodocumentoimpresora where tcaja=@caja and ISNULL(lfacturacionelectronica,0)=1) and 
					isnull(lEstadoFacturacion,0)=0 and 
					 tcaja=@caja and
					 MONTH(fregistro)>month(getdate()-30)
					 order by 2
				END
		END
	if @tipooper=2 --- Lista de Documentos Electronicos Bislinz
		BEGIN

			select tcaja as caja,tdocumento as nro_efact ,fregistro,ttipodocumento as tipodocu, 
			isnull((select Descripcion from vcliente where codigo=dbo.mdocumento.tcodigocliente),'') as razonsocial, lestadofacturacion as cdr, 
			Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tdocumento,12),3) +'-'+RIGHT(tdocumento,8) as   numerorefe, 
			tusuario as cajero ,'D' as DOC , 'BOLETA/FACTURA' AS DocDescripcion, isnull(trespfacturacion,'') +' / '+ isnull(tImprTermica,'') as Firma
			from dbo.mdocumento 
			where ttipodocumento<>'00' and convert(date,fregistro) between convert(date,@FechaIni) 
			and convert(date,@FechaFin) and tcaja=@caja

			union

			select tcaja as caja,tnotacredito as nro_efact ,fregistro,ttipodocumento as tipodocu, 
			isnull((select Descripcion from vcliente where codigo=(Select isnull(tcodigocliente,'') from mdocumento where tdocumento=dbo.mnotacredito.tdocumento)),'') as razonsocial, lestadofacturacion as cdr, 
			Case When ISNULL(lestadofacturacion,0)  = 1 then 'Enviado' else 'No Enviado' end As cdrDes, LEFT(tdocumento,1) + LEFT(RIGHT(tnotacredito,12),3) +'-'+RIGHT(tnotacredito,8) as   numerorefe, 
			tusuario as cajero ,'N' as DOC, 'NOTA DE CREDITO' AS DocDescripcion, isnull(trespfacturacion,'') +' / '+ isnull(tImprTermica,'') as Firma
			from dbo.mnotacredito 
			where ttipodocumento<>'00' and convert(date,fregistro) between convert(date,@FechaIni) 
			and convert(date,@FechaFin) and tcaja=@caja

		END
end

