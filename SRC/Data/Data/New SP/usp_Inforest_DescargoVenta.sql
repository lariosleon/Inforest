if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Inforest_DescargoVenta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Inforest_DescargoVenta]
GO

CREATE Procedure [dbo].[usp_Inforest_DescargoVenta]
@Almacen   nvarchar(max),
@fechaIni  nvarchar(max),
@fechaFin  nvarchar(max),
@sTemporal nvarchar(max),
@Local     nvarchar(max),
@Pedido    nvarchar(max),
@tipooper  integer
as
set dateformat ymd
set nocount off
Begin
	declare @Isql nvarchar(MAX)
	Declare @Comilla as nvarchar(max)
	set @comilla = '''' 

if @tipooper=1 -- LLENADO DE LA INFORMACION DE TEMPORAL PARA DESCARGO
	begin 
		-------------------------------------------------------------------------Venta por Receta Venta-----------------------------------------------------------------------------------------------------
		set @Isql = 'insert into ' + @sTemporal 
		set @Isql = @Isql +' SELECT T1.tCodigoPedido, T1.fFecha, T1.tCodigoProducto AS Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, ' 
		set @Isql = @Isql + @Almacen+'.dbo.DRECETAVENTA.tCodigoProducto, ' + @Almacen + '.dbo.DRECETAVENTA.nCantidad AS nRecetaCantidad, ' 
		set @Isql = @Isql + @Almacen+'.dbo.DRECETAVENTA.tSubArea, ISNULL(T1.tSubAlmacen,'+@comilla+@comilla+'), ' 
		set @Isql = @Isql + @Almacen+ '.dbo.TPRODUCTO.lRecetaBase, '+@Almacen+'.dbo.DRECETAVENTA.lProducto, T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable ' 

		set @Isql = @Isql +' From ( ' 
		set @Isql = @Isql +'  SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha,' 
		set @Isql = @Isql +'  dbo.DPEDIDO.tCodigoProducto, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.MPEDIDO.tTipoPedido ,'
		set @Isql = @Isql +'  dbo.DPEDIDO.tSubalmacen, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla+@comilla+') as tCodigoUnicoEtiqueta,'
		set @Isql = @Isql +'  ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla+@comilla+') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla+@comilla+') as fDiaContable'
		set @Isql = @Isql +'  FROM dbo.TCANALVENTA INNER JOIN dbo.DPEDIDO ON dbo.TCANALVENTA.tCodigoCanalVenta = dbo.DPEDIDO.tTipoPedido LEFT OUTER JOIN '
		set @Isql = @Isql +'  dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto RIGHT OUTER JOIN ' 
		set @Isql = @Isql +'  dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido '
		set @Isql = @Isql +'  WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +'  (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') '
		set @Isql = @Isql +'  AND (dbo.MPEDIDO.fFecha<= '+@comilla+@fechaFin+@comilla+') AND ' 
		set @Isql = @Isql +'  (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=0) OR '
		set @Isql = @Isql +'  ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +'  (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +'  (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0))'
		set @Isql = @Isql +'  ) T1 '

		set @Isql = @Isql +'  INNER JOIN '+@Almacen+'.dbo.MRECETAVENTA ON T1.tEnlace = '+@Almacen+'.dbo.MRECETAVENTA.tRecetaVenta INNER JOIN '+@Almacen+'.dbo.DRECETAVENTA ON '
		set @Isql = @Isql +'  '+@Almacen+'.dbo.MRECETAVENTA.tRecetaVenta = '+@Almacen+'.dbo.DRECETAVENTA.tRecetaVenta AND '+@Almacen+'.dbo.MRECETAVENTA.tLocal = '+@Almacen+'.dbo.DRECETAVENTA.tLocal '
		set @Isql = @Isql +'  INNER JOIN '+@Almacen+'.dbo.TPRODUCTO ON '+@Almacen+'.dbo.TPRODUCTO.tCodigoProducto = '+@Almacen+'.dbo.DRECETAVENTA.tCodigoProducto '
		set @Isql = @Isql +'  WHERE ('+@Almacen+'.dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '+@comilla+'01'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal1 = 1) or '
		set @Isql = @Isql +'  (T1.tTipoPedido = '+@comilla+'02'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '+@comilla+'03'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal3= 1) or '
		set @Isql = @Isql +'  (T1.tTipoPedido = '+@comilla+'04'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal4 = 1) or (T1.tTipoPedido = '+@comilla+'05'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal5 = 1)) and '
		set @Isql = @Isql +'  '+@Almacen+'.dbo.MRECETAVENTA.tLocal='+@comilla+@Local+@comilla+' And '+@Almacen+'.dbo.MRECETAVENTA.lActivo = 1 '
		print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		----------------------------------------------------------------------------------------- Venta por Descargo Directo ------------------------------------------------------------------------------------------
		set @Isql = 'insert into ' + @sTemporal 
		set @Isql = @Isql +' SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.DPEDIDO.tCodigoProducto AS Plato, '
		set @Isql = @Isql +' dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TPRODUCTO.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea,'
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tSubalmacen,'+@comilla+@comilla+'), 0, 0, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla+@comilla+'), ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla+@comilla+') as tDocumento, '
		set @Isql = @Isql +' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla+@comilla+') as fDiaContable '

		set @Isql = @Isql +' FROM dbo.TCANALVENTA INNER JOIN dbo.DPEDIDO ON dbo.TCANALVENTA.tCodigoCanalVenta = dbo.DPEDIDO.tTipoPedido LEFT OUTER JOIN dbo.TPRODUCTO LEFT OUTER JOIN '
		set @Isql = @Isql +' dbo.vArea ON dbo.TPRODUCTO.tArea = dbo.vArea.Codigo ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido '
		set @Isql = @Isql +' WHERE  (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.TPRODUCTO.tDescargo = '+@comilla+'D'+@comilla+') AND (ISNULL(dbo.TPRODUCTO.tEnlace, N'+@comilla+@comilla+') <> '+@comilla+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and'
		set @Isql = @Isql +' dbo.TPRODUCTO.lCombinacion=0) OR  ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR'
		set @Isql = @Isql +' dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND  (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND '
		set @Isql = @Isql +' (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.TPRODUCTO.tDescargo = '+@comilla+'D'+@comilla+') AND (ISNULL(dbo.TPRODUCTO.tEnlace, N'+@comilla+@comilla+') <> '+@comilla+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0)) '
		print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		------------------------------------------------------------------------------Quita los Sin de las Ventas-----------------------------------------------------------------------------------------------------------------
		set @Isql = ' delete from ' + @sTemporal +' where tCodigoPedido + tItem + Plato + tCodigoProducto in '
		set @Isql = @Isql +' (SELECT     dbo.TPRODUCTOPROPIEDAD.tCodigoPedido + dbo.TPRODUCTOPROPIEDAD.tItem + dbo.TPRODUCTOPROPIEDAD.tProducto + dbo.TPRODUCTOPROPIEDAD.tEnlace '
		set @Isql = @Isql +' FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN dbo.MPEDIDO ON dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.DPEDIDO ON '
		set @Isql = @Isql +' dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.TPRODUCTOPROPIEDAD.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TCANALVENTA ON '
		set @Isql = @Isql +' dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta '
		set @Isql = @Isql +' WHERE (((dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = '+@comilla+'9999'+@comilla+') AND (dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR '
		set @Isql = @Isql +' dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0)) OR '
		set @Isql = @Isql +' ((dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = '+@comilla+'9999'+@comilla+') AND'
		set @Isql = @Isql +' (dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND'
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1)))) '
		print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		---------------------------------------------------------------------------------------------Propiedades con Receta-------------------------------------------------------------------------------------------------
		set @Isql = 'insert into ' + @sTemporal 
		set @Isql = @Isql +' SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, ' +@Almacen+'.dbo.DRECETAPROPIEDAD.tCodigoProducto, '
		set @Isql = @Isql + @Almacen+'.dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, ' +@Almacen+'.dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,'+@comilla+@comilla+'), '
		set @Isql = @Isql + @Almacen+'.dbo.tProducto.lRecetaBase , ' +@Almacen+'.dbo.DRECETAPROPIEDAD.lProducto, T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable '
		set @Isql = @Isql + ' From (SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido,(CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql + ' dbo.TPRODUCTOPROPIEDAD.tProducto AS Plato, dbo.DPEDIDO.nCantidad * isnull(tproductopropiedad.ncantidad,1) ncantidad, dbo.DPEDIDO.tItem, '
		set @Isql = @Isql + ' (CASE WHEN LEN(dbo.TPRODUCTOPROPIEDAD.tEnlace)= 5 THEN '+@comilla+'R'+@comilla+' ELSE '+@comilla+'D'+@comilla+' END) AS tDescargo, dbo.TPRODUCTOPROPIEDAD.tEnlace, '
		set @Isql = @Isql + ' dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubalmacen, '
		set @Isql = @Isql + ' ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+') as tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, '
		set @Isql = @Isql + ' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable '
		set @Isql = @Isql + ' FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPRODUCTOPROPIEDAD ON '
		set @Isql = @Isql + ' dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.TPRODUCTO ON '
		set @Isql = @Isql + ' dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPRODUCTO.tCodigoProducto '
		set @Isql = @Isql + ' INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta WHERE '
		set @Isql = @Isql + ' (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') '
		set @Isql = @Isql + ' AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.fFecha >='+@comilla+@fechaIni+@comilla+') '
		set @Isql = @Isql + ' AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and '
		set @Isql = @Isql + ' dbo.TPRODUCTO.lCombinacion=0) OR  ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+')'
		set @Isql = @Isql + ' AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.fProgramacion >='+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql + ' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql + ' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0))) T1  '
		set @Isql = @Isql + ' INNER JOIN  '+@Almacen+'.dbo.MRECETAPROPIEDAD ON T1.tEnlace = '+@Almacen+'.dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN  '
		set @Isql = @Isql + @Almacen+'.dbo.DRECETAPROPIEDAD ON '+@Almacen+'.dbo.MRECETAPROPIEDAD.tRecetaPropiedad = '+@Almacen+'.dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND '
		set @Isql = @Isql + @Almacen+'.dbo.MRECETAPROPIEDAD.tLocal = '+@Almacen+'.dbo.DRECETAPROPIEDAD.tLocal '
		set @Isql = @Isql + ' INNER JOIN '+@Almacen+'.dbo.TPRODUCTO ON '+@Almacen+'.dbo.DRECETAPROPIEDAD.tCodigoProducto = '+@Almacen+'.dbo.TPRODUCTO.tCodigoProducto '
		set @Isql = @Isql + ' WHERE ('+@Almacen+'.dbo.DRECETAPROPIEDAD.lDescargo = 1) and '+@Almacen+'.dbo.MRECETAPROPIEDAD.tLocal='+@comilla+@Local+@comilla
		print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		--------------------------------------------------------------------------------------Propiedades con Descargo Directo-------------------------------------------------------------------------------------------------
		set @Isql = 'insert into ' + @sTemporal 
		set @Isql = @Isql +' SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido,(CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.DPEDIDO.tCodigoProducto AS Plato, dbo.DPEDIDO.nCantidad * isnull(tproductopropiedad.ncantidad,1) NCANTIDAD, dbo.TPRODUCTOPROPIEDAD.tItem, '+@comilla+'D'+@comilla+' AS tDescargo, '
		set @Isql = @Isql +' dbo.TPRODUCTOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TPRODUCTOPROPIEDAD.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tSubalmacen, N'+@comilla++@comilla+'), 0, 0, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+'), ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, '
		set @Isql = @Isql +' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable '
		set @Isql = @Isql +' FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPROPIEDAD INNER JOIN '
		set @Isql = @Isql +' dbo.TPRODUCTOPROPIEDAD ON dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = dbo.TPRODUCTOPROPIEDAD.tProducto '
		set @Isql = @Isql +' ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.vArea ON '
		set @Isql = @Isql +' dbo.TPROPIEDAD.tArea = dbo.vArea.Codigo INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta inner join dbo.TPRODUCTO on '
		set @Isql = @Isql +' dbo.TPRODUCTO.tCodigoProducto=dbo.DPEDIDO.tCodigoProducto '
		set @Isql = @Isql +' WHERE  (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and '
		set @Isql = @Isql +' dbo.TPRODUCTO.lCombinacion=0) OR ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+')'
		set @Isql = @Isql +' AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0)) '
		print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		--------------------------------------------------------------------------------------Combos por Recetas-------------------------------------------------------------------------------------------------
		set @Isql = 'insert into ' + @sTemporal
		set @Isql = @Isql +' SELECT T1.tCodigoPedido, T1.fFecha, T1.tCodigoProducto AS Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAVENTA.tCodigoProducto, '+@Almacen+'.dbo.DRECETAVENTA.nCantidad AS nRecetaCantidad, '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAVENTA.tSubArea, ISNULL(T1.tSubAlmacen,'+@comilla++@comilla+'), '+@Almacen+'.dbo.TPRODUCTO.lRecetaBase, '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAVENTA.lProducto,T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable '
		set @Isql = @Isql +' From (SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.CPEDIDO.tProductoCombo AS tCodigoProducto, dbo.CPEDIDO.nCantidad, dbo.DPEDIDO.tItem, TPRODUCTO_1.tDescargo, TPRODUCTO_1.tEnlace, dbo.MPEDIDO.tTipoPedido , dbo.DPEDIDO.tSubalmacen, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+') as tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, '
		set @Isql = @Isql +' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable '
		set @Isql = @Isql +' FROM dbo.CPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TPRODUCTO ON '
		set @Isql = @Isql +' dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto INNER JOIN dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto INNER JOIN '
		set @Isql = @Isql +' dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND'
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1))) T1 '
		set @Isql = @Isql +' INNER JOIN '+@Almacen+'.dbo.MRECETAVENTA ON T1.tEnlace = '+@Almacen+'.dbo.MRECETAVENTA.tRecetaVenta INNER JOIN '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAVENTA ON '+@Almacen+'.dbo.MRECETAVENTA.tRecetaVenta = '+@Almacen+'.dbo.DRECETAVENTA.tRecetaVenta AND '
		set @Isql = @Isql +' '+@Almacen+'.dbo.MRECETAVENTA.tLocal = '+@Almacen+'.dbo.DRECETAVENTA.tLocal INNER JOIN '+@Almacen+'.dbo.TPRODUCTO ON '
		set @Isql = @Isql +' '+@Almacen+'.dbo.TPRODUCTO.tCodigoProducto = '+@Almacen+'.dbo.DRECETAVENTA.tCodigoProducto '
		set @Isql = @Isql +' WHERE ('+@Almacen+'.dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '+@comilla+'01'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal1 = 1) or (T1.tTipoPedido = '+@comilla+'02'+@comilla+' AND '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '+@comilla+'03'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal3= 1) or '
		set @Isql = @Isql +' (T1.tTipoPedido = '+@comilla+'04'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal4 = 1) or (T1.tTipoPedido = '+@comilla+'05'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal5 = 1)) and '
		set @Isql = @Isql +' '+@Almacen+'.dbo.MRECETAVENTA.tLocal='+@comilla+@Local+@comilla
		print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		--------------------------------------------------------------------------------------Combo por Descargo Directo-------------------------------------------------------------------------------------------------------
		set @Isql = 'insert into ' + @sTemporal
		set @Isql = @Isql +' SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad, dbo.DPEDIDO.tItem, TPRODUCTO_1.tDescargo, TPRODUCTO_1.tEnlace, dbo.MPEDIDO.tTipoPedido, '
		set @Isql = @Isql +' TPRODUCTO_1.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea, ISNULL(dbo.DPEDIDO.tSubalmacen, N'+@comilla++@comilla+') , 0 , 0, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+'), ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable '
		set @Isql = @Isql +' FROM dbo.CPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN '
		set @Isql = @Isql +' dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto INNER JOIN dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto INNER JOIN ' 
		set @Isql = @Isql +' dbo.vArea ON TPRODUCTO_1.tArea = dbo.vArea.Codigo INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta RIGHT OUTER JOIN '
		set @Isql = @Isql +' dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (TPRODUCTO_1.tDescargo = '+@comilla+'D'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha  >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (TPRODUCTO_1.tDescargo = '+@comilla+'D'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1)) '
		--print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		--------------------------------------------------------------------------------------Quita los Sin de los Combos-------------------------------------------------------------------------------------------------------
		set @Isql =' delete from '+ @sTemporal +' where tcodigoPedido + tItem + Plato + tCodigoProducto '
		set @Isql = @Isql +' in (SELECT     dbo.TCOMBOPROPIEDAD.tCodigoPedido + dbo.TCOMBOPROPIEDAD.tItem + dbo.TCOMBOPROPIEDAD.tProducto + dbo.TCOMBOPROPIEDAD.tEnlace FROM dbo.TCOMBOPROPIEDAD INNER JOIN dbo.MPEDIDO ON '
		set @Isql = @Isql +' dbo.TCOMBOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido CROSS JOIN dbo.TCANALVENTA '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = '+@comilla+'9999'+@comilla+') AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0)) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR '
		set @Isql = @Isql +' dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND (dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = '+@comilla+'9999'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1))))'
		--print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		----------------------------------------------------------------------------Propiedades de los Combos con Recetas-------------------------------------------------------------------------------------------------------
		set @Isql = 'insert into ' + @sTemporal
		set @Isql = @Isql +' SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, '+@Almacen+'.dbo.DRECETAPROPIEDAD.tCodigoProducto, '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, '+@Almacen+'.dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,'+@comilla++@comilla+'), 0, 0, '
		set @Isql = @Isql +' T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable '
		set @Isql = @Isql +' FROM (SELECT     dbo.TCOMBOPROPIEDAD.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad * isnull(tcombopropiedad.ncantidad,1) ncantidad, dbo.TCOMBOPROPIEDAD.tItem, '
		set @Isql = @Isql +' (CASE WHEN LEN(dbo.TCOMBOPROPIEDAD.tEnlace) = 5 THEN '+@comilla+'R'+@comilla+' ELSE '+@comilla+'D'+@comilla+' END) AS tDescargo, dbo.TCOMBOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubalmacen, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+') as  tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable '
		set @Isql = @Isql +' FROM dbo.TPRODUCTO INNER JOIN dbo.CPEDIDO INNER JOIN dbo.TCOMBOPROPIEDAD ON dbo.CPEDIDO.tCodigoPedido = dbo.TCOMBOPROPIEDAD.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.TCOMBOPROPIEDAD.tItem AND '
		set @Isql = @Isql +' dbo.CPEDIDO.tItemCombo = dbo.TCOMBOPROPIEDAD.tItemCombo INNER JOIN dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido ON'
		set @Isql = @Isql +'  dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto INNER JOIN '
		set @Isql = @Isql +' dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' or dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' or dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' or dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' or dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1))) T1 '
		set @Isql = @Isql +' INNER JOIN '+@Almacen+'.dbo.MRECETAPROPIEDAD ON T1.tEnlace = '+@Almacen+'.dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN '+@Almacen+'.dbo.DRECETAPROPIEDAD ON '
		set @Isql = @Isql +' '+@Almacen+'.dbo.MRECETAPROPIEDAD.tRecetaPropiedad = '+@Almacen+'.dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND '
		set @Isql = @Isql +' '+@Almacen+'.dbo.MRECETAPROPIEDAD.tLocal = '+@Almacen+'.dbo.DRECETAPROPIEDAD.tLocal '
		set @Isql = @Isql +' WHERE ('+@Almacen+'.dbo.DRECETAPROPIEDAD.lDescargo = 1) and '+@Almacen+'.dbo.MRECETAPROPIEDAD.tLocal='+@comilla+@Local+@comilla
		print @isql
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		-----------------------------------------------------------------------Propiedades de los Combos con Descargo Directo--------------------------------------------------------------------------------------------------
		set @Isql = 'insert into ' + @sTemporal
		set @Isql = @Isql +' SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad * isnull(TCOMBOPROPIEDAD.ncantidad,1) ncantidad, dbo.TCOMBOPROPIEDAD.tItem, '+@comilla+'D'+@comilla+' AS tDescargo, '
		set @Isql = @Isql +' dbo.TCOMBOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TCOMBOPROPIEDAD.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.TPROPIEDAD.tArea AS tSubArea, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tSubalmacen, N'+@comilla++@comilla+'), 0, 0 ,ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+'), ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, '
		set @Isql = @Isql +' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable '
		set @Isql = @Isql +' FROM dbo.CPEDIDO INNER JOIN dbo.TCOMBOPROPIEDAD ON dbo.CPEDIDO.tCodigoPedido = dbo.TCOMBOPROPIEDAD.tCodigoPedido AND '
		set @Isql = @Isql +' dbo.CPEDIDO.tItem = dbo.TCOMBOPROPIEDAD.tItem AND dbo.CPEDIDO.tItemCombo = dbo.TCOMBOPROPIEDAD.tItemCombo INNER JOIN dbo.TPROPIEDAD ON '
		set @Isql = @Isql +' dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto INNER JOIN '
		set @Isql = @Isql +' dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.MPEDIDO ON '
		set @Isql = @Isql +' dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta '
		set @Isql = @Isql +' INNER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and '
		set @Isql = @Isql +' dbo.TPRODUCTO.lCombinacion=1) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1))'
		print @isql
		EXECUTE sp_executesql @Isql
	end
if @tipooper=2 -- PROCESO DE DESCARGO "PROCESODESCARGO" 
	begin
		set @Isql =  ' SELECT TEMP.tCodigoPedido, fFecha, Plato, nCantidad, tItem, tDescargo, tEnlace, tTipoPedido, TEMP.tCodigoProducto, '
		set @Isql =  @Isql +' nRecetaCantidad, tSubAreaAlm, tSubAreaInf, nFactor, nPrecioPromedio, TEMP.lRecetaBase, TEMP.lProducto,TEMP.tCodigoUnicoEtiqueta, '
		set @Isql =  @Isql +' TEMP.tDocumento, IsNull(mDocumento.tTipoDocumento,'+@comilla++@comilla+') as tTipoDocumento, TEMP.fDiaContable as fDiaContable '
		set @Isql =  @Isql +' FROM ' + @sTemporal+' TEMP INNER JOIN '
		set @Isql =  @Isql +' '+@Almacen+'.dbo.TPRODUCTO TP ON TEMP.tCodigoProducto = TP.tCodigoProducto LEFT JOIN mDocumento on '
		set @Isql =  @Isql +' TEMP.tDocumento = mDocumento.tDocumento where TEMP.tCodigoPedido = '+@comilla+@Pedido+@comilla+' and TEMP.titem='+@comilla+@Local+@comilla+' order by tCodigoPedido, tItem '
		--print @isql
		EXECUTE sp_executesql @Isql 
	end
end 


