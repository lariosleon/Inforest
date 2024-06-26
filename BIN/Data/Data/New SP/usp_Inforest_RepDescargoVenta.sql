if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_RepInforest_DescargoVenta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_RepInforest_DescargoVenta]
GO

create Procedure [dbo].[usp_RepInforest_DescargoVenta]
@Almacen   nvarchar(max),
@fechaIni  nvarchar(max),
@fechaFin  nvarchar(max),
@sTemporal nvarchar(max),
@Local     nvarchar(max),
@Grupo     nvarchar(50),
@SubGrupo  nvarchar(50),
@Insumo    nvarchar(50),
@Area      nvarchar(50),
@Descargo  nvarchar(50),
@tipooper  integer
as
set dateformat ymd
set nocount on
Begin
	declare @Isql nvarchar(MAX)
	Declare @Comilla as nvarchar(max)
	Declare @filtro as nvarchar(max)
	set @comilla = '''' 
	 CREATE TABLE #DBTRANS (tCodigoPedido nVarChar(50) collate Modern_Spanish_CI_AS, Fecha datetime ,PlatoVenta nVarChar(7) collate Modern_Spanish_CI_AS, CantidadPlato float, 
							Item nVarChar(3) collate Modern_Spanish_CI_AS, tDescargo nVarChar(2) collate Modern_Spanish_CI_AS, Enlace nVarChar(50) collate Modern_Spanish_CI_AS, 
							TipoPedido nVarChar(50) collate Modern_Spanish_CI_AS, CodigoProducto nVarChar(50) collate Modern_Spanish_CI_AS, CantidadReceta Float,
							SubAreaAlm nVarchar(50) collate Modern_Spanish_CI_AS, SubareaInf nVarChar(50) collate Modern_Spanish_CI_AS, RecetaBase bit , lProducto bit, 
							tCodigoUnicoEtiqueta nVarchar(50) collate Modern_Spanish_CI_AS, tDocumento nVarchar(50) collate Modern_Spanish_CI_AS, fDiaContable datetime, lTransferido bit, EnlaceDes nVarChar(400) collate Modern_Spanish_CI_AS)
	 CREATE TABLE #DBTRANS2 (tCodigoPedido nVarChar(50) collate Modern_Spanish_CI_AS, Fecha datetime ,PlatoVenta nVarChar(7) collate Modern_Spanish_CI_AS, CantidadPlato float, 
							Item nVarChar(3) collate Modern_Spanish_CI_AS, tDescargo nVarChar(2) collate Modern_Spanish_CI_AS, Enlace nVarChar(50) collate Modern_Spanish_CI_AS, 
							TipoPedido nVarChar(50) collate Modern_Spanish_CI_AS, CodigoProducto nVarChar(50) collate Modern_Spanish_CI_AS, CantidadReceta Float,
							SubAreaAlm nVarchar(50) collate Modern_Spanish_CI_AS, SubareaInf nVarChar(50) collate Modern_Spanish_CI_AS, RecetaBase bit , lProducto bit, 
							tCodigoUnicoEtiqueta nVarchar(50) collate Modern_Spanish_CI_AS, tDocumento nVarchar(50) collate Modern_Spanish_CI_AS, fDiaContable datetime, lTransferido bit, EnlaceDes nVarChar(400) collate Modern_Spanish_CI_AS)

	set @filtro=''
	if @Grupo <> ''
		set @filtro = @filtro + ' and TPI.tGrupo= ' +@Comilla+@Grupo+@Comilla
	if @SubGrupo <> ''
		set @filtro = @filtro + ' and TPI.tSubGrupo= ' +@Comilla+@SubGrupo+@Comilla
	if @Insumo <>''
		set @filtro = @filtro + ' and TEMP.CodigoProducto= ' +@Comilla+@Insumo+@Comilla
	if @Area <>''
		set @filtro = @filtro + ' and TEMP.SubAreaAlm= ' +@Comilla+@Area+@Comilla
	IF @Descargo='D'
		set @filtro = @filtro + ' and (select isnull(ltransferido,0) from dpedido where tcodigopedido=TEMP.tcodigopedido and titem=TEMP.item) = 1 '
	IF @Descargo='ND'
		set @filtro = @filtro + ' and (select isnull(ltransferido,0) from dpedido where tcodigopedido=TEMP.tcodigopedido and titem=TEMP.item) = 0 '


		-------------------------------------------------------------------------Venta por Receta Venta-----------------------------------------------------------------------------------------------------
		set @Isql = 'insert into #DBTRANS ' 
		set @Isql = @Isql +' SELECT T1.tCodigoPedido, T1.fFecha, T1.tCodigoProducto AS Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, ' 
		set @Isql = @Isql + @Almacen+'.dbo.DRECETAVENTA.tCodigoProducto, ' + @Almacen + '.dbo.DRECETAVENTA.nCantidad AS nRecetaCantidad, ' 
		set @Isql = @Isql + @Almacen+'.dbo.DRECETAVENTA.tSubArea, ISNULL(T1.tSubAlmacen,'+@comilla+@comilla+'), ' 
		set @Isql = @Isql + @Almacen+ '.dbo.TPRODUCTO.lRecetaBase, '+@Almacen+'.dbo.DRECETAVENTA.lProducto, T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable, T1.lTransferido,  '+@Almacen+'.dbo.MRECETAVENTA.tDescripcion ' 

		set @Isql = @Isql +' From ( ' 
		set @Isql = @Isql +'  SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha,' 
		set @Isql = @Isql +'  dbo.DPEDIDO.tCodigoProducto, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.MPEDIDO.tTipoPedido ,'
		set @Isql = @Isql +'  dbo.DPEDIDO.tSubalmacen, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla+@comilla+') as tCodigoUnicoEtiqueta,'
		set @Isql = @Isql +'  ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla+@comilla+') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla+@comilla+') as fDiaContable,  ISNULL(dbo.DPEDIDO.lTransferido, 0) lTransferido '
		set @Isql = @Isql +'  FROM dbo.TCANALVENTA INNER JOIN dbo.DPEDIDO ON dbo.TCANALVENTA.tCodigoCanalVenta = dbo.DPEDIDO.tTipoPedido LEFT OUTER JOIN '
		set @Isql = @Isql +'  dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto RIGHT OUTER JOIN ' 
		set @Isql = @Isql +'  dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido '
		set @Isql = @Isql +'  WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +'  (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') '
		set @Isql = @Isql +'  AND (dbo.MPEDIDO.fFecha<= '+@comilla+@fechaFin+@comilla+') AND ' 
		set @Isql = @Isql +'  (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=0) OR '
		set @Isql = @Isql +'  ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +'  (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +'  (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0))'
		set @Isql = @Isql +'  ) T1 '

		set @Isql = @Isql +'  INNER JOIN '+@Almacen+'.dbo.MRECETAVENTA ON T1.tEnlace = '+@Almacen+'.dbo.MRECETAVENTA.tRecetaVenta INNER JOIN '+@Almacen+'.dbo.DRECETAVENTA ON '
		set @Isql = @Isql +'  '+@Almacen+'.dbo.MRECETAVENTA.tRecetaVenta = '+@Almacen+'.dbo.DRECETAVENTA.tRecetaVenta AND '+@Almacen+'.dbo.MRECETAVENTA.tLocal = '+@Almacen+'.dbo.DRECETAVENTA.tLocal '
		set @Isql = @Isql +'  INNER JOIN '+@Almacen+'.dbo.TPRODUCTO ON '+@Almacen+'.dbo.TPRODUCTO.tCodigoProducto = '+@Almacen+'.dbo.DRECETAVENTA.tCodigoProducto '
		set @Isql = @Isql +'  WHERE ('+@Almacen+'.dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '+@comilla+'01'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal1 = 1) or '
		set @Isql = @Isql +'  (T1.tTipoPedido = '+@comilla+'02'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '+@comilla+'03'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal3= 1) or '
		set @Isql = @Isql +'  (T1.tTipoPedido = '+@comilla+'04'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal4 = 1) or (T1.tTipoPedido = '+@comilla+'05'+@comilla+' AND '+@Almacen+'.dbo.DRECETAVENTA.lCanal5 = 1)) and '
		set @Isql = @Isql +'  '+@Almacen+'.dbo.MRECETAVENTA.tLocal='+@comilla+@Local+@comilla+' And '+@Almacen+'.dbo.MRECETAVENTA.lActivo = 1 '
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		----------------------------------------------------------------------------------------- Venta por Descargo Directo ------------------------------------------------------------------------------------------
		set @Isql = 'insert into #DBTRANS '
		set @Isql = @Isql +' SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, dbo.DPEDIDO.tCodigoProducto AS Plato, '
		set @Isql = @Isql +' dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TPRODUCTO.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea,'
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tSubalmacen,'+@comilla+@comilla+'), 0, 0, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla+@comilla+'), ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla+@comilla+') as tDocumento, '
		set @Isql = @Isql +' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla+@comilla+') as fDiaContable, ISNULL(dbo.DPEDIDO.lTransferido, 0), '+@comilla+@comilla+' '

		set @Isql = @Isql +' FROM dbo.TCANALVENTA INNER JOIN dbo.DPEDIDO ON dbo.TCANALVENTA.tCodigoCanalVenta = dbo.DPEDIDO.tTipoPedido LEFT OUTER JOIN dbo.TPRODUCTO LEFT OUTER JOIN '
		set @Isql = @Isql +' dbo.vArea ON dbo.TPRODUCTO.tArea = dbo.vArea.Codigo ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido '
		set @Isql = @Isql +' WHERE  (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.TPRODUCTO.tDescargo = '+@comilla+'D'+@comilla+') AND (ISNULL(dbo.TPRODUCTO.tEnlace, N'+@comilla+@comilla+') <> '+@comilla+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and'
		set @Isql = @Isql +' dbo.TPRODUCTO.lCombinacion=0) OR  ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR'
		set @Isql = @Isql +' dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND  (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND '
		set @Isql = @Isql +' (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.TPRODUCTO.tDescargo = '+@comilla+'D'+@comilla+') AND (ISNULL(dbo.TPRODUCTO.tEnlace, N'+@comilla+@comilla+') <> '+@comilla+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0)) '
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		------------------------------------------------------------------------------Quita los Sin de las Ventas-----------------------------------------------------------------------------------------------------------------
		set @Isql = ' delete from #DBTRANS where tCodigoPedido + Item + Platoventa + CodigoProducto in '
		set @Isql = @Isql +' (SELECT     dbo.TPRODUCTOPROPIEDAD.tCodigoPedido + dbo.TPRODUCTOPROPIEDAD.tItem + dbo.TPRODUCTOPROPIEDAD.tProducto + dbo.TPRODUCTOPROPIEDAD.tEnlace '
		set @Isql = @Isql +' FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN dbo.MPEDIDO ON dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.DPEDIDO ON '
		set @Isql = @Isql +' dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.TPRODUCTOPROPIEDAD.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TCANALVENTA ON '
		set @Isql = @Isql +' dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta '
		set @Isql = @Isql +' WHERE (((dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = '+@comilla+'9999'+@comilla+') AND (dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR '
		set @Isql = @Isql +' dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0)) OR '
		set @Isql = @Isql +' ((dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = '+@comilla+'9999'+@comilla+') AND'
		set @Isql = @Isql +' (dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND'
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1)))) '
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		---------------------------------------------------------------------------------------------Propiedades con Receta-------------------------------------------------------------------------------------------------
		set @Isql = 'insert into #DBTRANS '
		set @Isql = @Isql +' SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, ' +@Almacen+'.dbo.DRECETAPROPIEDAD.tCodigoProducto, '
		set @Isql = @Isql + @Almacen+'.dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, ' +@Almacen+'.dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,'+@comilla+@comilla+'), '
		set @Isql = @Isql + @Almacen+'.dbo.tProducto.lRecetaBase , ' +@Almacen+'.dbo.DRECETAPROPIEDAD.lProducto, T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable , T1.lTransferido,  '+@Almacen+'.dbo.MRECETAPROPIEDAD.tDescripcion '
		set @Isql = @Isql + ' From (SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido,(CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql + ' dbo.TPRODUCTOPROPIEDAD.tProducto AS Plato, dbo.DPEDIDO.nCantidad * isnull(tproductopropiedad.ncantidad,1) ncantidad, dbo.DPEDIDO.tItem, '
		set @Isql = @Isql + ' (CASE WHEN LEN(dbo.TPRODUCTOPROPIEDAD.tEnlace)= 5 THEN '+@comilla+'R'+@comilla+' ELSE '+@comilla+'D'+@comilla+' END) AS tDescargo, dbo.TPRODUCTOPROPIEDAD.tEnlace, '
		set @Isql = @Isql + ' dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubalmacen, '
		set @Isql = @Isql + ' ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+') as tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, '
		set @Isql = @Isql + ' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable ,ISNULL(dbo.DPEDIDO.lTransferido, 0) lTransferido  '
		set @Isql = @Isql + ' FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPRODUCTOPROPIEDAD ON '
		set @Isql = @Isql + ' dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.TPRODUCTO ON '
		set @Isql = @Isql + ' dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPRODUCTO.tCodigoProducto '
		set @Isql = @Isql + ' INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta WHERE '
		set @Isql = @Isql + ' (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') '
		set @Isql = @Isql + ' AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.MPEDIDO.fFecha >='+@comilla+@fechaIni+@comilla+') '
		set @Isql = @Isql + ' AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and '
		set @Isql = @Isql + ' dbo.TPRODUCTO.lCombinacion=0) OR  ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+')'
		set @Isql = @Isql + ' AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.MPEDIDO.fProgramacion >='+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql + ' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql + ' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0))) T1  '
		set @Isql = @Isql + ' INNER JOIN  '+@Almacen+'.dbo.MRECETAPROPIEDAD ON T1.tEnlace = '+@Almacen+'.dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN  '
		set @Isql = @Isql + @Almacen+'.dbo.DRECETAPROPIEDAD ON '+@Almacen+'.dbo.MRECETAPROPIEDAD.tRecetaPropiedad = '+@Almacen+'.dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND '
		set @Isql = @Isql + @Almacen+'.dbo.MRECETAPROPIEDAD.tLocal = '+@Almacen+'.dbo.DRECETAPROPIEDAD.tLocal '
		set @Isql = @Isql + ' INNER JOIN '+@Almacen+'.dbo.TPRODUCTO ON '+@Almacen+'.dbo.DRECETAPROPIEDAD.tCodigoProducto = '+@Almacen+'.dbo.TPRODUCTO.tCodigoProducto '
		set @Isql = @Isql + ' WHERE ('+@Almacen+'.dbo.DRECETAPROPIEDAD.lDescargo = 1) and '+@Almacen+'.dbo.MRECETAPROPIEDAD.tLocal='+@comilla+@Local+@comilla
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		--------------------------------------------------------------------------------------Propiedades con Descargo Directo-------------------------------------------------------------------------------------------------
		set @Isql = 'insert into #DBTRANS ' 
		set @Isql = @Isql +' SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido,(CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.DPEDIDO.tCodigoProducto AS Plato, dbo.DPEDIDO.nCantidad * isnull(tproductopropiedad.ncantidad,1) NCANTIDAD, dbo.TPRODUCTOPROPIEDAD.tItem, '+@comilla+'D'+@comilla+' AS tDescargo, '
		set @Isql = @Isql +' dbo.TPRODUCTOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TPRODUCTOPROPIEDAD.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tSubalmacen, N'+@comilla++@comilla+'), 0, 0, ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+'), ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, '
		set @Isql = @Isql +' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable, ISNULL(dbo.DPEDIDO.lTransferido, 0) , '+@comilla++@comilla+' '
		set @Isql = @Isql +' FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPROPIEDAD INNER JOIN '
		set @Isql = @Isql +' dbo.TPRODUCTOPROPIEDAD ON dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = dbo.TPRODUCTOPROPIEDAD.tProducto '
		set @Isql = @Isql +' ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.vArea ON '
		set @Isql = @Isql +' dbo.TPROPIEDAD.tArea = dbo.vArea.Codigo INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta inner join dbo.TPRODUCTO on '
		set @Isql = @Isql +' dbo.TPRODUCTO.tCodigoProducto=dbo.DPEDIDO.tCodigoProducto '
		set @Isql = @Isql +' WHERE  (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and '
		set @Isql = @Isql +' dbo.TPRODUCTO.lCombinacion=0) OR ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+')'
		set @Isql = @Isql +' AND (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=0)) '
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		--------------------------------------------------------------------------------------Combos por Recetas-------------------------------------------------------------------------------------------------
		set @Isql = 'insert into #DBTRANS '
		set @Isql = @Isql +' SELECT T1.tCodigoPedido, T1.fFecha, T1.tCodigoProducto AS Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAVENTA.tCodigoProducto, '+@Almacen+'.dbo.DRECETAVENTA.nCantidad AS nRecetaCantidad, '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAVENTA.tSubArea, ISNULL(T1.tSubAlmacen,'+@comilla++@comilla+'), '+@Almacen+'.dbo.TPRODUCTO.lRecetaBase, '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAVENTA.lProducto,T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable, T1.lTransferido , '+@Almacen+'.dbo.MRECETAVENTA.tdescripcion  '
		set @Isql = @Isql +' From (SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.CPEDIDO.tProductoCombo AS tCodigoProducto, dbo.CPEDIDO.nCantidad, dbo.DPEDIDO.tItem, TPRODUCTO_1.tDescargo, TPRODUCTO_1.tEnlace, dbo.MPEDIDO.tTipoPedido , dbo.DPEDIDO.tSubalmacen, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+') as tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, '
		set @Isql = @Isql +' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable , ISNULL(dbo.DPEDIDO.lTransferido, 0) lTransferido'
		set @Isql = @Isql +' FROM dbo.CPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TPRODUCTO ON '
		set @Isql = @Isql +' dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto INNER JOIN dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto INNER JOIN '
		set @Isql = @Isql +' dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND'
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
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
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		--------------------------------------------------------------------------------------Combo por Descargo Directo-------------------------------------------------------------------------------------------------------
		set @Isql = 'insert into #DBTRANS '
		set @Isql = @Isql +' SELECT dbo.DPEDIDO.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad, dbo.DPEDIDO.tItem, TPRODUCTO_1.tDescargo, TPRODUCTO_1.tEnlace, dbo.MPEDIDO.tTipoPedido, '
		set @Isql = @Isql +' TPRODUCTO_1.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea, ISNULL(dbo.DPEDIDO.tSubalmacen, N'+@comilla++@comilla+') , 0 , 0, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+'), ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable , ISNULL(dbo.DPEDIDO.lTransferido, 0), '+@comilla++@comilla+' '
		set @Isql = @Isql +' FROM dbo.CPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN '
		set @Isql = @Isql +' dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto INNER JOIN dbo.TPRODUCTO AS TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto INNER JOIN ' 
		set @Isql = @Isql +' dbo.vArea ON TPRODUCTO_1.tArea = dbo.vArea.Codigo INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta RIGHT OUTER JOIN '
		set @Isql = @Isql +' dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (TPRODUCTO_1.tDescargo = '+@comilla+'D'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha  >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.DPEDIDO.tEstadoItem = '+@comilla+'N'+@comilla+') AND (TPRODUCTO_1.tDescargo = '+@comilla+'D'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1)) '

		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		--------------------------------------------------------------------------------------Quita los Sin de los Combos-------------------------------------------------------------------------------------------------------
		set @Isql =' delete from #DBTRANS where tcodigoPedido + Item + Platoventa + CodigoProducto '
		set @Isql = @Isql +' in (SELECT     dbo.TCOMBOPROPIEDAD.tCodigoPedido + dbo.TCOMBOPROPIEDAD.tItem + dbo.TCOMBOPROPIEDAD.tProducto + dbo.TCOMBOPROPIEDAD.tEnlace FROM dbo.TCOMBOPROPIEDAD INNER JOIN dbo.MPEDIDO ON '
		set @Isql = @Isql +' dbo.TCOMBOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido CROSS JOIN dbo.TCANALVENTA '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = '+@comilla+'9999'+@comilla+') AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0)) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR '
		set @Isql = @Isql +' dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND (dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = '+@comilla+'9999'+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1))))'
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		----------------------------------------------------------------------------Propiedades de los Combos con Recetas-------------------------------------------------------------------------------------------------------
		set @Isql = 'insert into #DBTRANS ' 
		set @Isql = @Isql +' SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, '+@Almacen+'.dbo.DRECETAPROPIEDAD.tCodigoProducto, '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, '+@Almacen+'.dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,'+@comilla++@comilla+'),  '--0, 0,
		set @Isql = @Isql +' '+@Almacen+'.dbo.tProducto.lRecetaBase , ' +@Almacen+'.dbo.DRECETAPROPIEDAD.lProducto, '
		set @Isql = @Isql +' T1.tCodigoUnicoEtiqueta, T1.tDocumento, T1.fDiaContable, ISNULL(T1.lTransferido, 0), '+@Almacen+'.dbo.MRECETAPROPIEDAD.tdescripcion   '
		set @Isql = @Isql +' FROM (SELECT     dbo.TCOMBOPROPIEDAD.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad * isnull(tcombopropiedad.ncantidad,1) ncantidad, dbo.TCOMBOPROPIEDAD.tItem, '
		set @Isql = @Isql +' (CASE WHEN LEN(dbo.TCOMBOPROPIEDAD.tEnlace) = 5 THEN '+@comilla+'R'+@comilla+' ELSE '+@comilla+'D'+@comilla+' END) AS tDescargo, dbo.TCOMBOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubalmacen, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+') as  tCodigoUnicoEtiqueta, ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable, ISNULL(dbo.DPEDIDO.lTransferido, 0) lTransferido  '
		set @Isql = @Isql +' FROM dbo.TPRODUCTO INNER JOIN dbo.CPEDIDO INNER JOIN dbo.TCOMBOPROPIEDAD ON dbo.CPEDIDO.tCodigoPedido = dbo.TCOMBOPROPIEDAD.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.TCOMBOPROPIEDAD.tItem AND '
		set @Isql = @Isql +' dbo.CPEDIDO.tItemCombo = dbo.TCOMBOPROPIEDAD.tItemCombo INNER JOIN dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido ON'
		set @Isql = @Isql +'  dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto INNER JOIN '
		set @Isql = @Isql +' dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' or dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' or dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and dbo.TPRODUCTO.lCombinacion=1) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' or dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' or dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1))) T1 '
		set @Isql = @Isql +' INNER JOIN '+@Almacen+'.dbo.MRECETAPROPIEDAD ON T1.tEnlace = '+@Almacen+'.dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN '+@Almacen+'.dbo.DRECETAPROPIEDAD ON '
		set @Isql = @Isql +' '+@Almacen+'.dbo.MRECETAPROPIEDAD.tRecetaPropiedad = '+@Almacen+'.dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND '
		set @Isql = @Isql +' '+@Almacen+'.dbo.MRECETAPROPIEDAD.tLocal = '+@Almacen+'.dbo.DRECETAPROPIEDAD.tLocal '
		set @Isql = @Isql +'  INNER JOIN '+@Almacen+'.dbo.TPRODUCTO ON '
		set @Isql = @Isql +' '+@Almacen+'.dbo.DRECETAPROPIEDAD.tCodigoProducto = '+@Almacen+'.dbo.TPRODUCTO.tCodigoProducto '
		set @Isql = @Isql +' WHERE ('+@Almacen+'.dbo.DRECETAPROPIEDAD.lDescargo = 1) and '+@Almacen+'.dbo.MRECETAPROPIEDAD.tLocal='+@comilla+@Local+@comilla
		EXECUTE sp_executesql @Isql
		-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

		-----------------------------------------------------------------------Propiedades de los Combos con Descargo Directo--------------------------------------------------------------------------------------------------
		set @Isql = 'insert into #DBTRANS ' 
		set @Isql = @Isql +' SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, (CASE WHEN dbo.TCANALVENTA.lCanalCentralPedidos = 0 THEN dbo.MPEDIDO.fFecha ELSE dbo.MPEDIDO.fProgramacion END) AS fFecha, '
		set @Isql = @Isql +' dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad * isnull(TCOMBOPROPIEDAD.ncantidad,1) ncantidad, dbo.TCOMBOPROPIEDAD.tItem, '+@comilla+'D'+@comilla+' AS tDescargo, '
		set @Isql = @Isql +' dbo.TCOMBOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TCOMBOPROPIEDAD.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.TPROPIEDAD.tArea AS tSubArea, '
		set @Isql = @Isql +' ISNULL(dbo.DPEDIDO.tSubalmacen, N'+@comilla++@comilla+'), 0, 0 ,ISNULL(dbo.DPEDIDO.tCodigoEtiqueta,'+@comilla++@comilla+'), ISNULL(dbo.DPEDIDO.tDocumento,'+@comilla++@comilla+') as tDocumento, '
		set @Isql = @Isql +' ISNULL(dbo.MPEDIDO.fDiaContable,'+@comilla++@comilla+') as fDiaContable,  ISNULL(dbo.DPEDIDO.lTransferido, 0), '+@comilla++@comilla+'  '
		set @Isql = @Isql +' FROM dbo.CPEDIDO INNER JOIN dbo.TCOMBOPROPIEDAD ON dbo.CPEDIDO.tCodigoPedido = dbo.TCOMBOPROPIEDAD.tCodigoPedido AND '
		set @Isql = @Isql +' dbo.CPEDIDO.tItem = dbo.TCOMBOPROPIEDAD.tItem AND dbo.CPEDIDO.tItemCombo = dbo.TCOMBOPROPIEDAD.tItemCombo INNER JOIN dbo.TPROPIEDAD ON '
		set @Isql = @Isql +' dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto INNER JOIN '
		set @Isql = @Isql +' dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.MPEDIDO ON '
		set @Isql = @Isql +' dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TCANALVENTA ON dbo.DPEDIDO.tTipoPedido = dbo.TCANALVENTA.tCodigoCanalVenta '
		set @Isql = @Isql +' INNER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto '
		set @Isql = @Isql +' WHERE (((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fFecha >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fFecha <= '+@comilla+@fechaFin+@comilla+') AND (dbo.TCANALVENTA.lCanalCentralPedidos = 0) and '
		set @Isql = @Isql +' dbo.TPRODUCTO.lCombinacion=1) OR '
		set @Isql = @Isql +' ((dbo.MPEDIDO.tEstadoPedido = '+@comilla+'02'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'04'+@comilla+' OR dbo.MPEDIDO.tEstadoPedido = '+@comilla+'05'+@comilla+') AND '
		set @Isql = @Isql +' (ISNULL(dbo.DPEDIDO.lTransferido, 0) <> 2) AND (LEN(dbo.TPROPIEDAD.tEnlace) = 7) AND (dbo.MPEDIDO.fProgramacion >= '+@comilla+@fechaIni+@comilla+') AND '
		set @Isql = @Isql +' (dbo.MPEDIDO.fProgramacion <= '+@comilla+@fechaFin+@comilla+') AND (ISNULL(dbo.MPEDIDO.lEntregado,0)=1) AND '
		set @Isql = @Isql +' (dbo.TCANALVENTA.lCanalCentralPedidos = 1) and dbo.TPRODUCTO.lCombinacion=1))'
		EXECUTE sp_executesql @Isql

		insert into #DBTRANS2 select *  from #DBTRANS where recetabase=1 and lproducto=0
		delete from #DBTRANS where recetabase=1 and lproducto=0
		
			WHILE (select count(*) from #DBTRANS2)>0
			BEGIN 
				set @Isql = 'insert Into #DBTRANS '
				set @Isql = @Isql +' select T.tCodigopedido,T.Fecha,T.PlatoVenta,T.CantidadPlato,T.Item,T.tDescargo,T.Enlace,T.TipoPedido, RB.tCodigoProducto, '
				set @Isql = @Isql +' ((T.CantidadReceta/T.Factor))*(RB.ncantidadReceta) as CantidadReceta,RB.tSubarea as SubareaAlm,T.SubareaInf,RB.lRecetaBase as RecetaBase,RB.lProducto,T.tCodigoUnicoEtiqueta, '
				set @Isql = @Isql +' T.tdocumento,T.fDiaContable,T.lTransferido, t.enlacedes from  '
				set @Isql = @Isql +' (select *,(select isnull(nfactor,1) from '+@Almacen+'.dbo.tproducto where tCodigoProducto= #DBTRANS2.CodigoProducto) as Factor from #DBTRANS2 '
				set @Isql = @Isql +' ) T Inner Join ( '
				set @Isql = @Isql +' SELECT MRB.tCodigoProducto as CProducto,RB.tRecetaBase, RB.tSubArea, RB.tCodigoProducto, isnull(PR.lRecetaBase, 0) as lRecetaBase, isnull(RB.lProducto, 0) as lProducto,  '
				set @Isql = @Isql +' isnull(RB.nCantidad, 0) as nCantidadReceta,  '
				set @Isql = @Isql +' isnull(PR.nFactor, 0) as nFactor, ISNULL(PR.nPrecioPromedio, 0) AS nPrecioPromedio  '
				set @Isql = @Isql +' FROM '+@Almacen+'.dbo.DRECETABASE RB INNER JOIN '+@Almacen+'.dbo.TPRODUCTO PR ON RB.tCodigoProducto = PR.tCodigoProducto INNER JOIN '+@Almacen+'.dbo.MRECETABASE MRB '
				set @Isql = @Isql +' ON RB.tRecetaBase = MRB.tRecetaBase '
				set @Isql = @Isql +' WHERE  (RB.lDescargo=1)) RB   '
				set @Isql = @Isql +' on T.codigoproducto=RB.CProducto  '
				EXECUTE sp_executesql @Isql

				Delete from #DBTRANS2
				insert into #DBTRANS2 select *  from #DBTRANS where recetabase=1 and lproducto=0
				delete from #DBTRANS where recetabase=1 and lproducto=0

				CONTINUE  
			END

if @tipooper=1 -- LLENADO DE LA INFORMACION DE TEMPORAL PARA DESCARGO
	begin 
		set @Isql =  ' SELECT TEMP.tCodigoPedido, TEMP.Item, TPI.Grupo, TPI.SubGrupo, TPI.TipoProducto, TPI.Descripcion as PlatoVenta,  TEMP.CantidadPlato CantVenta, convert(date,TEMP.Fecha) as Fecha,  TEMP.PlatoVenta CodPlatoVenta, '
		set @Isql =  @Isql +' case TEMP.tDescargo when '+@comilla+'R'+@comilla+' then '+@comilla+'RECETA'+@comilla+' '
		set @Isql =  @Isql +' when  '+@comilla+'D'+@comilla+' then '+@comilla+'DIRECTO'+@comilla+' '
		set @Isql =  @Isql +' when  '+@comilla+'M'+@comilla+' then '+@comilla+'MENU'+@comilla+' '
		set @Isql =  @Isql +' when  '+@comilla+'N'+@comilla+' then '+@comilla+'SIN DESCARGO'+@comilla+' '
		set @Isql =  @Isql +' else '+@comilla+'SIN DESCARGO'+@comilla+' end  as Descargo,  TEMP.Enlace, '
		set @Isql =  @Isql +' case TEMP.tDescargo when '+@comilla+'R'+@comilla+' then Temp.enlacedes '
		set @Isql =  @Isql +' when  '+@comilla+'D'+@comilla+' then (select tresumido from '+@Almacen+'.dbo.vproducto WHERE tcodigoproducto= TEMP.Enlace ) '
		set @Isql =  @Isql +' when  '+@comilla+'M'+@comilla+' then '+@comilla+'MENU'+@comilla+' '
		set @Isql =  @Isql +' when  '+@comilla+'N'+@comilla+' then '+@comilla+'SIN DESCARGO'+@comilla+' '
		set @Isql =  @Isql +' else '+@comilla+'SIN DESCARGO'+@comilla+' end  as DescEnlace,  '
		set @Isql =  @Isql +' TipoPedido, TEMP.CodigoProducto,  TP.tResumido as Descripcion , '
		set @Isql =  @Isql +' CantidadReceta/nfactor as CantidadReceta, nPrecioPromedio, TEMP.CantidadPlato*(CantidadReceta/nfactor) as TotalConsumo, case when (select isnull(ltransferido,0) from dpedido where tcodigopedido=TEMP.tcodigopedido and titem=TEMP.item) = 1 then '+@comilla+'DESCARGADO'+@comilla+' else '+@comilla+'NO DESCARGADO'+@comilla+' end as Descarga, TEMP.CantidadPlato*(CantidadReceta/nfactor)*nPrecioPromedio as CostoTotal , TEMP.tDocumento, '
		set @Isql =  @Isql +' (select descripcion from '+@Almacen+'.dbo.varea where codigo= case when isnull(TEMP.SubAreaInf,'+@comilla+ @comilla+')='+@comilla+ @comilla+' then TEMP.SubAreaAlm else TEMP.SubAreaInf end ) as DescAreaAlm, ( select  nStockActual  from '+@Almacen+'.dbo.tsubstock where tcodigoproducto= TEMP.CodigoProducto  and tcodigosubarea= case when isnull(TEMP.SubAreaInf,'+@comilla+ @comilla+')='+@comilla+ @comilla+' then TEMP.SubAreaAlm else TEMP.SubAreaInf end ) as Stock  '
		set @Isql =  @Isql +' FROM #DBTRANS TEMP INNER JOIN '
		set @Isql =  @Isql +' '+@Almacen+'.dbo.TPRODUCTO TP ON TEMP.CodigoProducto = TP.tCodigoProducto'
		set @Isql =  @Isql +' LEFT JOIN VPRODUCTO TPI ON TEMP.PlatoVenta = TPI.CODIGO'
		set @Isql =  @Isql +' LEFT JOIN mDocumento on '
		set @Isql =  @Isql +' TEMP.tDocumento = mDocumento.tDocumento  where TEMP.tCodigoPedido <> '+@comilla+'00000000'+@comilla+ @filtro+' ORDER BY 1, 2,14 '
		EXECUTE sp_executesql @Isql

	end
if @tipooper=2 -- Reporte Resumido
	begin

		set @Isql =  ' SELECT TEMP.tCodigoPedido,convert(date,TEMP.Fecha) as Fecha, '
		--TEMP.Item, TPI.Grupo, TPI.SubGrupo, TPI.TipoProducto, TPI.Descripcion as PlatoVenta,  TEMP.CantidadPlato CantVenta, convert(date,TEMP.Fecha) as Fecha,  TEMP.PlatoVenta CodPlatoVenta, '
		--set @Isql =  @Isql +' case TEMP.tDescargo when '+@comilla+'R'+@comilla+' then '+@comilla+'RECETA'+@comilla+' '
		--set @Isql =  @Isql +' when  '+@comilla+'D'+@comilla+' then '+@comilla+'DIRECTO'+@comilla+' '
		--set @Isql =  @Isql +' when  '+@comilla+'M'+@comilla+' then '+@comilla+'MENU'+@comilla+' '
		--set @Isql =  @Isql +' when  '+@comilla+'N'+@comilla+' then '+@comilla+'SIN DESCARGO'+@comilla+' '
		--set @Isql =  @Isql +' else '+@comilla+'SIN DESCARGO'+@comilla+' end  as Descargo,  TEMP.Enlace, '
		--set @Isql =  @Isql +' case TEMP.tDescargo when '+@comilla+'R'+@comilla+' then Temp.enlacedes '
		--set @Isql =  @Isql +' when  '+@comilla+'D'+@comilla+' then (select tresumido from '+@Almacen+'.dbo.vproducto WHERE tcodigoproducto= TEMP.Enlace ) '
		--set @Isql =  @Isql +' when  '+@comilla+'M'+@comilla+' then '+@comilla+'MENU'+@comilla+' '
		--set @Isql =  @Isql +' when  '+@comilla+'N'+@comilla+' then '+@comilla+'SIN DESCARGO'+@comilla+' '
		--set @Isql =  @Isql +' else '+@comilla+'SIN DESCARGO'+@comilla+' end  as DescEnlace,  '
		set @Isql =  @Isql +'  TEMP.CodigoProducto,  TP.tResumido as Descripcion , '
		set @Isql =  @Isql +' sum(CantidadReceta/nfactor) as CantidadReceta, nPrecioPromedio, sum(TEMP.CantidadPlato*(CantidadReceta/nfactor)) as TotalConsumo, sum(TEMP.CantidadPlato*(CantidadReceta/nfactor)*nPrecioPromedio) as CostoTotal , '
		set @Isql =  @Isql +' (select descripcion from '+@Almacen+'.dbo.varea where codigo=TEMP.SubAreaAlm) as DescAreaAlm, ( select  nStockActual  from '+@Almacen+'.dbo.tsubstock where tcodigoproducto= TEMP.CodigoProducto  and tcodigosubarea= case when isnull(TEMP.SubAreaInf,'+@comilla+ @comilla+')='+@comilla+ @comilla+' then TEMP.SubAreaAlm else TEMP.SubAreaInf end ) as Stock,( select  nStockActual  from '+@Almacen+'.dbo.tsubstock where tcodigoproducto= TEMP.CodigoProducto  and tcodigosubarea= case when isnull(TEMP.SubAreaInf,'+@comilla+ @comilla+')='+@comilla+ @comilla+' then TEMP.SubAreaAlm else TEMP.SubAreaInf end ) - sum(TEMP.CantidadPlato*(CantidadReceta/nfactor)) as Diferencia  '
		set @Isql =  @Isql +' FROM #DBTRANS TEMP INNER JOIN '
		set @Isql =  @Isql +' '+@Almacen+'.dbo.TPRODUCTO TP ON TEMP.CodigoProducto = TP.tCodigoProducto'
		set @Isql =  @Isql +' LEFT JOIN VPRODUCTO TPI ON TEMP.PlatoVenta = TPI.CODIGO'
		set @Isql =  @Isql +' LEFT JOIN mDocumento on '
		set @Isql =  @Isql +' TEMP.tDocumento = mDocumento.tDocumento  where TEMP.tCodigoPedido <> '+@comilla+'00000000'+@comilla+ @filtro+' '
		set @Isql =  @Isql +' Group by TEMP.tCodigoPedido, convert(date,TEMP.Fecha),TEMP.CodigoProducto, TP.tResumido, nPrecioPromedio, TEMP.SubAreaAlm,TEMP.SubAreaInf order by 1,2 '
		EXECUTE sp_executesql @Isql

	end
end 

