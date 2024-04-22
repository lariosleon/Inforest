if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_RepLiquidacion3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_RepLiquidacion3]
GO
create procedure [dbo].[sp_RepLiquidacion3]
@finicio as datetime,
@ffin as datetime,
@turno nvarchar(50),
@tipooper int
as
begin
	if @tipooper= 1
	begin
			select dbo.MDOCUMENTO.tEstadoDocumento,
			dbo.TTIPODOCUMENTO.tCodigoTipoDocumento ,isnull(UPPER(M.Descripcion),'SIN SALON') as Salon, M.tDetallado as Mesa,
			dbo.TTIPODOCUMENTO.tDescripcion,dbo.mdocumento.tdocumento,dbo.vproducto.tgrupo,dbo.vproducto.grupo,dbo.vproducto.tSubGrupo,dbo.vproducto.SubGrupo,dbo.vProducto.Descripcion,
			sum(dbo.ddocumento.nPrecioNeto*dbo.ddocumento.ncantidad) as Neto,sum(dbo.ddocumento.nImpuesto1) as Impuesto1,sum(dbo.ddocumento.nImpuesto2) as Impuesto2,
			sum(dbo.ddocumento.nImpuesto3) as Impuesto3,sum(dbo.ddocumento.nVenta) as nVenta--, sum( isnull(dbo.DPAGODOCUMENTO.nPropina,0)) as Propina,DBO.DPAGODOCUMENTO.tTipoPago--,isnull(DBO.DPAGODOCUMENTO.nMonto,0) AS MontoPago
			INTO #DATOS
			from dbo.MDOCUMENTO left join dbo.DDOCUMENTO on dbo.MDOCUMENTO.tDocumento=dbo.DDOCUMENTO.tDocumento  
			inner join dbo.TTIPODOCUMENTO on dbo.MDOCUMENTO.tTipoDocumento=dbo.TTIPODOCUMENTO.tCodigoTipoDocumento
			inner join dbo.vProducto on dbo.DDOCUMENTO.tCodigoProducto=dbo.vProducto.Codigo
			LEFT join (select dbo.vsalon.Codigo as Codsalon,dbo.vSalon.Descripcion,dbo.TMESA.tCodigoMesa,dbo.TMESA.tDetallado from dbo.vSalon inner join dbo.TMESA on dbo.vsalon.Codigo=dbo.TMESA.tSalon ) M on 
			(select isnull(tMesa,'') from mpedido where MPEDIDO.tCodigoPedido=dbo.DDOCUMENTO.tCodigoPedido)=m.tcodigomesa
			where 	DBO.MDOCUMENTO.tTurno=@turno and dbo.MDOCUMENTO.tTipoDocumento<>'00'
			group by dbo.MDOCUMENTO.tEstadoDocumento,dbo.TTIPODOCUMENTO.tDescripcion,dbo.TTIPODOCUMENTO.tCodigoTipoDocumento,dbo.mdocumento.tdocumento,dbo.DDOCUMENTO.tCodigoPedido,
			dbo.vproducto.grupo,dbo.vproducto.tgrupo,dbo.vproducto.tSubGrupo,dbo.vproducto.SubGrupo,dbo.vProducto.Descripcion,M.Descripcion,M.tDetallado--,DBO.DPAGODOCUMENTO.tTipoPago--,DBO.DPAGODOCUMENTO.nMonto
	
			SELECT DBO.MDOCUMENTO.tTipoDocumento,DBO.DPAGODOCUMENTO.tMoneda,DBO.DPAGODOCUMENTO.tTipoPago,
			case when DBO.DPAGODOCUMENTO.tmoneda='01' then SUM(isnull(DBO.DPAGODOCUMENTO.nMonto,0)) else SUM(isnull(DBO.DPAGODOCUMENTO.nMonto*DBO.DPAGODOCUMENTO.ntipocambio,0)) end nMonto,SUM(isnull(DBO.DPAGODOCUMENTO.nPropina,0)) nPropina,
			isnull(sum((DBO.DPAGODOCUMENTO.nPropina)*(DBO.TTARJETACREDITO.nFactorRetencion/100)),0) AS fRetencion
			INTO #DATOSPAGO
			FROM  DBO.MDOCUMENTO  INNER JOIN DBO.DPAGODOCUMENTO  ON DBO.MDOCUMENTO.tDocumento=DBO.DPAGODOCUMENTO.tDocumento --AND T.tDocumento=DP.tDocumento
			left join DBO.TTARJETACREDITO on DBO.DPAGODOCUMENTO.tTarjeta=DBO.TTARJETACREDITO.tCodigoTarjeta
			WHERE DBO.MDOCUMENTO.tTurno=@turno 
			GROUP BY DBO.MDOCUMENTO.tTipoDocumento,DBO.DPAGODOCUMENTO.tMoneda,DBO.DPAGODOCUMENTO.tTipoPago,DBO.DPAGODOCUMENTO.tTarjeta

			--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			SELECT '01' AS Orden,UPPER(grupo) AS GRUPO,tdescripcion as Doc, sum(Neto) TOTAL FROM #DATOS D --right join vGrupo G on d.tGrupo=g.Codigo
			where tEstadoDocumento IN ('01','02','03')
			group by grupo,tDescripcion
			UNION
			SELECT '02'AS Orden,'SUBTOTAL'AS GRUPO, tdescripcion as Doc,CAST( sum(Neto) AS DECIMAL (10,2)) TOTAL FROM #DATOS
			where tEstadoDocumento IN ('01','02','03')
			group by tDescripcion
			UNION
			SELECT '03'AS Orden,(Select upper(tImpuesto1) from TPARAMETRO)AS GRUPO,tdescripcion as Doc, sum(Impuesto1)  TOTAL FROM #DATOS
			where tEstadoDocumento IN ('01','02','03')
			group by tDescripcion
			union
			--if (Select upper(tImpuesto2) from TPARAMETRO)<>''
			--	begin
			--	SELECT '04'AS Orden,(Select upper(tImpuesto2) from TPARAMETRO)AS GRUPO,tdescripcion as Doc, sum(Impuesto2)  TOTAL FROM #DATOS
			--	where tEstadoDocumento IN ('01','02','03')
			--	group by tDescripcion
			--	--union
			--end
			SELECT '05'AS Orden,'TOTAL'AS GRUPO,tdescripcion as Doc, sum(nVenta) TOTAL FROM #DATOS
			where tEstadoDocumento IN ('01','02','03')
			group by tDescripcion
			union
			SELECT '06'AS Orden,'PROPINAS' AS GRUPO,tdescripcion as Doc, 
			(select  sum(nPropina) from #DATOSPAGO where tTipoPago='02' and tTipoDocumento=#DATOS.tCodigoTipoDocumento  )  TOTAL FROM #DATOS  
			where tEstadoDocumento IN ('01','02','03')
			group by tDescripcion,tCodigoTipoDocumento
			union
			SELECT '07'AS Orden,'TOTAL DE INGRESOS' AS GRUPO,tdescripcion as Doc,
			isnull((select  sum(nPropina) from #DATOSPAGO where tTipoPago='02' and tTipoDocumento=#DATOS.tCodigoTipoDocumento  ),0)+sum(nVenta)  TOTAL FROM #DATOS
			where tEstadoDocumento IN ('01','02','03')
			group by tDescripcion,tCodigoTipoDocumento
			union
			SELECT '08'AS Orden,'TARJETAS DE CREDITO' AS GRUPO,tDescripcion,
			(select  sum(nmonto)+sum(nPropina) from #DATOSPAGO where tTipoPago='02' and tTipoDocumento=#DATOS.tCodigoTipoDocumento  ) AS TOTAL
			FROM #DATOS
			where tEstadoDocumento IN ('01','02','03')
			GROUP BY tDescripcion,#DATOS.tCodigoTipoDocumento
			--union
			--SELECT '09'AS Orden,'OTROS PAGOS' AS Grupo,tDescripcion,
			--(select sum(nmonto) from #DATOSPAGO where tTipoPago not in('02','01') and tTipoDocumento=#DATOS.tCodigoTipoDocumento  ) AS TOTAL
			--FROM #DATOS 
			--where tEstadoDocumento IN ('01','02','03')
			--GROUP BY tDescripcion, tCodigoTipoDocumento
			union
			SELECT '10'AS Orden,'VENTAS X COBRAR'AS Grupo,tdescripcion as Doc, 
			ISNULL((SELECT SUM(nVenta) FROM #DATOS H WHERE H.tEstadoDocumento='03' AND H.tCodigoTipoDocumento=#DATOS.tCodigoTipoDocumento ),0) FROM #DATOS
			group by #DATOS.tCodigoTipoDocumento,tDescripcion
			union
			SELECT '12'AS Orden,'VENTAS EFECTIVAS'AS Grupo,tdescripcion as Doc, 
			isnull((select  sum(nPropina) from #DATOSPAGO where tTipoPago='02' and tTipoDocumento=#DATOS.tCodigoTipoDocumento  ),0)+ISNULL(sum(nVenta),0)-ISNULL((select sum(nmonto) from #DATOSPAGO where tTipoPago not in('01') and tTipoDocumento=#DATOS.tCodigoTipoDocumento  ),0) AS  TOTAL FROM #DATOS
			where tEstadoDocumento IN ('02')
			group by tDescripcion,tCodigoTipoDocumento
			union
			SELECT '13'AS Orden,'PROPINAS NETAS' AS GRUPO,tdescripcion as Doc, 
			(select  sum(nPropina)- sum(fretencion) from #DATOSPAGO where tTipoPago='02' and tTipoDocumento=#DATOS.tCodigoTipoDocumento  )  TOTAL FROM #DATOS  
			where tEstadoDocumento IN ('01','02','03')
			group by tDescripcion,tCodigoTipoDocumento
			union
			SELECT '14'AS Orden,'EFECTIVO FISICO' AS GRUPO,tDescripcion,
			(select  sum(nmonto) from #DATOSPAGO where tTipoPago='01' and tTipoDocumento=#DATOS.tCodigoTipoDocumento  ) AS TOTAL
			FROM #DATOS
			where tEstadoDocumento IN ('01','02','03')
			GROUP BY tDescripcion,#DATOS.tCodigoTipoDocumento
			--SELECT '14' AS Orden, 'RECIBO EGRESO' as GRUPO,'RECIBO' as Doc,
			--SUM(CASE WHEN tMoneda='01'THEN ISNULL(nMonto,0) ELSE CASE WHEN tMoneda='02' THEN ISNULL((nTipoCambio*nMonto),0) END END ) AS TOTAL FROM MEGRESO 
			--WHERE tTurno=@turno and tEstadoDocumento='01'
			--union
			--SELECT '15' AS Orden, 'RECIBO INGRESO' as GRUPO,'RECIBO' as Doc,
			--SUM(CASE WHEN tMoneda='01'THEN ISNULL(nMonto,0) ELSE CASE WHEN tMoneda='02' THEN ISNULL((nTipoCambio*nMonto),0) END END ) AS TOTAL FROM MINGRESO 
			--WHERE tTurno=@turno and tEstadoDocumento='02'
			union 
			SELECT '16'AS Orden,'' AS GRUPO, tdescripcion as Doc, null as  TOTAL FROM #DATOS group by tDescripcion,tCodigoTipoDocumento
			union 
			SELECT '17'AS Orden,'----ZONAS----' AS GRUPO, tdescripcion as Doc, null as  TOTAL FROM #DATOS  group by tDescripcion,tCodigoTipoDocumento
			union
			SELECT '18'AS Orden,Salon AS GRUPO,tdescripcion as Doc, 
			(sum(nVenta) )  TOTAL FROM #DATOS  
			where tEstadoDocumento IN ('01','02','03')
			group by Salon,tDescripcion,tCodigoTipoDocumento
			union
			SELECT '19'AS Orden,'TOTAL SALONES' AS GRUPO,tdescripcion as Doc, 
			(sum(nVenta) )  TOTAL FROM #DATOS  
			where tEstadoDocumento IN ('01','02','03')
			group by Salon,tDescripcion,tCodigoTipoDocumento

		ORDER BY 1,2;

	DROP TABLE #DATOS
	DROP TABLE #DATOSPAGO
	end
	else
	begin
		SELECT '01' AS Orden,UPPER('gggg') AS GRUPO,'ffdfg' as Doc, 10.25 TOTAL --FROM #DATOS D --right join vGrupo G on d.tGrupo=g.Codigo
		--where tEstadoDocumento IN ('01','02','03')
			--group by grupo,tDescripcion
	end
end
