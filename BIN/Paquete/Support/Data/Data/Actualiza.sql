if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MGUIA]') and OBJECTPROPERTY(id, N'IsTable') = 1)
   DROP TABLE MGUIA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DGUIA]') and OBJECTPROPERTY(id, N'IsTable') = 1)
   DROP TABLE DGUIA
GO

--Insertar Datos
delete from TTABLA where tTabla='FRECUENCIA'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','00','Diario','Diario',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','01','Lunes','Lunes',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','02','Martes','Martes',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','03','Miercoles','Miercoles',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','04','Jueves','Jueves',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','05','Viernes','Viernes',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','06','Sabados','Sabados',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','07','Domingos','Domingo',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','99','Fecha Especial','Fecha Espcial',0,1)

delete FROM TTABLA where TTABLA='MOTIVOELIMINACION' AND tCodigo='000'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('MOTIVOELIMINACION','000','Otro','Otro',0,1)

delete FROM TTABLA where TTABLA='MOTIVODESCUENTO' AND tCodigo='000'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, nValor, lActivo) values ('MOTIVODESCUENTO','000','Otro','Otro',0,0,1)

delete FROM TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='000'
delete FROM TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='001'
delete FROM TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='002'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCANCELACION','000','Otro','Otro',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCANCELACION','001','Recibo Ingreso','Recibo Ingreso',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCANCELACION','002','Nota de Crédito','Nota de Crédito',0,1)

delete FROM TTABLA where TTABLA='ESTADOGUIA'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOGUIA','01','EMITIDO','EMITIDO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOGUIA','02','FACTURADO','FACTURADO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOGUIA','03','ANULADO','ANULADO',0,1)

delete FROM TTABLA where TTABLA='ESTADOPEDIDO' AND tCodigo='04'
delete FROM TTABLA where TTABLA='ESTADOPEDIDO' AND tCodigo='05'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOPEDIDO','04','CUENTA CORRIENTE','CUENTA CORRIENTE',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOPEDIDO','05','CARGADO','CARGADO',0,1)

delete FROM TTABLA where TTABLA='TIPOPAGO' and TCODIGO='04'
delete FROM TTABLA where TTABLA='TIPOPAGO' and TCODIGO='05'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','04','Varios','Varios',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','05','Puntos','Puntos',0,1)

delete FROM TTABLA where TTABLA='TIPOCLIENTEFRECUENTE' and TCODIGO='00'
--delete FROM TTABLA where TTABLA='TIPOCLIENTEFRECUENTE' and TCODIGO='01'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCLIENTEFRECUENTE','00','SIN TIPO','SIN TIPO',0,1)
--INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCLIENTEFRECUENTE','01','CLIENTE DELIVERY','CLIENTE DELIVERY',0,1)

DELETE FROM TTABLA WHERE TTABLA='TIPOCANCELACION' and tCodigo='000'
DELETE FROM TTABLA WHERE TTABLA='TIPOCANCELACION' and tCodigo='001'
DELETE FROM TTABLA WHERE TTABLA='TIPOCANCELACION' and tCodigo='002'
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOCANCELACION', '000',2,'Otro','Otro',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOCANCELACION', '001',2,'Recibo de Ingreso','Recibo de Ingreso',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOCANCELACION', '002',2,'Nota de Credito','Nota de Credito',1)

DELETE FROM TTABLA WHERE TTABLA='TIPOUBICACION'
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '01',2,'Avenida','AV',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '02',2,'Jiron','JR',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '03',2,'Calle','CL',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '04',2,'Pasaje','PJ',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '05',2,'Alameda','AL',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '06',2,'Malecon','MA',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '07',2,'Ovalo','OV',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '08',2,'Parque','PQ',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '09',2,'Plaza','PL',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '10',2,'Carretera','CA',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '11',2,'Plazuela','PL',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '12',2,'Bajada','BA',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '13',2,'Conjunto Residencial','RS',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '14',2,'Barrio Fiscal','BF',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '15',2,'Urbanizacion','UR',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '16',2,'Urbanizacion Popular','UP',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '17',2,'Cooperativa','CO',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '18',2,'Villa Militar','VM',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '19',2,'Unidad Vecinal','UV',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '20',2,'Asociacion','AS',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '21',2,'Paseo','PS',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '22',2,'Prolongacion','PR',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '23',2,'Asentamiento Humano','AH',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '24',2,'Agrupacion','AG',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '25',2,'Otros','OT',1)

DELETE FROM TTABLA WHERE TTABLA='UNIDADNEGOCIO' AND TCODIGO='00'
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('UNIDADNEGOCIO', '00',2,'SIN UNIDAD NEGOCIO','SIN UNIDAD NEGOCIO',1)

delete from TACCESO
INSERT INTO TACCESO values ('00000001','03','Boton Platos','mdiAdministracion','BT','cmdopcion1',null,null)
INSERT INTO TACCESO values ('00000002','03','Boton Clientes','mdiAdministracion','BT','cmdopcion2',null,null)
INSERT INTO TACCESO values ('00000003','03','Boton Mesas','mdiAdministracion','BT','cmdopcion3',null,null)
INSERT INTO TACCESO values ('00000004','03','Boton Usuarios','mdiAdministracion','BT','cmdopcion4',null,null)
INSERT INTO TACCESO values ('00000005','03','Boton Backup','mdiAdministracion','BT','cmdopcion5',null,null)
INSERT INTO TACCESO values ('00000006','03','Boton Restore','mdiAdministracion','BT','cmdopcion6',null,null)
INSERT INTO TACCESO values ('00000007','04','Vistas de Mesas','mdiConsulta','BT','cmdopcion1',null,null)
INSERT INTO TACCESO values ('00000008','04','Coorelativo de Pedidos','mdiConsulta','BT','cmdopcion2',null,null)
INSERT INTO TACCESO values ('00000009','04','Coorelativo de Documentos','mdiConsulta','BT','cmdopcion3',null,null)
INSERT INTO TACCESO values ('00000010','04','Liquidacion de Cajero','mdiConsulta','BT','cmdopcion4',null,null)
INSERT INTO TACCESO values ('00000011','04','Registro de Ventas	','mdiConsulta','BT','cmdopcion5',null,null)
INSERT INTO TACCESO values ('00000012','04','Paloteo de Produccion','mdiConsulta','BT','cmdopcion6',null,null)
INSERT INTO TACCESO values ('00000013','04','Propinas','mdiConsulta','BT','cmdopcion7',null,null)
INSERT INTO TACCESO values ('00000014','02','Apertura','mdiPuntoVenta','BT','cmdopcion1',null,null)
INSERT INTO TACCESO values ('00000015','02','Punto de Venta','mdiPuntoVenta','BT','cmdopcion2',null,null)
INSERT INTO TACCESO values ('00000016','02','Cierre','mdiPuntoVenta','BT','cmdopcion3',null,null)
INSERT INTO TACCESO values ('00000017','02','Mesas','mdiPuntoVenta','BT','cmdopcion4',null,null)
INSERT INTO TACCESO values ('00000018','02','Recibo de Egresos','mdiPuntoVenta','BT','cmdopcion5',null,null)
INSERT INTO TACCESO values ('00000019','02','Recibo de Ingresos','mdiPuntoVenta','BT','cmdopcion6',null,null)
INSERT INTO TACCESO values ('00000020','02','Reservas','mdiPuntoVenta','BT','cmdopcion7',null,null)
INSERT INTO TACCESO values ('00000021','02','Correlativo de Pedidos','mdiPuntoVenta','BT','cmdopcion8',null,null)
INSERT INTO TACCESO values ('00000022','02','Correlativo de Documentos','mdiPuntoVenta','BT','cmdopcion9',null,null)
INSERT INTO TACCESO values ('00000023','02','Cuentas Corrientes','mdiPuntoVenta','BT','cmdopcion10',null,null)
INSERT INTO TACCESO values ('00000024','02','Cuentas por Cobrar','mdiPuntoVenta','BT','cmdopcion11',null,null)
INSERT INTO TACCESO values ('00000025','02','Carta de Productos','mdiPuntoVenta','BT','cmdopcion12',null,null)
INSERT INTO TACCESO values ('00000026','02','Delivery en Tránsito','mdiPuntoVenta','BT','cmdopcion13',null,null)
INSERT INTO TACCESO values ('00000027','02','Delivery Entregados','mdiPuntoVenta','BT','cmdopcion14',null,null)

INSERT INTO TACCESO values ('10100000','03','Configuración','mdiAdministracion','MN','mnuConfiguracion',null,null)
INSERT INTO TACCESO values ('10110000','03','Parámetros Generales','mdiAdministracion','MN','mnuParametro',NULL,12)
INSERT INTO TACCESO values ('10120000','03','Canal de Venta','mdiAdministracion','MN','mnuTipoPedido',NULL,null)
INSERT INTO TACCESO values ('10130000','03','Mantenimientos de Impresora','mdiAdministracion','MN','mnuImpresora',null,6)
INSERT INTO TACCESO values ('10140000','03','Configuración de Cajas','mdiAdministracion','MN','mnuConfiguraCaja',null,2)
INSERT INTO TACCESO values ('10150000','03','Configuración de Impresoras','mdiAdministracion','MN','mnuConfiguracionImpresora',null,null)
INSERT INTO TACCESO values ('10160000','03','Configura Mensaje','mdiAdministracion','MN','mnuMensaje','TMENSAJE',7)
INSERT INTO TACCESO values ('10161000','03','Mantenimiento de Establecimientos (Locales)','mdiAdministracion','MN','mnuManteEstable',null,null)
INSERT INTO TACCESO values ('10162000','03','Mantenimiento de Agrupacion de Puntos de Ventas','mdiAdministracion','MN','mnuMantSectorVentas',null,null)
INSERT INTO TACCESO values ('10170000','03','Cierre de Periodo','mdiAdministracion','MN','mnuCierre',null,36)
INSERT INTO TACCESO values ('10190000','03','Actualización de Tablas','mdiAdministracion','MN','mnuTablaReplica',null,null)

--CESAR
INSERT INTO TACCESO values ('10175000','03','Usuarios','mdiAdministracion','MN','mnuUsuario','TUSUARIO',18)
INSERT INTO TACCESO values ('10175010','03','Opción Agregar','frmUsuario','MN','cmdopcion0','TUSUARIO',null)
INSERT INTO TACCESO values ('10175020','03','Opción Modificar','frmUsuario','MN','cmdopcion1','TUSUARIO',null)
INSERT INTO TACCESO values ('10175030','03','Opción Eliminar','frmUsuario','MN','cmdOpcion2','TUSUARIO',null)

INSERT INTO TACCESO values ('10180000','03','Grupo de Usuarios','mdiAdministracion','MN','mnuGrupoUsuario','GRUPOUSUARIO',1)
INSERT INTO TACCESO values ('10185000','03','Tipos de Cambio','mdiAdministracion','MN','mnuTipoCambio','TTIPOCAMBIO',17)
INSERT INTO TACCESO values ('10200000','03','Tablas','mdiAdministracion','MN','mnuTabla',null,null)
INSERT INTO TACCESO values ('10201000','03','Tipos de Identidad','mdiAdministracion','MN','mnuIdentidad','TTABLA-TIPOIDENTIDAD',29)
INSERT INTO TACCESO values ('10202000','03','Tipos de Documento','mdiAdministracion','MN','mnuTipoDocumento','TTABLA-TIPODOCUMENTO',28)
INSERT INTO TACCESO values ('10203000','03','Tipos de Clientes Frecuentes','mdiAdministracion','MN','mnuTipocliente',null,27)
INSERT INTO TACCESO values ('10204000','03','Tipos de Cuenta Corriente','mdiAdministracion','MN','mnuTipoCtaCte',null,25)
INSERT INTO TACCESO values ('10205000','03','Cuentas Contables de Cancelación','mdiAdministracion','MN','mnuCuentaContable','TTABLA-TIPOPAGO',32)
INSERT INTO TACCESO values ('10206000','03','Otros Tipos de Cancelación','mdiAdministracion','MN','mnuCancelacion','TTABLA-TIPOCANCELACION',26)
--INSERT INTO TACCESO values ('10207000','03','Clientes Cuentas Corrientes','mdiAdministracion','MN','mnuCliente',null,4)
INSERT INTO TACCESO values ('10208000','03','Clientes Frecuentes','mdiAdministracion','MN','mnuDelivery',null,4)
INSERT INTO TACCESO values ('10209000','03','Clientes Facturados','mdiAdministracion','MN','mnuClienteFactura',null,3)
--INSERT INTO TACCESO values ('10209500','03','Maitres','mdiAdministracion','MN','mnuMaitre','TTABLA-MAITRE',39)
INSERT INTO TACCESO values ('10210000','03','Mozos','mdiAdministracion','MN','mnuMozo','TTABLA-MOZO',23)
INSERT INTO TACCESO values ('10211000','03','Motorizados','mdiAdministracion','MN','mnuMotorizados','TTABLA-MOTORIZADO',24)
INSERT INTO TACCESO values ('10212000','03','Empacadores','mdiAdministracion','MN','mnuEmpacador','TTABLA-EMPACADOR',33)
INSERT INTO TACCESO values ('10213000','03','Zonas','mdiAdministracion','MN','mnuZona','TTABLA-ZONA',34)
INSERT INTO TACCESO values ('10214000','03','Distritos','mdiAdministracion','MN','mnuDistritos','TTABLA-DISTRITO',21)
INSERT INTO TACCESO values ('10215000','03','Mesas','mdiAdministracion','MN','mnuMesas','TMESA',8)
INSERT INTO TACCESO values ('10216000','03','Motivos de Cortesías','mdiAdministracion','MN','mnuCortesia','TTABLA-CORTESIA',20)
INSERT INTO TACCESO values ('10217000','03','Motivos de Eliminación','mdiAdministracion','MN','mnuEliminacion','TTABLA-MOTIVOELIMINACION',22)
INSERT INTO TACCESO values ('10218000','03','Motivos de Descuentos','mdiAdministracion','MN','mnuDescuento','TMOTIVODESCUENTO',9)
INSERT INTO TACCESO values ('10218100','03','Estado de Clientes Frecuentes','mdiAdministracion','MN','mnuEstadoClienteFrecuente','TTABLA-ESTADOFRECUENTE',40)
INSERT INTO TACCESO values ('10218200','03','Tipos de Egreso','mdiAdministracion','MN','mnuTipoEgreso',null,null)
INSERT INTO TACCESO values ('10219000','03','Tarjetas Bancarias','mdiAdministracion','MN','mnuTarjetaCredito','TTARJETACREDITO',16)
INSERT INTO TACCESO values ('10220000','03','Areas de Producción','mdiAdministracion','MN','mnuArea','TTABLA-AREA',35)
INSERT INTO TACCESO values ('10300000','03','Productos de Venta','mdiAdministracion','MN','mnuProd',null,null)
INSERT INTO TACCESO values ('10301000','03','Tipos de Producto','mdiAdministracion','MN','mnuTipoProducto','TTABLA-TIPOPRODUCTO',30)
INSERT INTO TACCESO values ('10301500','03','Unidad de Negocio','mdiAdministracion','MN','mnuUnidadNegocio','TTABLA-UNIDADNEGOCIO',31)
INSERT INTO TACCESO values ('10301550','03','Sucursales','mdiAdministracion','MN','mnuSucursales','TTABLA-SUCURSAL',37)
INSERT INTO TACCESO values ('10302000','03','Operadores','mdiAdministracion','MN','mnuOperador','TOPERADOR',11)
INSERT INTO TACCESO values ('10303000','03','Propiedades','mdiAdministracion','MN','mnuPropiedad','TPROPIEDAD',14)
INSERT INTO TACCESO values ('10304000','03','Grupos y SubGrupos','mdiAdministracion','MN','mnuGrupo','GRUPO',5)
INSERT INTO TACCESO values ('10305000','03','Productos y Precios','mdiAdministracion','MN','mnuProducto','PRODUCTO',13)
INSERT INTO TACCESO values ('10305500','03','Agrupacion de Caja Rápida','mdiAdministracion','MN','mnuAgrupacion','AGRUPACION',19)
INSERT INTO TACCESO values ('10305600','03','Insumos/Platos de Stock Crítico','mdiAdministracion','MN','mnuInsumoCritico','TINSUMO',38)
INSERT INTO TACCESO values ('10306000','03','Ofertas','mdiAdministracion','MN','mnuOferta','TOFERTA',10)
--INSERT INTO TACCESO values ('10308000','03','Equivalencia de Productos','mdiAdministracion','MN','mnuEquivalencias','TPRODUCTOXPRODUCTO',15)
--INSERT INTO TACCESO values ('10308000','03','Clientes Frecuentes','mdiAdministracion','MN','mnuDelivery',null,15)
INSERT INTO TACCESO values ('10400000','03','Utilitario','mdiAdministracion','MN','mnuUtilitario',null,null)
INSERT INTO TACCESO values ('10401000','03','Optimizador de la BD','mdiAdministracion','MN','mnuOptimizador',null,null)
INSERT INTO TACCESO values ('10402000','03','Backup','mdiAdministracion','MN','mnuBackup',null,null)
INSERT INTO TACCESO values ('10403000','03','Restore','mdiAdministracion','MN','mnuRestore',null,null)
INSERT INTO TACCESO values ('10404000','03','Transferencia Almacén','mdiAdministracion','MN','mnuAlmacen',null,null)
INSERT INTO TACCESO values ('10406000','03','Cambiar de Local','mdiAdministracion','MN','mnuCambiaLocal',null,null)
INSERT INTO TACCESO values ('10405000','03','Eliminar Cortesía','mdiAdministracion','MN','mnuElimina',null,null)


INSERT INTO TACCESO values ('20100000','04','Correlativos','mdiConsulta','MN','mnuCuentas',null,null)
INSERT INTO TACCESO values ('20101000','04','Pedidos','mdiConsulta','MN','mnuCorrelativoPedido',null,null)
INSERT INTO TACCESO values ('20102000','04','Documentos','mdiConsulta','MN','mnuCorrelativoDocumento',null,null)
INSERT INTO TACCESO values ('20103000','04','Cuentas Corrientes','mdiConsulta','MN','mnuCtaCte',null,null)
INSERT INTO TACCESO values ('20104000','04','Cuentas por Cobrar','mdiConsulta','MN','mnuCuentaCobrar',null,null)
INSERT INTO TACCESO values ('20105000','04','Notas de Crédito','mdiConsulta','MN','mnuNotaCredito',null,null)
INSERT INTO TACCESO values ('20106000','04','Reservas','mdiConsulta','MN','mnuReserva',null,null)
INSERT INTO TACCESO values ('20107000','04','Recibos de Egreso','mdiConsulta','MN','mnuRecibo',null,null)
INSERT INTO TACCESO values ('20108000','04','Recibos de Ingreso','mdiConsulta','MN','mnuIngreso',null,null)
INSERT INTO TACCESO values ('20109000','04','Turnos','mdiConsulta','MN','mnuTurno',null,null)
INSERT INTO TACCESO values ('20200000','04','Reportes','mdiConsulta','MN','mnuReporte',null,null)
INSERT INTO TACCESO values ('20201000','04','De Control','mdiConsulta','MN','mnuControl',null,null)
INSERT INTO TACCESO values ('20201010','04','Liquidacion de Cajero','mdiConsulta','MN','mnuLiquidacion',null,null)
INSERT INTO TACCESO values ('20201015','04','Paloteo de Producción','mdiConsulta','MN','mnuPaloteo',null,null)
INSERT INTO TACCESO values ('20201020','04','Paloteo de Propiedades','mdiConsulta','MN','mnuPropiedades',null,null)
INSERT INTO TACCESO values ('20201025','04','Paloteo de Insumos','mdiConsulta','MN','mnuPaloteoInsumo',null,null)
INSERT INTO TACCESO values ('20201026','04','Paloteo de Equivalencias','mdiConsulta','MN','mnuRepEquivalencias',null,null)
INSERT INTO TACCESO values ('20201027','04','Paloteo de Ofertas','mdiConsulta','MN','mnuPaloteoOfertas',null,null)
INSERT INTO TACCESO values ('20201028','04','Paloteo de Productos por Meses','mdiConsulta','MN','mnuPaloteoProductoMes',null,null)
INSERT INTO TACCESO values ('20201030','04','Comandas','mdiConsulta','MN','mnuComanda',null,null)
INSERT INTO TACCESO values ('20201035','04','Pedidos','mdiConsulta','MN','mnuCorrelativo',null,null)
INSERT INTO TACCESO values ('20201040','04','Propinas','mdiConsulta','MN','mnuPropina',null,null)
INSERT INTO TACCESO values ('20201045','04','Cortesias','mdiConsulta','MN','mnuCortesia',null,null)
INSERT INTO TACCESO values ('20201050','04','Descuentos','mdiConsulta','MN','mnuDescuento',null,null)
INSERT INTO TACCESO values ('20201055','04','Cuentas Corrientes','mdiConsulta','MN','mnuRepCtaCte',null,null)
INSERT INTO TACCESO values ('20201060','04','Cuentas por Cobrar','mdiConsulta','MN','mnuClienteDeuda',null,null)
INSERT INTO TACCESO values ('20201065','04','Contactos','mdiConsulta','MN','mnuContacto',null,null)
INSERT INTO TACCESO values ('20201070','04','Control de Transacciones','mdiConsulta','MN','mnuAnulado',null,null)
INSERT INTO TACCESO values ('20201075','04','Diferencias entre Paloteo vs Liquidacion','mdiConsulta','MN','mnuDiferencias',null,null)
INSERT INTO TACCESO values ('20202000','04','Contables','mdiConsulta','MN','mnuContables',null,null)
INSERT INTO TACCESO values ('20202010','04','Registro de Ventas','mdiConsulta','MN','mnuRegistroVenta',null,null)
INSERT INTO TACCESO values ('20202020','04','Principales Clientes','mdiConsulta','MN','mnuPrincipal',null,null)
INSERT INTO TACCESO values ('20202030','04','Cobranzas','mdiConsulta','MN','mnuCobranza',null,null)
INSERT INTO TACCESO values ('20203000','04','Estadisticos','mdiConsulta','MN','mnuEstadistico',null,null)
INSERT INTO TACCESO values ('20203010','04','Ranking de Produccion','mdiConsulta','MN','mnuRanking',null,null)
INSERT INTO TACCESO values ('20203015','04','Analítico de Productos por Mozos','mdiConsulta','MN','mnuMozo',null,null)
INSERT INTO TACCESO values ('20203020','04','Analítico de Productos por Motorizados','mdiConsulta','MN','mnuMotorizado',null,null)
INSERT INTO TACCESO values ('20203030','04','Analítico de Clientes Frecuentes','mdiConsulta','MN','mnuFrecuente',null,null)
INSERT INTO TACCESO values ('20203035','04','Producción por Mozos','mdiConsulta','MN','mnuEstPropina',null,null)
INSERT INTO TACCESO values ('20203036','04','Planilla de Motorizados','mdiConsulta','MN','mnuPlanillaMotorizados',null,null)
INSERT INTO TACCESO values ('20203040','04','Tiempos en Salon','mdiConsulta','MN','mnuTiempoSalon',null,null)
INSERT INTO TACCESO values ('20203045','04','Tiempos Delivery','mdiConsulta','MN','mnuTiempoDelivery',null,null)
INSERT INTO TACCESO values ('20203046','04','Tiempos KDS','mdiConsulta','MN','mnuTiempoKDS',null,null)
INSERT INTO TACCESO values ('20203047','04','Tiempos Chef Control','mdiConsulta','MN','mnuTiempoChefControl',null,null)
INSERT INTO TACCESO values ('20203050','04','Diferencias de Tiempos Delivery','mdiConsulta','MN','mnuDiferenciaDelivery',null,null)
INSERT INTO TACCESO values ('20203055','04','Rotación de Mesas','mdiConsulta','MN','mnuRotacion',null,null)
INSERT INTO TACCESO values ('20203060','04','Ocupabilidad de Mesas','mdiConsulta','MN','mnuOcupabilidad',null,null)
INSERT INTO TACCESO values ('20203065','04','Paloteo Comparativo','mdiConsulta','MN','mnuPaloteoComparativo',null,null)
INSERT INTO TACCESO values ('20203070','04','Resultados Operativos de Ventas','mdiConsulta','MN','mnuResultadoOperativo',null,null)
INSERT INTO TACCESO values ('20204000','04','Gerencial','mdiConsulta','MN','mnuAnalitico',null,null)
INSERT INTO TACCESO values ('20204010','04','Venta Anual por Meses','mdiConsulta','MN','mnuVenta',null,null)
INSERT INTO TACCESO values ('20204020','04','Venta Mensual por Fechas','mdiConsulta','MN','mnuVentasFechas',null,null)
INSERT INTO TACCESO values ('20204030','04','Venta Comparativa Anual','mdiConsulta','MN','mnuVentaComparada',null,null)
INSERT INTO TACCESO values ('20204040','04','Venta Comparativo Mensual','mdiConsulta','MN','mnuVentaComparadaMensual',null,null)
INSERT INTO TACCESO values ('20204050','04','Venta por Turnos','mdiConsulta','MN','mnuVentaTurno',null,null)
INSERT INTO TACCESO values ('20204060','04','Cobranza Mensual por Fechas','mdiConsulta','MN','mnuCobranzaFecha',null,null)
INSERT INTO TACCESO values ('20300000','04','Ticketera	','mdiConsulta','MN','mnuTicketera',null,null)
INSERT INTO TACCESO values ('20301000','04','Liquidación de Cajero','mdiConsulta','MN','mnuLiquidacionTicket',null,null)
INSERT INTO TACCESO values ('20302000','04','Paloteo de Producción','mdiConsulta','MN','mnuPaloteoTicket',null,null)
INSERT INTO TACCESO values ('20500000','04','Conexión	','mdiConsulta','MN','mnuConexion',null,null)
INSERT INTO TACCESO values ('20501000','04','Cambiar de Local','mdiConsulta','MN','mnuCambiaLocal',null,null)

INSERT INTO TACCESO values ('30100000','02','Movimientos','mdiConsulta','MN','mnuMovimiento',null,null)
INSERT INTO TACCESO values ('30101000','02','Apertura de Turno','mdiPuntoVenta','MN','mnuInicio',null,null)
INSERT INTO TACCESO values ('30102000','02','Punto de Venta','mdiPuntoVenta','MN','mnuVenta',null,null)
INSERT INTO TACCESO values ('30103000','02','Activar PinPad','mdiPuntoVenta','MN','mnuPinPad',null,null)
INSERT INTO TACCESO values ('30104000','02','Cierre de Turno','mdiPuntoVenta','MN','mnuCierre',null,null)
INSERT INTO TACCESO values ('30105000','02','Enumeración de Documentos','mdiPuntoVenta','MN','mnuCorrelativo',null,null)
INSERT INTO TACCESO values ('30106000','02','Mesas','mdiPuntoVenta','MN','mnuMesa',null,null)
INSERT INTO TACCESO values ('30106100','02','Insumos/Platos de Stock Crítico','mdiPuntoVenta','MN','mnuInsumoCritico',null,null)
INSERT INTO TACCESO values ('30200000','02','Correlativos','mdiPuntoVenta','MN','mnuCorrelativo',null,null)
INSERT INTO TACCESO values ('30201000','02','Correlativo de Pedidos','mdiPuntoVenta','MN','mnuCorrelativoPedido',null,null)
INSERT INTO TACCESO values ('30202000','02','Correlativo de Documentos','mdiPuntoVenta','MN','mnuCorrelativoDocumento',null,null)
INSERT INTO TACCESO values ('30203000','02','Cuentas Corrientes','mdiPuntoVenta','MN','mnuCtaCte',null,null)

INSERT INTO TACCESO values ('30204000','02','Recibo de Egresos','mdiPuntoVenta','MN','mnuRecibo',null,null)
INSERT INTO TACCESO values ('30204010','02','Opción Modificar','frmReciboEgresoDetalle','MN','cmdopcion1','MEGRESO',null)

INSERT INTO TACCESO values ('30205000','02','Recibo de Ingresos','mdiPuntoVenta','MN','mnuReciboIngreso',null,null)
INSERT INTO TACCESO values ('30206000','02','Notas de Crédito','mdiPuntoVenta','MN','mnuNotaCredito',null,null)
INSERT INTO TACCESO values ('30207000','02','Reservas','mdiPuntoVenta','MN','mnuReserva',null,null)
INSERT INTO TACCESO values ('30208000','02','Cuentas por Cobrar','mdiPuntoVenta','MN','mnuCuentaCobrar',null,null)
INSERT INTO TACCESO values ('30300000','02','Conexión','mdiPuntoVenta','MN','mnuConexion',null,null)
INSERT INTO TACCESO values ('30301000','02','Cambiar de Local','mdiPuntoVenta','MN','mnuCambiaLocal',null,null)
GO

if (select count(tCodigo) from TTABLA where tTabla='ETIQUETA') = 0 
BEGIN
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('ETIQUETA', '01',2,'','',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('ETIQUETA', '02',2,'','',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('ETIQUETA', '03',2,'','',1)
END
GO

if (select count(tCodigo) from TTABLA where tTabla='CAJARAPIDA') = 0 
BEGIN
   insert into TTABLA select 'CAJARAPIDA','11',null,'','',null,1,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','12',null,'','',null,2,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','13',null,'','',null,3,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','14',null,'','',null,4,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','15',null,'','',null,5,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','16',null,'','',null,6,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','17',null,'','',null,7,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','18',null,'','',null,8,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','19',null,'','',null,9,'',0,1,null,''

   insert into TTABLA select 'CAJARAPIDA','21',null,'','',null,1,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','22',null,'','',null,2,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','23',null,'','',null,3,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','24',null,'','',null,4,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','25',null,'','',null,5,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','26',null,'','',null,6,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','27',null,'','',null,7,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','28',null,'','',null,8,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','29',null,'','',null,9,'',0,1,null,''

   insert into TTABLA select 'CAJARAPIDA','31',null,'','',null,1,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','32',null,'','',null,2,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','33',null,'','',null,3,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','34',null,'','',null,4,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','35',null,'','',null,5,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','36',null,'','',null,6,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','37',null,'','',null,7,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','38',null,'','',null,8,'',0,1,null,''
   insert into TTABLA select 'CAJARAPIDA','39',null,'','',null,9,'',0,1,null,''
end
GO
delete from  tusuario where substring(tcodigousuario,1,1) like '*'

go

if (select count(TCODIGOUSUARIO) from TUSUARIO where SUBSTRING(TCODIGOUSUARIO,1,1)='*')=0 
begin

insert into TUSUARIO select '*0001','00','OMELERO','OMELERO','\:##_\@#',1,GETDATE(),'\¡5¡¡¡¡¡¡¡¡¡¡5','',''
insert into TUSUARIO select '*0002','00','OMELEROB','OMELEROB','>:_#¡@|5',1,GETDATE(),'','',''
insert into TUSUARIO select '*0003','00','LHEREDIA','LHEREDIA','>@|\@_|¬',1,GETDATE(),'','',''
insert into TUSUARIO select '*0004','00','ESANTILLANA','ESANTILLANA','>¡>__#¡>',1,GETDATE(),'','',''
insert into TUSUARIO select '*0006','00','JCARDENAS','JCARDENAS','>>¬#||5¡',1,GETDATE(),'','',''
insert into TUSUARIO select '*0007','00','SCASTILLO','SCASTILLO','¡||¡¡@##',1,GETDATE(),'','',''
insert into TUSUARIO select '*0009','00','SPENA','SPENA','>>@_::¬>',1,GETDATE(),'','',''
insert into TUSUARIO select '*0010','00','ADCASTILLO','ADCASTILLO','>#>¡|@@|',1,GETDATE(),'','',''
insert into TUSUARIO select '*0011','00','RMALLA','RMALLA','5¡\_|@:#',1,GETDATE(),'','',''
insert into TUSUARIO select '*0012','00','NINGAROCA','NINGAROCA','5¡@5@|:_',1,GETDATE(),'','',''
insert into TUSUARIO select '*0013','00','ADELGADO','ADELGADO','_>_@5#5>',1,GETDATE(),'','',''
insert into TUSUARIO select '*0014','00','RPADILLA','RPADILLA','>¡¬55\¡¬',1,GETDATE(),'','',''
insert into TUSUARIO select '*0015','00','RNUNEZ','RNUNEZ','>>::>¬||',1,GETDATE(),'','',''
insert into TUSUARIO select '*0016','00','DLUNA','DLUNA','>>¬@#>#|',1,GETDATE(),'','',''
insert into TUSUARIO select '*0017','00','TTORRES','TTORRES','>5_5@>\\',1,GETDATE(),'','',''
--insert into TUSUARIO select '*0018','00','LGRADOS','LGRADOS','>\|>|55',1,GETDATE(),'' ,'',''
insert into TUSUARIO select '*0019','00','CLAROSA','CLAROSA','>\\>_:>>',1,GETDATE(),'','',''
insert into TUSUARIO select '*0020','00','CNINAHUANCA','CNINAHUANCA','>#:>>¡\#',1,GETDATE(),'','',''
insert into TUSUARIO select '*0021','00','JAVILA','JAVILA','>#¬|¬_¡@',1,GETDATE(),'','',''
insert into TUSUARIO select '*0023','00','EVILLEGAS','EVILLEGAS','>\|:_@__',1,GETDATE(),'','',''
insert into TUSUARIO select '*0024','00','RMANTILLA','RMANTILLA','\:#¡__@>',1,GETDATE(),'','',''
insert into TUSUARIO select '*0025','00','FRIEGA','FRIEGA','>>##|@:@',1,GETDATE(),'','',''
insert into TUSUARIO select '*0026','00','ASALAS','ASALAS','>>¡@#|_|',1,GETDATE(),'','',''
insert into TUSUARIO select '*0027','00','LVEGA','LVEGA','>_5:5\5_',1,GETDATE(),'','',''
insert into TUSUARIO select '*0028','00','ACAYCHO','ACAYCHO','>_¡\|:>>',1,GETDATE(),'','',''
insert into TUSUARIO select '*0029','00','MVILLARAN','MVILLARAN','_#5|\\:@',1,GETDATE(),'','',''

end

go

DELETE FROM TTABLA WHERE TTABLA='MOZO' AND substring(TCODIGO,1,1)='*'
 
go

if (select count(TCODIGO) from TTABLA where SUBSTRING(TCODIGO,1,1)='*' and TTABLA='MOZO')=0 
begin
	insert into TTABLA select 'MOZO','*001',0,'OMELERO','OMELERO',		'\¡5¡¡¡¡¡¡¡¡¡¡5',0,'\:##_\@#',1,1,1,''
	insert into TTABLA select 'MOZO','*002',0,'OMELEROB','OMELEROB',	'',0,'>:_#¡@|5',1,1,1,''
	insert into TTABLA select 'MOZO','*003',0,'LHEREDIA','LHEREDIA',		'',0,'>@|\@_|¬',1,1,1,''
	insert into TTABLA select 'MOZO','*004',0,'ESANTILLANA','ESANTILLANA',	'',0,'>¡>__#¡>',1,1,1,''
	insert into TTABLA select 'MOZO','*006',0,'JCARDENAS','JCARDENAS',		'',0,'>>¬#||5¡',1,1,1,''
	insert into TTABLA select 'MOZO','*007',0,'SCASTILLO','SCASTILLO',		'',0,'¡||¡¡@##',1,1,1,''
	insert into TTABLA select 'MOZO','*009',0,'RBARBOZA','RBARBOZA',		'',0,'>>@_::¬>',1,1,1,''
	insert into TTABLA select 'MOZO','*010',0,'ADCASTILLO','ADCASTILLO',		'',0,'>#>¡|@@|',1,1,1,''
	insert into TTABLA select 'MOZO','*011',0,'RMALLA','RMALLA',		'',0,'5¡\_|@:#',1,1,1,''
	insert into TTABLA select 'MOZO','*012',0,'NINGAROCA','NINGAROCA',	'',0,'5¡@5@|:_',1,1,1,''
	insert into TTABLA select 'MOZO','*013',0,'ADELGADO','ADELGADO',		'',0,'_>_@5#5>',1,1,1,''
	insert into TTABLA select 'MOZO','*014',0,'RPADILLA','RPADILLA',		'',0,'>¡¬55\¡¬',1,1,1,''
	insert into TTABLA select 'MOZO','*015',0,'RNUNEZ','RNUNEZ',		'',0,'>>::>¬||',1,1,1,''
	insert into TTABLA select 'MOZO','*016',0,'DLUNA','DLUNA',		'',0,'>>¬@#>#|',1,1,1,''
	insert into TTABLA select 'MOZO','*017',0,'TTORRES','TTORRES',		'',0,'>5_5@>\\',1,1,1,''
	--insert into TTABLA select 'MOZO','*018',0,'LGRADOS','LGRADOS',		'',0,'>\|>|55',1,1,1,''
	insert into TTABLA select 'MOZO','*019',0,'CLAROSA','CLAROSA',		'',0,'>\\>_:>>',1,1,1,''
	insert into TTABLA select 'MOZO','*020',0,'CNINAHUANCA','CNINAHUANCA',	'',0,'>#:>>¡\#',1,1,1,''
	insert into TTABLA select 'MOZO','*021',0,'JAVILA','JAVILA',		'',0,'>#¬|¬_¡@',1,1,1,''
	insert into TTABLA select 'MOZO','*023',0,'EVILLEGAS','EVILLEGAS',		'',0,'>\|:_@__',1,1,1,''
	insert into TTABLA select 'MOZO','*024',0,'RMANTILLA','RMANTILLA',	'',0,'\:#¡__@>',1,1,1,''
	insert into TTABLA select 'MOZO','*025',0,'FRIEGA','FRIEGA',		'',0,'>>##|@:@',1,1,1,''
	insert into TTABLA select 'MOZO','*026',0,'ASALAS','ASALAS',	'',0,'>>¡@#|_|',1,1,1,''
	
	insert into TTABLA select 'MOZO','*027',0,'LVEGA','LVEGA',		'',0,'>_5:5\5_',1,1,1,''
	insert into TTABLA select 'MOZO','*028',0,'ACAYCHO','ACAYCHO',	'',0,'>_¡\|:>>',1,1,1,''

END

GO

if (select count(TGRUPOUSUARIO) from TGRUPOUSUARIO where TGRUPOUSUARIO='00')=0 
begin

	insert into TGRUPOUSUARIO select '00','INFHOTEL',1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,'00',1
END
else
begin
update TGRUPOUSUARIO
set tDetallado='INFHOTEL', lModulo01=1,lModulo02=1,lModulo03=1, lModulo04=1,  lOpcion01=1,
lOpcion02=1,lOpcion03=1,lOpcion04=1,lOpcion05=1,lOpcion06=1,lOpcion07=1,lOpcion08=1,lOpcion09=1,lOpcion10=1,lOpcion11=1,
lOpcion12=1,lOpcion13=1,lOpcion14=1,lOpcion15=1,lOpcion16=1,lOpcion17=1,lOpcion18=1,lOpcion19=1,lOpcion20=1,lOpcion21=1,
lactivo=1, lModulo05=1
where TGRUPOUSUARIO='00'
end
go


if (select count(tF1) from TMENSAJE) = 0
   INSERT INTO TMENSAJE (tF1) VALUES ('')
GO

--Actualiza Datos
UPDATE TPRODUCTO SET TUNIDADNEGOCIO='00' where isnull(TUNIDADNEGOCIO,'')=''
UPDATE DPEDIDO SET TUNIDADNEGOCIO='00' where isnull(TUNIDADNEGOCIO,'')=''

update tTipoDocumentoImpresora set lImpuesto1 = 0 where isnull(lImpuesto1,0)=0
update tTipoDocumentoImpresora set lImpuesto2 = 0 where isnull(lImpuesto2,0)=0
update tTipoDocumentoImpresora set lImpuesto3 = 0 where isnull(lImpuesto3,0)=0

update tProducto set lImpuesto4 = 0 where isnull(lImpuesto4,0)=0
update tProducto set lImpuesto5 = 0 where isnull(lImpuesto5,0)=0
update tProducto set lImpuesto6 = 0 where isnull(lImpuesto6,0)=0
update tProducto set lImpuesto7 = 0 where isnull(lImpuesto7,0)=0
update tProducto set lImpuesto8 = 0 where isnull(lImpuesto8,0)=0
update tProducto set lImpuesto9 = 0 where isnull(lImpuesto9,0)=0
update tProducto set lImpuesto10 = 0 where isnull(lImpuesto10,0)=0
update tProducto set lImpuesto11 = 0 where isnull(lImpuesto11,0)=0
update tProducto set lImpuesto12 = 0 where isnull(lImpuesto12,0)=0
update tProducto set lImpuesto13 = 0 where isnull(lImpuesto13,0)=0
update tProducto set lImpuesto14 = 0 where isnull(lImpuesto14,0)=0
update tProducto set lImpuesto15 = 0 where isnull(lImpuesto15,0)=0
update tproducto set lDescuento=1 where isnull(lDescuento,1)=1
update tProducto set lLocal = 1 where isnull(lLocal,1)=1
update tProducto set lDelivery = 1 where isnull(lDelivery,1)=1
update tProducto set lLlevar = 1 where isnull(lLlevar,1)=1
update tProducto set lCanal4 = 1 where isnull(lCanal4,1)=1
update tProducto set lCanal5 = 1 where isnull(lCanal5,1)=1
update tProducto set nPrecioDelivery = 0 where isnull(nPrecioDelivery,0)=0
update tProducto set nPrecioLlevar = 0 where isnull(nPrecioLlevar,0)=0
update tProducto set nPrecioCanal4 = 0 where isnull(nPrecioCanal4,0)=0
update tProducto set nPrecioCanal5 = 0 where isnull(nPrecioCanal5,0)=0
update tproducto set ninsumo=0 where isnull(nInsumo,0)=0
update tproducto set ninsumo2=0 where isnull(nInsumo2,0)=0
update tproducto set ninsumo3=0 where isnull(nInsumo3,0)=0
update tproducto set ninsumo4=0 where isnull(nInsumo4,0)=0
update tproducto set ninsumo5=0 where isnull(nInsumo5,0)=0
update tproducto set nManoObra=0 where isnull(nManoObra,0)=0
update tproducto set nManoObra2=0 where isnull(nManoObra2,0)=0
update tproducto set nManoObra3=0 where isnull(nManoObra3,0)=0
update tproducto set nManoObra4=0 where isnull(nManoObra4,0)=0
update tproducto set nManoObra5=0 where isnull(nManoObra5,0)=0
update tproducto set nGasto=0 where isnull(nGasto,0)=0
update tproducto set nGasto2=0 where isnull(nGasto2,0)=0
update tproducto set nGasto3=0 where isnull(nGasto3,0)=0
update tproducto set nGasto4=0 where isnull(nGasto4,0)=0
update tproducto set nGasto5=0 where isnull(nGasto5,0)=0

update TTABLA set tIcono = 0 where TTABLA = 'TIPODOCUMENTO' and isnull(tIcono,0)=0
update TTABLA set nTamano = 0 where TTABLA = 'MOZO' and isnull(nTamano,0)=0

UPDATE TDELIVERY SET lPuntos=1 where isnull(lPuntos,1)=1
UPDATE TDELIVERY SET nAcumulado=0 where isnull(nAcumulado,0)=0
UPDATE TDELIVERY SET nUtilizado=0 where isnull(nUtilizado,0)=0
UPDATE TDELIVERY SET nDisponible=0 where isnull(nDisponible,0)=0

UPDATE TCAJA SET tTipoPedido = '01' where isnull(tTipoPedido,'')=''
update TPARAMETRO set nDecimal = 2 where isnull(nDecimal,0)=0
UPDATE TTABLA SET nTamano = '1' where isnull(nTamano,1)=1 and TTABLA='TIPODOCUMENTO'

update TSUBGRUPO set nOrden=1 where isnull(nOrden,0)=0
delete from tproductoarea where tcodigoproducto in (select tcodigoproducto from tproducto where lcombinacion=1)

update mpedido set fLlegada =fRegistro where isnull(fLlegada,0)=0 and tTipoPedido='02'

update TOFERTA set lCanal4=1 where isnull(lCanal4,0)=0
update TOFERTA set lCanal5=1 where isnull(lCanal5,0)=0
update treserva set tprioridad=1

UPDATE TUSUARIO set tBandaMagnetica='' where isnull(tBandaMagnetica,'')=''
UPDATE TPROPIEDAD set tEnlace='' where isnull(tEnlace,'')='' 
UPDATE TPRODUCTOPROPIEDAD set tEnlace='' where isnull(tEnlace,'')='' 
UPDATE TCOMBOPROPIEDAD set tEnlace='' where isnull(tEnlace,'')='' 
UPDATE TMOTIVODESCUENTO set lTopePedido=0 where isnull(lTopePedido,0)=0
UPDATE TMOTIVODESCUENTO set lBloqueo=0 where isnull(lBloqueo,0)=0
UPDATE TMOTIVODESCUENTO set lRatio=0 where isnull(lRatio,0)=0
UPDATE TMOTIVODESCUENTO set lAplicablePedido=0 where isnull(lAplicablePedido,0)=0
GO

update TTABLA set tResumido = substring(tResumido,1,15) where tTABLA='MOZO' and len(ltrim(tResumido))>15
update TUSUARIO set tResumido = substring(tResumido,1,15) 
GO

delete from taccesoenvia where substring(tcodigoacceso,4,5)='00000'
go

update TLOCAL set tCodigoLocal='0'+LTRIM(RTRIM(tCodigoLocal)) where LEN(LTRIM(tCodigoLocal))=2 
update TTABLA set tValor='0'+LTRIM(RTRIM(tValor)) where LEN(LTRIM(tValor))=2 and TTABLA='SALON'
GO

--insert into TOPERADOR 
--select tcodigo, tdetallado, tresumido, tvalor, 0, nvalor, 0, nboton, getdate(), 'MASTER', lactivo, 1 from TTABLA where tTabla='OPERADOR'
--delete from TTABLA where tTabla='OPERADOR'

--INSERT INTO TMOTIVODESCUENTO 
--SELECT tCodigo, tDetallado, tResumido, nValor, 0, 1, 1, 1 
--FROM TTABLA where TTABLA='MOTIVODESCUENTO'
--delete ttabla where TTABLA='MOTIVODESCUENTO'

--DELETE FROM TTABLA WHERE TTABLA='GRUPOUSUARIO'
--DELETE FROM TGRUPOUSUARIO
--INSERT INTO TGRUPOUSUARIO (tGrupoUsuario, tDetallado, lModulo01, lModulo02, lModulo03, lOpcion01,lOpcion02,lOpcion03,lOpcion04,lOpcion05,lOpcion06,lOpcion07,lOpcion08,lOpcion09,lOpcion10, lOpcion11, lOpcion12, lOpcion13, lOpcion14, lOpcion15, lActivo) 
--values ('01','SUPERVISOR',1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1)
--INSERT INTO TGRUPOUSUARIO (tGrupoUsuario, tDetallado, lModulo01, lModulo02, lModulo03, lOpcion01,lOpcion02,lOpcion03,lOpcion04,lOpcion05,lOpcion06,lOpcion07,lOpcion08,lOpcion09,lOpcion10, lOpcion11, lOpcion12, lOpcion13, lOpcion14, lOpcion15, lActivo) 
--values ('02','CAJA',1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1)

--Update CPEDIDO set nInsumo = ISNULL(TPRODUCTO.nInsumo,0), nGasto=ISNULL(TPRODUCTO.nGasto,0), nManoObra=ISNULL(TPRODUCTO.nManoObra,0)
--From CPEDIDO, TPRODUCTO
--WHERE CPEDIDO.tProductoCombo=TPRODUCTO.tCodigoProducto and isnull(CPEDIDO.nInsumo,0)=0 


--insert into ttabla (TTABLA,TCODIGO,TDETALLADO,TRESUMIDO,LACTIVO) values('SECTOR','00','SECTOR0','SECTOR0',1)
go

update TPARAMETRO  set lImprimeDiaContable =0 where lImprimeDiaContable is null

go
if (select count(tCodigo) from TTABLA where tTabla='PAISORIGEN' and TCODIGO='000') = 0 
Begin
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
VALUES('PAISORIGEN', '000',2,'PERU','PERU',1,1)
end
GO

if (select count(tCodigo) from TTABLA where tTabla='PAISORIGEN' and TCODIGO='001') = 0 
Begin
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
VALUES('PAISORIGEN', '001',2,'BOLIVIA','BOLIVIA',0,1)
end
GO

if (select count(tCodigo) from TTABLA where tTabla='PAISORIGEN' and TCODIGO='002') = 0 
Begin
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
VALUES('PAISORIGEN', '002',2,'ECUADOR','ECUADOR',0,1)
end
GO

if (select count(tCodigolocal) from TLOCAL) = 0 
BEGIN
insert into tlocal (tcodigolocal,tdetallado,tresumido,lactivo)
values('01','PRINCIPAL','PRINCIPAL',1)
end
GO

 --Sucursal 
IF(select count(tCodigo) from TTABLA WHERE tCodigo ='00' and TTABLA ='SUCURSAL' )= 0
BEGIN 
insert into TTABLA( tTabla, tCodigo, tDetallado, tResumido, nValor, tValor, tIcono, lActivo) values ('SUCURSAL',  '00',  'SIN SUCURSAL',  'SIN SUCURSAL', 0,'','',1) 
END
 go
 
--diacontable
if (select case when ldiacontableautomatico=1 then 2 else 0 end + case when ldiacontablemanual=1 then 2 else 0 end from tparametro) =0
begin
update tparametro set ldiacontableautomatico=1 , thoracierrediacontable='06:00'
end 
GO

--SMALLDATETIME A DATETIME
if (SELECT    upper(systypes.name)      FROM         syscolumns INNER JOIN  sysobjects ON syscolumns.id = sysobjects.id INNER JOIN systypes ON systypes.xtype = syscolumns.xtype WHERE     (sysobjects.name = 'DPEDIDO') AND  (systypes.name <> 'sysname')  AND   syscolumns.name='fenvio')='SMALLDATETIME'
BEGIN
	exec RSP_ALTERACOLUMNA 'DPEDIDO','FENVIO','datetime','8','N'
END 
GO

if (select count(*) from ttabla where TTABLA='ESTADOFRECUENTE') = 0 
BEGIN
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOFRECUENTE','01','ACTIVO','ACTIVO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOFRECUENTE','02','SUSP CREDITO','SUSP CREDITO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOFRECUENTE','03','SUSP DERECHO','SUSP DERECHO',0,1)
end
go

if (select count(tCodigo) from TTABLA where tTabla='TIPODOCUMENTO') > 0 
BEGIN
	insert into TTIPODOCUMENTO  
	select tCodigo, tDetallado, tResumido, tValor, nValor, tIcono, 0, nTamano, lActivo, 0,'' from TTABLA where tTABLA='TIPODOCUMENTO'

	DELETE FROM TTABLA where tTABLA='TIPODOCUMENTO'
END
GO

if (select count(tCodigoCanalVenta) from TCANALVENTA) = 0
begin
   INSERT INTO TCANALVENTA
   select tCodigo, tDetallado, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, lActivo from ttabla where ttabla='TIPOPEDIDO'
end
GO
if (select count(tCodigoCanalVenta) from TCANALVENTA) = 1 
begin
   INSERT INTO TCANALVENTA select '02', '', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0
end
GO
if (select count(tCodigoCanalVenta) from TCANALVENTA) = 2
begin
   INSERT INTO TCANALVENTA select '03', '', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0
end
GO
if (select count(tCodigoCanalVenta) from TCANALVENTA) = 3 
begin
   INSERT INTO TCANALVENTA select '04', '', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0
end
GO
if (select count(tCodigoCanalVenta) from TCANALVENTA) = 4
begin
   INSERT INTO TCANALVENTA select '05', '', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0
end
GO

if (select count(TTABLA) from TTABLA where tTabla='TTIPOPEDIDO') > 0
Begin
	Update TCANALVENTA set lObligaMesa=TCAJA.lObliga, lObligaPax=TCAJA.lPax, lObligaMozo=TCAJA.lMozo, lObligaMotorizado=TCAJA.lMotorizado
	from TCAJA, TCANALVENTA
	where TCAJA.tTipoPedido=TCANALVENTA.tCodigoCanalVenta 
	delete FROM TTABLA where tTabla='TTIPOPEDIDO'
End
GO

if (select avg(len(ltrim(tCodigoDelivery))) FROM TDELIVERY) < 6
Begin
   update TDELIVERY set tCodigoDelivery = right('00'+ ltrim(tCodigoDelivery),7)
End
GO
if (select count(tCodigoCliente) FROM TCOMPANIA) > 0
Begin
	Declare @Correlativo as int
	set @Correlativo = (select ISNULL(max(tCodigoDelivery),0) FROM TDELIVERY)
	insert into TDELIVERY
	select substring('0000000',1,7-len(ltrim(str(@correlativo+tCodigoCliente))))+ltrim(str(@correlativo+tCodigoCliente)), 
	'00', tapecom, tnomsoc, tDireccion, tTelefono1, '','','','','','','',0,
	null,tEmail,'',lActivo,0,0,0,0, tUsuario, fRegistro,getdate(), lReplica,
	'00', 0, 1, nConsumo, nLinea, tTipoCtaCte, tSubTipoCtaCte, '',ISNULL(tIdentidad,''),'','',0,0
	From tCompania


	update MPEDIDO set tClienteDelivery= right('00'+tClienteDelivery,7) where len(tClienteDelivery)>0
	update MPEDIDO set tClienteCtaCte= substring('0000000',1,7-len(ltrim(str(@correlativo+tClienteCtaCte))))+ltrim(str(@correlativo+tClienteCtaCte))
	where len(tClienteCtaCte)>0
	delete from tCompania
End
GO

delete from TTRAMITE
insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
				values ('00001','Solicitud de Autorizacion', 0,1)
GO			
insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
				values ('00002','Solicitud de Autorizacion por Cambio de Software',1,1)
GO
insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
				values ('00003','Solicitud de Renovacion de Autorizacion', 1,1)				
				
GO
insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
				values ('00004','Baja de la Autorizacion', 1,1)				
GO
insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
				values ('00005','Inclusion de Puntos de Emision y Tipos de Documentos', 0,1)				
GO
insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
				values ('00006','Exclusion de Puntos de Emision y Tipos de Documentos', 1,1)						
GO

delete FROM TTABLA where TTABLA='ESTADOSOLICITUD'
GO
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','01','ACTIVO','ACTIVO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','02','APROBADO','APROBADO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','03','EN TRAMITE','EN TRAMITE',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','04','RECHAZADO','RECHAZADO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','05','ANULADO','ANULADO',0,1)	
go

delete FROM TTABLA where TTABLA='ESTADODETSOLICITUD'
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADODETSOLICITUD','01','APROBADO','APROBADO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADODETSOLICITUD','02','EN TRAMITE','EN TRAMITE',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADODETSOLICITUD','03','RECHAZADO','RECHAZADO',0,1)
INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADODETSOLICITUD','04','DE BAJA','ANULADO',0,1)
GO


update TPARAMETRO set lMobilePasswordCCaja=0 where lMobilePasswordCCaja is null
go
update TPARAMETRO set lMobileUnidadNegocio=0 where lMobileUnidadNegocio is null

GO
				
update tcaja set lBloqueaPrecuenta=0 where lBloqueaPrecuenta is null

go
update TCAJA set lMultiareasubgrupo=0 where lMultiareasubgrupo is null
go
update TCAJA set lMultiareacaja=0 where lMultiareacaja is null
go
if (select count(*) from TPARAMETRO where tVersion<'4.94.4111')>0
begin
		update TCAJA set lmultiareacaja=1 where isnull(tsubalmacen,'')<>''
end

go
update tcaja set tsectorventa='' where tsectorventa is null
GO

-----CESAR DOCUMENTOS VARIABLES
DELETE FROM TTABLA where tTabla='FORMULARIO'
INSERT TTABLA (TTABLA, TCODIGO, nTamano, tDetallado, tResumido, lActivo) values ('FORMULARIO','01',2,'TICKET','TICKET',1)
INSERT TTABLA (TTABLA, TCODIGO, nTamano, tDetallado, tResumido, lActivo) values ('FORMULARIO','02',2,'VARIABLE','VARIABLE',1)
INSERT TTABLA (TTABLA, TCODIGO, nTamano, tDetallado, tResumido, lActivo) values ('FORMULARIO','03',2,'TICKETVARIABLE','TICKETVARIABLE',1)
-------

----GRUPO USUARIO
UPDATE TGRUPOUSUARIO SET tNivel = tGrupoUsuario
----

UPDATE TPARAMETRO SET tVersion='4.94.4262'
GO
PRINT ' LISTO '




