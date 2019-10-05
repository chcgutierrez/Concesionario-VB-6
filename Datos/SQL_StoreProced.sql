
/*=====================================================================
Nombre: sp_buscar_tipodoc
Objetivo: Buscar el registro en la tabla tb_tipodoc segun criterio
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_buscar_tipodoc
@tipo_doc varchar (10)
as
select
id_tipodoc,
tipo_doc,
des_tip_doc,
est_tip_doc,
fec_act,
obs_gen
from
	tb_tipodoc (nolock)
where
tipo_doc = @tipo_doc



/*=====================================================================
Nombre: sp_editar_tipodoc
Objetivo: Actualizar el registro en la tabla tb_tipodoc segun criterio
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_editar_tipodoc
@tipo_doc varchar (10),
@des_tip_doc varchar (100),
@est_tip_doc varchar (10),
@obs_gen text
as
update tb_tipodoc set des_tip_doc = @des_tip_doc,
					  est_tip_doc = @est_tip_doc,
					  fec_act = getdate(),
					  obs_gen = @obs_gen where
					  tipo_doc = @tipo_doc



/*=====================================================================
Nombre: sp_guardar_tipodoc
Objetivo: Insertar registros en la tabla tb_tipodoc
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_guardar_tipodoc
@tipo_doc varchar (100),
@des_tip_doc varchar (100),
@est_tip_doc varchar (10),
@obs_gen text
as
insert into tb_tipodoc (tipo_doc, des_tip_doc, est_tip_doc, fec_act, obs_gen)
	values(@tipo_doc, @des_tip_doc, @est_tip_doc, getdate(), @obs_gen)



/*=====================================================================
Nombre: sp_mostrar_tipodoc
Objetivo: Ver todos los registros de la tabla tb_tipodoc
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_mostrar_tipodoc
as
select
	id_tipodoc,
	tipo_doc,
	des_tip_doc,
	est_tip_doc,
	fec_act,
	obs_gen
from
	tb_tipodoc (nolock)

/*
=========================================================================
Tabla Pais
tb_pais
=========================================================================
*/

/*=====================================================================
Nombre: sp_buscar_pais
Objetivo: Buscar el registro en la tabla tb_pais segun criterio
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_buscar_pais
@cod_pais varchar (10)
as
select
id_pais,
cod_pais,
nom_pais,
est_pais,
fec_act,
obs_gen
from
	tb_pais (nolock)
where
cod_pais = @cod_pais

/*=====================================================================
Nombre: sp_editar_pais
Objetivo: Actualizar el registro en la tabla tb_pais segun criterio
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_editar_pais
@cod_pais varchar (10),
@nom_pais varchar (100),
@est_pais varchar (10),
@obs_gen text
as
update tb_pais set nom_pais = @nom_pais,
				   est_pais = @est_pais,
				   fec_act = getdate(),
				   obs_gen = @obs_gen where
				   cod_pais = @cod_pais

/*=====================================================================
Nombre: sp_guardar_pais
Objetivo: Insertar registros en la tabla tb_pais
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_guardar_pais
@cod_pais varchar (10),
@nom_pais varchar (100),
@est_pais varchar (10),
@obs_gen text
as
insert into tb_pais (cod_pais, nom_pais, est_pais, fec_act, obs_gen)
	values(@cod_pais, @nom_pais, @est_pais, getdate(), @obs_gen)

/*=====================================================================
Nombre: sp_mostrar_pais
Objetivo: Ver todos los registros de la tabla tb_tipodoc
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_mostrar_pais
as
select
	id_pais,
	cod_pais,
	nom_pais,
	est_pais,
	fec_act,
	obs_gen
from
	tb_pais (nolock)

INSERT INTO tb_pais VALUES ('2001','ARGENTINA','A','2019-09-01','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2002','BOLIVIA','A','2019-09-02','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2003','BRASIL','I','2019-09-03','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2004','CHILE','A','2019-09-01','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2005','COLOMBIA','A','2019-09-01','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2006','ECUADOR','A','2019-09-01','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2007','GUYANA','I','2019-09-07','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2008','PARAGUAY','A','2019-09-08','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2009','PERÚ','A','2019-09-01','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2010','SURINAM','I','2019-09-01','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2011','URUGUAY','A','2019-09-11','CARGA_INICIAL');
INSERT INTO tb_pais VALUES ('2012','VENEZUELA','I','2019-09-12','CARGA_INICIAL');

	select * from tb_pais

/*
=========================================================================
Tabla Departamento
tb_depto
=========================================================================
*/

/*=====================================================================
Nombre: sp_buscar_depto
Objetivo: Buscar el registro en la tabla tb_depto segun criterios
Fecha Creacion: 05/10/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_buscar_depto
@codPais varchar (10),
@codDepto varchar (10)
as
select
A.id_depto,
A.id_pais,
B.cod_pais,
A.cod_depto,
A.nom_depto,
A.est_depto,
A.fec_act,
A.obs_gen
from
	tb_depto A (nolock) inner join tb_pais B (nolock)
		on B.id_pais = A.id_pais
where
A.cod_depto = @codDepto and
B.cod_pais = @codPais

select * from tb_pais WHERE COD_PAIS='2005'
select * from tb_depto

insert into tb_depto values('105','3001','ANTIOQUIA','A',getdate(),'CARGA_SQL')
insert into tb_depto values('105','3002','RISARALDA','A',getdate(),'CARGA_SQL')
insert into tb_depto values('105','3003','TOLIMA','A',getdate(),'CARGA_SQL')
insert into tb_depto values('105','3004','CUNDINAMARCA','A',getdate(),'CARGA_SQL')
insert into tb_depto values('105','3005','ATLANTICO','A',getdate(),'CARGA_SQL')
insert into tb_depto values('105','3006','AMAZONAS','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3007','ARAUCA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3008','BOGOTÁ','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3009','BOLÍVAR','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3010','BOYACÁ','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3011','CALDAS','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3012','CAQUETÁ','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3013','CASANARE','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3014','CAUCA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3015','CESAR','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3016','CHOCÓ','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3017','CÓRDOBA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3018','GUAINÍA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3019','GUAVIARE','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3020','HUILA','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3021','LA GUAJIRA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3022','MAGDALENA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3023','META','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3024','NARIÑO','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3025','NORTE','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3026','PUTUMAYO','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3027','QUINDÍO','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3028','SAN ANDRÉS Y PROVIDENCIA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3029','SANTANDER','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3030','SUCRE','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3031','VALLE','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3032','VAUPÉS','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3033','VICHADA','A',GETDATE(),'CARGA_SQL');
