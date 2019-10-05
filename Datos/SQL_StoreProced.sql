
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