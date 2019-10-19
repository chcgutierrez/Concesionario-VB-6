
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
INSERT INTO tb_pais VALUES ('2009','PER�','A','2019-09-01','CARGA_INICIAL');
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
insert into tb_depto values('105','3008','BOGOT�','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3009','BOL�VAR','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3010','BOYAC�','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3011','CALDAS','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3012','CAQUET�','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3013','CASANARE','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3014','CAUCA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3015','CESAR','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3016','CHOC�','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3017','C�RDOBA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3018','GUAIN�A','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3019','GUAVIARE','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3020','HUILA','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3021','LA GUAJIRA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3022','MAGDALENA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3023','META','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3024','NARI�O','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3025','NORTE DE SANTANDER','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3026','PUTUMAYO','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3027','QUIND�O','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3028','SAN ANDR�S Y PROVIDENCIA','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3029','SANTANDER','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3030','SUCRE','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3031','VALLE','I',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3032','VAUP�S','A',GETDATE(),'CARGA_SQL');
insert into tb_depto values('105','3033','VICHADA','A',GETDATE(),'CARGA_SQL');

/*=====================================================================
Nombre: sp_mostrar_depto
Objetivo: Mostrar los registros existentes en la tabla tb_depto
Fecha Creacion: 05/10/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_mostrar_depto
as
select
B.cod_pais,
A.cod_depto,
A.nom_depto,
A.est_depto,
A.fec_act,
A.obs_gen
from
	tb_depto A (nolock) inner join tb_pais B (nolock)
		on B.id_pais = A.id_pais

/*=====================================================================
Nombre: sp_guardar_depto
Objetivo: Insertar registros en la tabla tb_depto
Fecha Creacion: 05/10/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_guardar_depto
@idPais varchar (10),
@codDepto varchar (10),
@nomDepto varchar (100),
@estDepto varchar (10),
@obsGen text
as
insert into tb_depto (id_pais, cod_depto, nom_depto, est_depto, fec_act, obs_gen)
	values(@idPais, @codDepto, @nomDepto, @estDepto, getdate(), @obsGen)

/*=====================================================================
Nombre: sp_editar_depto
Objetivo: Actualizar el registro en la tabla tb_pais segun criterio
Fecha Creacion: 10/05/2019
Autor: chcgutierrez
=======================================================================*/
create procedure sp_editar_depto
@codPais varchar (10),
@codDepto varchar (10),
@nomDepto varchar (100),
@estDepto varchar (10),
@obsDepto text
as
update tb_depto set nom_depto = @nomDepto,
				    est_depto = @estDepto,
				    fec_act = getdate(),
				    obs_gen = @obsDepto where
				    id_pais = @codPais and
					cod_depto = @codDepto


select * from tb_pais
select * from tb_depto
select * from tb_ciudad

INSERT INTO tb_ciudad VALUES ('101','105','4001','ABEJORRAL','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4002','ABRIAQU�','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4003','ALEJANDR�A','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4004','AMAG�','I',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4005','AMALFI','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4006','ANDES','I',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4007','ANGEL�POLIS','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4008','ANGOSTURA','I',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4009','ANOR�','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4010','ANZA','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4011','APARTAD�','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4012','ARBOLETES','A',GETDATE(),'CARGA_SQL');
INSERT INTO tb_ciudad VALUES ('101','105','4013','ARGELIA','I',GETDATE(),'CARGA_SQL');

/*
=========================================================================
Tabla Ciudad
tb_ciudad
=========================================================================
*/

/*=====================================================================
Nombre: sp_buscar_ciudad
Objetivo: Buscar el registro en la tabla tb_depto segun criterios
Fecha Creacion: 05/10/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_buscar_ciudad
@codPais varchar (10),
@codDepto varchar (10),
@codCiudad varchar (10)
as
select
A.id_ciudad,
A.id_depto,
A.id_pais,
B.cod_depto,
C.cod_pais,
A.cod_ciudad,
A.nom_ciu,
A.est_ciu,
A.fec_act,
A.obs_gen
from
	tb_ciudad A (nolock) inner join tb_depto B (nolock)
		on B.id_depto = A.id_depto
		inner join tb_pais C (nolock)
		on B.id_pais = A.id_pais
where
B.cod_depto = @codDepto and
C.cod_pais = @codPais and
A.cod_ciudad = @codCiudad

/*=====================================================================
Nombre: sp_guardar_ciudad
Objetivo: Insertar registros en la tabla tb_ciudad
Fecha Creacion: 05/10/2019
Autor: chcgutierrez
=======================================================================*/
create proc sp_guardar_ciudad
@idDepto varchar (10),
@idPais varchar (10),
@codCiudad varchar (10),
@nomCiudad varchar (100),
@estCiudad varchar (10),
@obsGen text
as
insert into tb_ciudad (id_depto, id_pais, cod_ciudad, nom_ciu, est_ciu, fec_act, obs_gen)
	values(@idDepto, @idPais, @codCiudad, @nomCiudad, @estCiudad, getdate(), @obsGen)

/*=====================================================================
Nombre: sp_editar_ciudad
Objetivo: Actualizar el registro en la tabla tb_ciudad segun criterio
Fecha Creacion: 10/05/2019
Autor: chcgutierrez
=======================================================================*/
create procedure sp_editar_ciudad
@codDepto varchar (10),
@codPais varchar (10),
@codCiudad varchar (10),
@nomCiudad varchar (100),
@estCiudad varchar (10),
@obsCiudad text
as
update tb_ciudad set nom_ciu = @nomCiudad,
				    est_ciu = @estCiudad,
				    fec_act = getdate(),
				    obs_gen = @obsCiudad where
					id_depto = @codDepto and
				    id_pais = @codPais and
					cod_ciudad = @codCiudad

select * from tb_tipodoc

/*=====================================================================
Nombre: sp_buscar_pais_desc
Objetivo: Buscar el registro en la tabla tb_pais segun criterio
Fecha Creacion: 30/08/2019
Autor: chcgutierrez
=======================================================================*/
alter proc sp_buscar_pais_desc
@desc_pais varchar (150)
as
select
cod_pais,
nom_pais,
est_pais,
obs_gen
from
	tb_pais (nolock)
where
est_pais = 'A' and
nom_pais like '%'+ @desc_pais +'%'

exec sp_buscar_pais_desc 'AR'

select * from tb_pais
