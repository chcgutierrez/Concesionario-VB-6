
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

