create database almCarros;

create table tb_cliente(
id_cliente varchar (50) not null primary key,
tip_doc int not null,
num_doc varchar (50) not null,
nom_cli varchar (100) not null,
ape_cli varchar (100) not null,
sexo_cli varchar (10) not null,
dire_cli varchar (50),
tel_cli varchar (20) not null,
dir_mail varchar (100) not null,
est_cli varchar (10) not null,--activo, inactivo
fec_nace_cli date not null,
pais_cli int not null,
depto_cli int not null,
ciu_cli int not null,
obs_gen text
--Relaciones
constraint fk_cliente_tipodoc foreign key (tip_doc) references tb_tipodoc (id_tipodoc),
constraint fk_cliente_pais foreign key (pais_cli) references tb_pais (id_pais),
constraint fk_cliente_depto foreign key (depto_cli) references tb_depto (id_depto),
constraint fk_cliente_ciudad foreign key (ciu_cli) references tb_ciudad (id_ciudad)
);

create table tb_tipodoc(
id_tipodoc int identity (1001,1) not null primary key,
tipo_doc varchar (10) not null,
des_tip_doc varchar (100) not null,
est_tip_doc varchar (10) not null,
fec_act date not null,
obs_gen text
);

create table tb_pais(
id_pais int identity (101,1) not null primary key,
cod_pais varchar (10) not null,
nom_pais varchar (100) not null,
est_pais varchar (10) not null,
fec_act date not null,
obs_gen text
);

create table tb_depto(
id_depto int identity (101,1) not null primary key,
id_pais int not null,
cod_depto varchar (10) not null,
nom_depto varchar (100) not null,
est_depto varchar (10) not null,
fec_act date not null,
obs_gen text
--Relaciones
constraint fk_depto_pais foreign key (id_pais) references tb_pais (id_pais)
);

create table tb_ciudad(
id_ciudad int identity (101,1) not null primary key,
id_depto int not null,
id_pais int not null,
cod_ciudad varchar (10) not null,
nom_ciu varchar (100) not null,
est_ciu varchar (10) not null,
fec_act date not null,
obs_gen text
--Relaciones
constraint fk_ciudad_depto foreign key (id_depto) references tb_depto (id_depto),
constraint fk_ciudad_pais foreign key (id_pais) references tb_pais (id_pais)
);

create table vehiculo(
idvehiculo int identity (2001,1) not null primary key,
tipoveh varchar (50) not null,
marcaveh varchar (50) not null,
modeloveh varchar (10) not null,
colorveh varchar (50) not null,
lineaveh varchar (50) not null,
estadoveh varchar (10) not null,
regsisveh date not null,
comentveh text
);

create table mtipov(
idtipov int identity (101,1) not null primary key,
desctipov varchar (100) not null,
estadotipov varchar (10) not null,
regsistipov date not null,
obstipov text
);

create table mmarcav(
idmarcav int identity (101,1) not null primary key,
descmarcav varchar (100) not null,
estadomarcav varchar (10) not null,
regsismarcav date not null,
obsmarcav text
);
