create database PruebaBD
go
use PruebaBD
go
create table tb_prueba(
ClassID int identity(1,1)primary key not null,
ClassName varchar(100)not null
)
go