# Desafio-Paschoalotto

Passo-a-passo para executar o projeto

## Instalações Necessárias:
#
| Plugin | 
| ------ | 
| Excel | 
| [![PostgreSQL]()](https://www.postgresql.org/download/) |
| [![ODBC Driver Manager]()](https://www.postgresql.org/ftp/odbc/versions/msi/) | 

![image](https://user-images.githubusercontent.com/20867353/212007080-f1f4cdcb-28a5-4de3-9ab0-18565c68e4b4.png)

## Dependencias Necessárias:
> Nota: Caso for necessario 

#### Registrar a DLL 

>Iniciar o Prompt de comando como administrador e rodar o comando abaixo entre aspas
  
- MSFLXGRD.OCX -> 'regsvr32 C:\Windows\SysWOW64\MSFLXGRD.OCX'

## Execução:
> Nota: Scripts para criação do usuario, banco de dados e tabela para utilização


```
CREATE USER postgres WITH PASSWORD ‘1234‘;

CREATE DATABASE postgres OWNER postgres;

CREATE TABLE public.pokedex (
	id SERIAL primary key ,
	created_at timestamp NOT NULL DEFAULT now(),
	name_pokemon text NOT NULL,	
	Type_1 text NULL,
	Type_2 text NULL,	
	Total int4 NULL DEFAULT 0,
	HP int4 NULL DEFAULT 0,
	Attack int4 NULL DEFAULT 0,
	Defense int4 NULL DEFAULT 0
);
```

### No 'ODBC Driver Manager'

```
DataBase: postgres
Server / Host: localhost
UserName: postgres

Port: 5432
Password: 1234
```

![image](https://user-images.githubusercontent.com/20867353/212017563-ddf3ea7b-7cf2-42ff-9770-c1613606f6e7.png)
![image](https://user-images.githubusercontent.com/20867353/212017856-51e391e8-db74-41d0-9f0c-092124102309.png)
![image](https://user-images.githubusercontent.com/20867353/212018109-cb21e78c-891e-4811-857a-664662be1814.png)
