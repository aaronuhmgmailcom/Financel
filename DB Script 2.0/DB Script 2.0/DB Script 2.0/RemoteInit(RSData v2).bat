@set/p a1=SQL Server:

@set/p a2=Login Name:

@set/p a3=Password:
@set/p a4=Template Folder:

@del "%cd%\log.txt"

@echo Delete RSDataV2 database...

@sqlcmd -S %a1% -U %a2% -P %a3% -i DeleteDatabaseV2.sql -d master -o Temp.log 

@type Temp.log>>log.txt
 @echo Create RSDataV2 database...

@sqlcmd -S %a1% -U %a2% -P %a3% -i RSData-CreateDatabaseV2.sql -d RSDataV2 -o Temp.log 

@type Temp.log>>log.txt 

@echo Import RSDataV2 BaseData...

@sqlcmd -S %a1% -U %a2% -P %a3% -i BaseDataV2.sql -d RSDataV2 -o Temp.log 
@type Temp.log>>log.txt


@echo Initialize Template Folder...

@sqlcmd -S %a1% -U %a2% -P %a3% -Q "Delete dbo.FinTools_Settings" -d RSDataV2 -o Temp.log 
@sqlcmd -S %a1% -U %a2% -P %a3% -Q "insert into dbo.FinTools_Settings(ft_id,ft_folder) values(newid(),'%a4%')" -d RSDataV2 -o Temp.log 


@type Temp.log>>log.txt 







@del Temp.log
@echo Finished!
Pause
	