--use master
--go
--declare @dbName nvarchar(50)
--set @dbName='' --
--declare   @spid   nvarchar(20) 
--declare   cur_lock   cursor   for 
--SELECT DISTINCT request_session_id FROM master.sys.dm_tran_locks WHERE resource_type = 'DATABASE' AND resource_database_id = db_id(@dbName)
--open   cur_lock 
--fetch   cur_lock      into   @spid 
--while   @@fetch_status=0 
--    begin     
--    exec( 'kill '+@spid) 
--    fetch   Next From cur_lock into @spid
--    end     
--close   cur_lock
--deallocate   cur_lock

--IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'')
DROP DATABASE RSDataV2
GO 

create database RSDataV2
GO

--ALTER DATABASE [] COLLATE Chinese_Taiwan_Stroke_CI_AS