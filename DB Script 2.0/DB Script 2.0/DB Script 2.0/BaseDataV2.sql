          
----------
--ttt                                                                                                                                                         a3                      b4 a5                                                    b6 a7                      b8 a9                                                    b10 a11         b12 a13  b14 a15         b16 a17         b18 a19                       b20 a21         b22 c1111                     d
------------------------------------------------------------------------------------------------------------------------------------------------------------- ----------------------- -- ----------------------------------------------------- -- ----------------------- -- ----------------------------------------------------- --- ----------- --- ---- --- ----------- --- ----------- --- ------------------------- --- ----------- --- ------------------------- -
Delete [RSDataV2].[dbo].[FinTools_Users]
go

INSERT INTO [RSDataV2].[dbo].[FinTools_Users]
           (
           [FormUserID]
           ,[FormUserPassword]
           ,[WindowsUserID]
           ,[MachineName]
           ,[LoginType]
           ,[SUNUserIP]
           ,[SUNUserID]
           ,[SUNUserPass])
     VALUES('Peter',	'123'	,'rsimpson-PC\rsimpson'	,NULL	,NULL	,NULL	,NULL	,NULL)
GO


INSERT INTO [RSDataV2].[dbo].[FinTools_ProcessesMacros]
           (
           [ProcessMacroName]
           ,[Type])
     VALUES('Post',	'1'	)
GO

INSERT INTO [RSDataV2].[dbo].[FinTools_ProcessesMacros]
           (
           [ProcessMacroName]
           ,[Type])
     VALUES('Transaction update',	'1'	)
GO

INSERT INTO [RSDataV2].[dbo].[FinTools_ProcessesMacros]
           (
           [ProcessMacroName]
           ,[Type])
     VALUES('Re-open template',	'1'	)
GO

INSERT INTO [RSDataV2].[dbo].[FinTools_ProcessesMacros]
           (
           [ProcessMacroName]
           ,[Type])
     VALUES('Save',	'1'	)
GO
--

INSERT INTO [RSDataV2].[dbo].[FinTools_ProcessesMacros]
           (
           [ProcessMacroName]
           ,[Type])
     VALUES('AllocationMarker Update',	'1'	)
GO

INSERT INTO [RSDataV2].[dbo].[FinTools_ProcessesMacros]
           (
           [ProcessMacroName]
           ,[Type])
     VALUES('Create Text File',	'1'	)
GO
go
          
----------
--ttt                                                  a3                                                    b4 c1111       d
------------------------------------------------------ ----------------------------------------------------- -- ----------- -

--
--go
          
----------
--ttt                                                                     a3          b4 a5          b6 a7          b8 c1111       d
------------------------------------------------------------------------- ----------- -- ----------- -- ----------- -- ----------- -
  
--
--go
          
----------
--ttt                                                                                                                                                       a3          b4 a5          b6 a7                      b8 a9                                                                                                                                                                                                                                                                                                                                                                                                                  b10 a11                  b12 a13                                                                                                                                                                                                                                                                b14 a15                       b16 a17         b18 a19         b20 c1111       d
----------------------------------------------------------------------------------------------------------------------------------------------------------- ----------- -- ----------- -- ----------------------- -- ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- --- -------------------- --- ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ --- ------------------------- --- ----------- --- ----------- --- ----------- -
  
--
--go
          
----------
--ttt                                                                                                                                                                                     a3          b4 a5          b6 a7                                                                                                                                                                                                                                                                                                                                                                                                                  b8 a9          b10 a11                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 b12 a13  b14 a15                       b16 a17         b18 a19         b20 a21         b22 a23                       b24 a25         b26 c1111       d
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ----------- -- ----------- -- ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- -- ----------- --- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- --- ---- --- ------------------------- --- ----------- --- ----------- --- ----------- --- ------------------------- --- ----------- --- ----------- -
  
--
--go
          
----------
--ttt                                                                                            a3          b4 a5          b6 a7          b8 a9                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  b10 a11                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 b12 a13                       b14 c1111       d
------------------------------------------------------------------------------------------------ ----------- -- ----------- -- ----------- -- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- --- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- --- ------------------------- --- ----------- -
  
--
--go
          
----------
--ttt                                                                                                                                     a3          b4 a5          b6 a7          b8 a9          b10 a11         b12 a13         b14 a15         b16 a17                       b18 a19         b20 c1111                     d
----------------------------------------------------------------------------------------------------------------------------------------- ----------- -- ----------- -- ----------- -- ----------- --- ----------- --- ----------- --- ----------- --- ------------------------- --- ----------- --- ------------------------- -
  
--
--go
          
----------
--ttt                                                                                                                                                    a3          b4 a5          b6 a7          b8 a9          b10 a11         b12 a13         b14 a15         b16 a17         b18 a19         b20 a21                       b22 a23         b24 c1111                     d
-------------------------------------------------------------------------------------------------------------------------------------------------------- ----------- -- ----------- -- ----------- -- ----------- --- ----------- --- ----------- --- ----------- --- ----------- --- ----------- --- ------------------------- --- ----------- --- ------------------------- -
  
--
--go
          
----------
--ttt                                                                                                                       a3          b4 a5          b6 a7          b8 a9          b10 a11  b12 a13                       b14 a15         b16 a17                       b18 c1111       d
--------------------------------------------------------------------------------------------------------------------------- ----------- -- ----------- -- ----------- -- ----------- --- ---- --- ------------------------- --- ----------- --- ------------------------- --- ----------- -
  
--
--go
          
----------
--ttt                                                                                                                                a3          b4 a5                                                                                                      b6 a7                      b8 a9                      b10 a11                                                   b12 a13         b14 a15         b16 a17                       b18 a19                       b20 a21                       b22 c1111                     d
------------------------------------------------------------------------------------------------------------------------------------ ----------- -- ------------------------------------------------------------------------------------------------------- -- ----------------------- -- ----------------------- --- ----------------------------------------------------- --- ----------- --- ----------- --- ------------------------- --- ------------------------- --- ------------------------- --- ------------------------- -

--
--go
          
----------
--ttt                                                                                                                                                                                a3          b4 a5          b6 a7                                                    b8 a9          b10 a11                                                                                                                                                                                                                                                                b12 a13                                                   b14 a15         b16 a17         b18 a19                                                   b20 a21                                                                                                     b22 a23                                                   b24 a25         b26 a27         b28 a29         b30 c1111                     d
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ ----------- -- ----------- -- ----------------------------------------------------- -- ----------- --- ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ --- ----------------------------------------------------- --- ----------- --- ----------- --- ----------------------------------------------------- --- ------------------------------------------------------------------------------------------------------- --- ----------------------------------------------------- --- ----------- --- ----------- --- ----------- --- ------------------------- -

--
--go
          
----------
--ttt                                                                                            a3          b4 a5                   b6 a7                   b8 a9                   b10 c1111                d
------------------------------------------------------------------------------------------------ ----------- -- -------------------- -- -------------------- -- -------------------- --- -------------------- -
  
--
--go