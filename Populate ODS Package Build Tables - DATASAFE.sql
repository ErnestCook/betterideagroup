USE DATASAFE_ODS

--STEP 1------------------------------------------------------------------------------------------
--Insert into [PACKAGE_BUILD].[dbo].[DATASAFE_ODS_PACKAGE_XML]
INSERT INTO PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_XML](XMLData)
SELECT CONVERT(VARCHAR(MAX), BulkColumn) AS BulkColumn
FROM OPENROWSET(BULK 'C:\AutoGeneratedSSISPackages\ODS\DATASAFE_ODS_Template.dtsx', SINGLE_BLOB) AS x;

--STEP 2------------------------------------------------------------------------------------------
--Insert into PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE]
INSERT INTO PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE]
           ([SOURCE_TABLE_NAME]
           ,[TARGET_TABLE_NAME]
		   ,[TASK_1_CODE]
		   ,[TASK_3_CODE]
		   ,[PACKAGE_NAME]
          )
SELECT 'DATASAFE_'+TABLE_NAME AS SOURCE_TABLE_NAME, 
		TABLE_NAME AS TARGET_TABLE_NAME,
        'SELECT DISTINCT BATCH_AS_OF_DATE FROM DATASAFE_' +  +TABLE_NAME + ' WHERE BATCH_AS_OF_DATE IS NOT NULL ORDER BY BATCH_AS_OF_DATE',
		'DELETE FROM STAGE..DATASAFE_' + TABLE_NAME + ' WHERE BATCH_AS_OF_DATE = ?',
		'ETL_LOAD_DATASAFE_ODS_' + TABLE_NAME
FROM    INFORMATION_SCHEMA.TABLES

--STEP 3------------------------------------------------------------------------------------------
--UPDATE TASK_2_CODE ON PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE]
DECLARE @TABLE_ID INT = 1

WHILE @TABLE_ID <= (SELECT MAX(TABLE_ID) FROM PACKAGE_BUILD..DATASAFE_ODS_PACKAGE_CODE)
BEGIN

declare  @targettable NVARCHAR(MAX)

SET @targettable = (SELECT TARGET_TABLE_NAME FROM PACKAGE_BUILD..DATASAFE_ODS_PACKAGE_CODE WHERE TABLE_ID = @TABLE_ID)

declare @return int;
declare @sql Nvarchar(max) = ''
declare @list Nvarchar(max) = '';

SELECT @list = @list + [name] +', '
from sys.columns
where object_id = object_id(@targettable)

SELECT @list = LEFT(@list, LEN(@list) - 1)

-- --------------------------------------------------------------------------------
DECLARE @OUTPUT1 VARCHAR(MAX)

SET @OUTPUT1 = 'INSERT INTO ' + @targettable + ' 
SELECT ' + @list + ' 
FROM ('

SELECT @list = @list + 's.' + [name] +', '
from sys.columns
where object_id = object_id(@targettable)

DECLARE @OUTPUT2 VARCHAR(MAX)

SET @OUTPUT2 = 'MERGE [dbo].[' + @targettable + '] AS T' + ' 
USING (SELECT * FROM STAGE..DATASAFE_' + @targettable + ' WHERE BATCH_AS_OF_DATE = ?) as S'

-- Get the join columns ----------------------------------------------------------
SET @list = ''
select     @list = @list + 'T.[' + c.COLUMN_NAME + '] = S.[' +  c.COLUMN_NAME + '] AND '
from     INFORMATION_SCHEMA.TABLE_CONSTRAINTS pk ,
    INFORMATION_SCHEMA.KEY_COLUMN_USAGE c
where     pk.TABLE_NAME = @targettable
and    CONSTRAINT_TYPE = 'PRIMARY KEY'
and    c.TABLE_NAME = pk.TABLE_NAME
and    c.CONSTRAINT_NAME = pk.CONSTRAINT_NAME
and c.COLUMN_NAME <> 'BATCH_AS_OF_DATE'

SELECT @list =  LEFT(@list, LEN(@list) -3)

DECLARE @OUTPUT3 VARCHAR(MAX)
SET @OUTPUT3 = ' ON ( ' + @list + ' AND T.CURRENT_FLAG =1 )' + ' WHEN NOT MATCHED BY TARGET THEN'

-- WHEN NOT MATCHED BY TARGET ------------------------------------------------

-- get the values list
SET @list = ''

SELECT @list = @list + '' +[name] +', '
FROM sys.columns
WHERE object_id = OBJECT_ID(@targettable)

SELECT @list = LEFT(@list, LEN(@list) - 1)

SELECT @list = REPLACE(@list,'DW_UPDATE_DATE','GETDATE()')

SELECT @list = REPLACE(@list,'DW_INSERT_DATE','GETDATE()')

SELECT @list = REPLACE(@list,'CURRENT_FLAG','1')
SELECT @list = REPLACE(@list,'PROCESS_FLAG','1')


DECLARE @OUTPUT4 VARCHAR(MAX)
SET @OUTPUT4 =  '  INSERT VALUES(' + @list +  ')' + 
'WHEN MATCHED AND (T.HASHBYTES_VALUE != S.HASHBYTES_VALUE) 
	THEN UPDATE SET T.CURRENT_FLAG = 0, T.DW_UPDATE_DATE = GETDATE()'

--get output list
SET @list = ''

SELECT @list = @list + 'S.' +[name] +', '
FROM sys.columns
WHERE object_id = OBJECT_ID(@targettable)

SELECT @list = LEFT(@list, LEN(@list) - 1)
SELECT @list = REPLACE(@list,'s.DW_UPDATE_DATE','GETDATE() DW_UPDATE_DATE')

SELECT @list = REPLACE(@list,'s.DW_INSERT_DATE','GETDATE() DW_INSERT_DATE')

SELECT @list = REPLACE(@list,'s.CURRENT_FLAG','1 CURRENT_FLAG')
SELECT @list = REPLACE(@list,'s.PROCESS_FLAG','1 PROCESS_FLAG')

DECLARE @OUTPUT5 VARCHAR(MAX)

SET @OUTPUT5= 'OUTPUT $Action Action_Out, ' +@list + ')
 AS MERGE_OUT
 WHERE MERGE_OUT.Action_Out= ' + '''UPDATE''' + ';'

UPDATE PACKAGE_BUILD..DATASAFE_ODS_PACKAGE_CODE
SET TASK_2_CODE = @OUTPUT1 + @OUTPUT2 + @OUTPUT3 + @OUTPUT4 + @OUTPUT5
WHERE TABLE_ID = @TABLE_ID 

SET @TABLE_ID += 1

END

--STEP 4------------------------------------------------------------------------------------------

--Update XML_CODE column on PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE] WITH XMLDATA VALUE from DATASAFE_ODS_PACKAGE_XML table
UPDATE PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE]
SET XML_CODE = X.XMLData
FROM PACKAGE_BUILD..DATASAFE_ODS_PACKAGE_XML X

--Update XML_CODE on DATASAFE_ODS_PACKAGE_CODE with TASK_1_CODE
UPDATE PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE]
SET XML_CODE = P.UPDATED_XML
FROM PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE] PC
	INNER JOIN (SELECT TABLE_ID, I.XML_CODE, REPLACE(I.XML_CODE, 'SELECT DISTINCT CAST(GETDATE() AS DATE) AS DATE_PART',I.TASK_1_CODE) UPDATED_XML
					FROM PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE] I
				) P ON PC.TABLE_ID = P.TABLE_ID

--Update XML_CODE on DATASAFE_ODS_PACKAGE_CODE with TASK_2_CODE
UPDATE PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE]
SET XML_CODE = P.UPDATED_XML
FROM PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE] PC
	INNER JOIN (SELECT TABLE_ID, I.XML_CODE, REPLACE(I.XML_CODE, 'SELECT ?',I.TASK_2_CODE) UPDATED_XML
					FROM PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE] I
				) P ON PC.TABLE_ID = P.TABLE_ID

--Update XML_CODE on DATASAFE_ODS_PACKAGE_CODE with TASK_3_CODE
UPDATE PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE]
SET XML_CODE = P.UPDATED_XML
FROM PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE] PC
	INNER JOIN (SELECT TABLE_ID, I.XML_CODE, REPLACE(I.XML_CODE, 'SELECT 3, ?',I.TASK_3_CODE) UPDATED_XML
					FROM PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE] I
				) P ON PC.TABLE_ID = P.TABLE_ID

--Update XML_CODE on DATASAFE_ODS_PACKAGE_CODE to change default package name to correct name based on the table name
UPDATE PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE]
SET XML_CODE = P.UPDATED_XML
FROM PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE] PC
	INNER JOIN (SELECT TABLE_ID, I.XML_CODE, REPLACE(I.XML_CODE, 'ETL_ODS_LOAD_TEMPLATE',I.PACKAGE_NAME) UPDATED_XML
					FROM PACKAGE_BUILD..[DATASAFE_ODS_PACKAGE_CODE] I
) P ON PC.TABLE_ID = P.TABLE_ID

--Generate BCP Script
SELECT 'bcp "SELECT [XML_CODE] FROM PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE] WHERE TARGET_TABLE_NAME = '''  + T.TABLE_NAME 
		+ '''" queryout C:\AutoGeneratedSSISPackages\ODS\' + T.PACKAGE_NAME + '.dtsx -c -T -S SQL-M360 '
FROM (SELECT TARGET_TABLE_NAME TABLE_NAME , PACKAGE_NAME FROM PACKAGE_BUILD.[dbo].[DATASAFE_ODS_PACKAGE_CODE]) T
