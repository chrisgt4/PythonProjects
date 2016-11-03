USE SOX2016;
DROP TABLE dbo.DBA_TAB_PRIVS;
CREATE TABLE dbo.DBA_TAB_PRIVS (
DBNAME			VARCHAR(30)
,VERSION		VARCHAR(30)
,GRANTEE		VARCHAR(30)
,OWNER			VARCHAR(30)
,TABLE_NAME		VARCHAR(30)
,GRANTOR		VARCHAR(30)
,PRIVILEGE		VARCHAR(40)
,GRANTABLE		VARCHAR(3)
,HIERARCHY		VARCHAR(3)
,COMMON			VARCHAR(3)
,TYPE			VARCHAR(30)	
)
