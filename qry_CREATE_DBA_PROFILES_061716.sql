USE SOX2016;
DROP TABLE dbo.DBA_PROFILES;
CREATE TABLE dbo.DBA_PROFILES (
DBNAME			VARCHAR(30)
,VERSION		VARCHAR(30)
,PROFILE		VARCHAR(30)
,RESOURCE_NAME	VARCHAR(32)
,RESOURCE_TYPE	VARCHAR(8)
,LIMIT			VARCHAR(40)
,COMMON			VARCHAR(3)
);