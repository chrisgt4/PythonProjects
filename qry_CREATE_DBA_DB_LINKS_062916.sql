USE DBMS3;
DROP TABLE dbo.DBA_DB_LINKS;
CREATE TABLE dbo.DBA_DB_LINKS (
DBNAME							VARCHAR(30)
,VERSION						VARCHAR(30)
,OWNER							VARCHAR(30)
,DB_LINK						VARCHAR(128)
,USERNAME						VARCHAR(30)
,HOST							VARCHAR(2000)
,CREATED						DATE
);
