SELECT DISTINCT na.DBName
	, na.VERSION
	, na.USERNAME
	, na.Account_Status
	, na.LOCK_DATE
	, na.EXPIRY_DATE
	, na.Profile
	, na.Created
	, na.AUTHENTICATION_TYPE
	, na.Last_Login
	, na.COMMON
	, na.ORACLE_MAINTAINED
	, utr.GRANTED_ROLE As Role_Name
	, op.PRIVILEGE
INTO Role_Assigned_Obj_Privs
FROM (dbo.New_Accts AS na 
	INNER JOIN dbo.DBA_ROLE_PRIVS AS utr ON na.DBNAME = utr.DBNAME AND na.USERNAME = utr.GRANTEE)
	INNER JOIN dbo.DBA_TAB_PRIVS AS op ON utr.GRANTED_ROLE = op.GRANTEE AND utr.DBNAME = op.DBNAME;

