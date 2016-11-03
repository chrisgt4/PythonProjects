SELECT na.DBName
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
	, 'User Privilege' AS Role_Name
	, sp.PRIVILEGE
INTO User_Assigned_Sys_Privs
FROM dbo.New_Accts AS na 
	INNER JOIN dbo.DBA_SYS_PRIVS AS sp ON na.UserName = sp.GRANTEE AND na.DBNAME = sp.DBNAME;