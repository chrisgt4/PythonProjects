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
	, 'User Privilege' As Role_Name
	, obj.PRIVILEGE
INTO User_Assigned_Obj_Privs
FROM dbo.New_Accts AS na 
	INNER JOIN dbo.DBA_TAB_PRIVS AS obj ON obj.GRANTEE = na.USERNAME AND obj.DBNAME = na.DBNAME

