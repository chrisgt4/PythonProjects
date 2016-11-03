SELECT subq1.* INTO dbo.ALL_PRIVS FROM (
	SELECT * FROM dbo.User_Assigned_Sys_Privs
	UNION ALL
	SELECT * FROM dbo.Role_Assigned_Sys_Privs
	UNION ALL
	SELECT * FROM dbo.User_Assigned_Obj_Privs
	UNION ALL
	SELECT * FROM dbo.Role_Assigned_Obj_Privs
	) AS subq1
ORDER BY subq1.DBNAME, subq1.USERNAME