SELECT new.* 
FROM dbo.Group_To_User_091015 AS new
LEFT OUTER JOIN
	(SELECT DISTINCT gtu.* 
	FROM dbo.Group_To_User_123114 AS gtu
	WHERE gtu.Fully_Qual_Name IN
		('NASDCORP\app_ax_tech',
		'CORP\SQLADMIN',
		'CORP\priv_mssql_p_admin_access',
		'CORP\role_MSSQL_DBA'
		)
		AND
		gtu.IA_Group_Member_Clean <> 'role_MSSQL_DBA'
	) AS old ON old.Fully_Qual_Name = new.Fully_Qual_Name AND old.IA_Group_Member_Clean = new.IA_Group_Member_Clean
WHERE new.Fully_Qual_Name IN
	('NASDCORP\app_ax_tech',
	'CORP\SQLADMIN',
	'CORP\priv_mssql_p_admin_access',
	'CORP\role_MSSQL_DBA'
	)
	AND
	new.IA_Group_Member_Clean <> 'role_MSSQL_DBA'
	AND old.IA_Group_Member_Clean IS NULL


