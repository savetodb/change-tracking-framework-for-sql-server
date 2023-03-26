-- =============================================
-- Change Tracking Framework for Microsoft SQL Server
-- Version 4.2, January 9, 2023
--
-- Copyright 2017-2023 Gartle LLC
--
-- License: MIT
-- =============================================

IF OBJECT_ID('logs.xl_actions_drop_all_change_tracking_triggers', 'P') IS NOT NULL
EXEC logs.xl_actions_drop_all_change_tracking_triggers NULL, 1
GO
IF OBJECT_ID('[logs].[usp_translations]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[usp_translations];
GO
IF OBJECT_ID('[logs].[usp_translations_change]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[usp_translations_change];
GO
IF OBJECT_ID('[logs].[xl_actions_clear_logs]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_clear_logs];
GO
IF OBJECT_ID('[logs].[xl_actions_create_change_tracking_triggers]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_create_change_tracking_triggers];
GO
IF OBJECT_ID('[logs].[xl_actions_drop_all_change_tracking_triggers]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_drop_all_change_tracking_triggers];
GO
IF OBJECT_ID('[logs].[xl_actions_drop_change_tracking_triggers]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_drop_change_tracking_triggers];
GO
IF OBJECT_ID('[logs].[xl_actions_restore_current_record]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_restore_current_record];
GO
IF OBJECT_ID('[logs].[xl_actions_restore_previous_record]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_restore_previous_record];
GO
IF OBJECT_ID('[logs].[xl_actions_restore_record]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_restore_record];
GO
IF OBJECT_ID('[logs].[xl_actions_select_lookup_id]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_select_lookup_id];
GO
IF OBJECT_ID('[logs].[xl_actions_select_record]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_select_record];
GO
IF OBJECT_ID('[logs].[xl_actions_select_records]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_select_records];
GO
IF OBJECT_ID('[logs].[xl_actions_set_role_permissions]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_actions_set_role_permissions];
GO
IF OBJECT_ID('[logs].[xl_export_settings]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_export_settings];
GO
IF OBJECT_ID('[logs].[xl_import_handlers]', 'P') IS NOT NULL
DROP PROCEDURE [logs].[xl_import_handlers];
GO

IF OBJECT_ID('[logs].[view_administrator_handlers]', 'V') IS NOT NULL
DROP VIEW [logs].[view_administrator_handlers];
GO
IF OBJECT_ID('[logs].[view_captured_objects]', 'V') IS NOT NULL
DROP VIEW [logs].[view_captured_objects];
GO
IF OBJECT_ID('[logs].[view_query_list]', 'V') IS NOT NULL
DROP VIEW [logs].[view_query_list];
GO
IF OBJECT_ID('[logs].[view_translations]', 'V') IS NOT NULL
DROP VIEW [logs].[view_translations];
GO
IF OBJECT_ID('[logs].[view_user_handlers]', 'V') IS NOT NULL
DROP VIEW [logs].[view_user_handlers];
GO

IF OBJECT_ID('[logs].[get_escaped_parameter_name]', 'FN') IS NOT NULL
DROP FUNCTION [logs].[get_escaped_parameter_name];
GO
IF OBJECT_ID('[logs].[get_translated_string]', 'FN') IS NOT NULL
DROP FUNCTION [logs].[get_translated_string];
GO
IF OBJECT_ID('[logs].[get_unescaped_parameter_name]', 'FN') IS NOT NULL
DROP FUNCTION [logs].[get_unescaped_parameter_name];
GO

IF OBJECT_ID('[logs].[base_tables]', 'U') IS NOT NULL
DROP TABLE [logs].[base_tables];
GO
IF OBJECT_ID('[logs].[change_logs]', 'U') IS NOT NULL
DROP TABLE [logs].[change_logs];
GO
IF OBJECT_ID('[logs].[formats]', 'U') IS NOT NULL
DROP TABLE [logs].[formats];
GO
IF OBJECT_ID('[logs].[handlers]', 'U') IS NOT NULL
DROP TABLE [logs].[handlers];
GO
IF OBJECT_ID('[logs].[tables]', 'U') IS NOT NULL
DROP TABLE [logs].[tables];
GO
IF OBJECT_ID('[logs].[translations]', 'U') IS NOT NULL
DROP TABLE [logs].[translations];
GO
IF OBJECT_ID('[logs].[workbooks]', 'U') IS NOT NULL
DROP TABLE [logs].[workbooks];
GO


DECLARE @sql nvarchar(max) = ''

SELECT
    @sql = @sql + 'ALTER ROLE ' + QUOTENAME(r.name) + ' DROP MEMBER ' + QUOTENAME(m.name) + ';' + CHAR(13) + CHAR(10)
FROM
    sys.database_role_members rm
    INNER JOIN sys.database_principals r ON r.principal_id = rm.role_principal_id
    INNER JOIN sys.database_principals m ON m.principal_id = rm.member_principal_id
WHERE
    r.name IN ('log_admins', 'log_users')

IF LEN(@sql) > 1
    BEGIN
    EXEC (@sql);
    PRINT @sql
    END
GO

IF DATABASE_PRINCIPAL_ID('log_admins') IS NOT NULL
DROP ROLE [log_admins];
GO
IF DATABASE_PRINCIPAL_ID('log_users') IS NOT NULL
DROP ROLE [log_users];
GO

IF SCHEMA_ID('logs') IS NOT NULL
DROP SCHEMA [logs];
GO


IF DATABASE_PRINCIPAL_ID('log_app') IS NOT NULL
DROP USER log_app
GO

print 'Change Tracking Framework removed';
