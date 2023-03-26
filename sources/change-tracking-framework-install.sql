-- =============================================
-- Change Tracking Framework for Microsoft SQL Server
-- Version 4.2, January 9, 2023
--
-- Copyright 2017-2023 Gartle LLC
--
-- License: MIT
-- =============================================

SET NOCOUNT ON
GO

CREATE USER log_app WITHOUT LOGIN;
GO

CREATE SCHEMA logs AUTHORIZATION log_app;
GO

CREATE ROLE log_admins AUTHORIZATION log_app;
GO
CREATE ROLE log_users AUTHORIZATION log_app;
GO

CREATE TABLE logs.base_tables (
    ID int IDENTITY(1,1) NOT NULL
    , OBJECT_SCHEMA nvarchar(128) NOT NULL
    , OBJECT_NAME nvarchar(128) NOT NULL
    , BASE_TABLE_SCHEMA nvarchar(128) NOT NULL
    , BASE_TABLE_NAME nvarchar(128) NOT NULL
    , CONSTRAINT PK_base_tables PRIMARY KEY (ID)
    , CONSTRAINT IX_base_tables UNIQUE (OBJECT_SCHEMA, OBJECT_NAME)
);
GO

CREATE INDEX IX_base_tables_base_table_schema_base_table_name_logs ON logs.base_tables (BASE_TABLE_SCHEMA, BASE_TABLE_NAME);
GO

CREATE TABLE logs.change_logs (
    change_id int IDENTITY(1,1) NOT NULL
    , object_id int NOT NULL
    , id int NULL
    , keys nvarchar(445) NULL
    , inserted xml NULL
    , deleted xml NULL
    , change_type tinyint NOT NULL
    , change_date datetime NOT NULL
    , change_user nvarchar(128) NOT NULL
    , CONSTRAINT PK_change PRIMARY KEY (change_id)
);
GO

CREATE INDEX IX_change_logs_object_id_id_logs ON logs.change_logs (change_id, id);
GO

CREATE INDEX IX_change_logs_object_id_keys_logs ON logs.change_logs (object_id, keys);
GO

CREATE TABLE logs.formats (
    ID int IDENTITY(1,1) NOT NULL
    , TABLE_SCHEMA nvarchar(128) NOT NULL
    , TABLE_NAME nvarchar(128) NOT NULL
    , TABLE_EXCEL_FORMAT_XML xml NULL
    , APP nvarchar(50) NULL
    , CONSTRAINT PK_formats PRIMARY KEY (ID)
    , CONSTRAINT IX_formats UNIQUE (TABLE_SCHEMA, TABLE_NAME, APP)
);
GO

CREATE TABLE logs.handlers (
    ID int IDENTITY(1,1) NOT NULL
    , TABLE_SCHEMA nvarchar(20) NOT NULL
    , TABLE_NAME nvarchar(128) NOT NULL
    , COLUMN_NAME nvarchar(128) NULL
    , EVENT_NAME varchar(25) NOT NULL
    , HANDLER_SCHEMA nvarchar(20) NULL
    , HANDLER_NAME nvarchar(128) NULL
    , HANDLER_TYPE nvarchar(25) NULL
    , HANDLER_CODE nvarchar(max) NULL
    , TARGET_WORKSHEET nvarchar(128) NULL
    , MENU_ORDER int NULL
    , EDIT_PARAMETERS bit NULL
    , CONSTRAINT PK_handlers PRIMARY KEY (ID)
    , CONSTRAINT IX_handlers UNIQUE (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME)
);
GO

CREATE TABLE logs.tables (
    object_id int NOT NULL
    , TABLE_SCHEMA nvarchar(128) NOT NULL
    , TABLE_NAME nvarchar(128) NOT NULL
    , CONSTRAINT PK_tables PRIMARY KEY (object_id)
);
GO

CREATE INDEX IX_tables_logs ON logs.tables (TABLE_SCHEMA, TABLE_NAME);
GO

CREATE TABLE logs.translations (
    ID int IDENTITY(1,1) NOT NULL
    , TABLE_SCHEMA nvarchar(128) NOT NULL
    , TABLE_NAME nvarchar(128) NOT NULL
    , COLUMN_NAME nvarchar(128) NULL
    , LANGUAGE_NAME varchar(10) NOT NULL
    , TRANSLATED_NAME nvarchar(128) NULL
    , TRANSLATED_DESC nvarchar(255) NULL
    , TRANSLATED_COMMENT nvarchar(4000) NULL
    , CONSTRAINT PK_translations PRIMARY KEY (ID)
    , CONSTRAINT IX_translations UNIQUE (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME)
);
GO

CREATE TABLE logs.workbooks (
    ID int IDENTITY(1,1) NOT NULL
    , NAME nvarchar(128) NOT NULL
    , TEMPLATE nvarchar(255) NULL
    , DEFINITION nvarchar(max) NOT NULL
    , TABLE_SCHEMA nvarchar(128) NULL
    , CONSTRAINT PK_workbooks PRIMARY KEY (ID)
    , CONSTRAINT IX_workbooks UNIQUE (NAME)
);
GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Returns the escaped parameter name
-- =============================================

CREATE FUNCTION [logs].[get_escaped_parameter_name]
(
    @name nvarchar(128) = NULL
)
RETURNS nvarchar(255)
AS
BEGIN

RETURN
    REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
    REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
    REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
    REPLACE(REPLACE(@name
    , ' ', '_x0020_'), '!', '_x0021_'), '"', '_x0022_'), '#', '_x0023_'), '$', '_x0024_')
    , '%', '_x0025_'), '&', '_x0026_'), '''', '_x0027_'), '(', '_x0028_'), ')', '_x0029_')
    , '*', '_x002A_'), '+', '_x002B_'), ',', '_x002C_'), '-', '_x002D_'), '.', '_x002E_')
    , '/', '_x002F_'), ':', '_x003A_'), ';', '_x003B_'), '<', '_x003C_'), '=', '_x003D_')
    , '>', '_x003E_'), '?', '_x003F_'), '@', '_x0040_'), '[', '_x005B_'), '\', '_x005C_')
    , ']', '_x005D_'), '^', '_x005E_'), '`', '_x0060_'), '{', '_x007B_'), '|', '_x007C_')
    , '}', '_x007D_'), '~', '_x007E_')

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Returns the translated string
-- =============================================

CREATE FUNCTION [logs].[get_translated_string]
(
    @string nvarchar(128) = NULL
    , @data_language varchar(10) = NULL
)
RETURNS nvarchar(128)
AS
BEGIN

RETURN COALESCE((
    SELECT
        COALESCE(TRANSLATED_DESC, TRANSLATED_NAME)
    FROM
        logs.translations
    WHERE
        TABLE_SCHEMA = 'logs' AND TABLE_NAME = 'strings'
        AND COLUMN_NAME = @string AND LANGUAGE_NAME = COALESCE(@data_language, 'en')
    ), @string)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Returns the unescaped parameter name
-- =============================================

CREATE FUNCTION [logs].[get_unescaped_parameter_name]
(
    @name nvarchar(255) = NULL
)
RETURNS nvarchar(128)
AS
BEGIN

RETURN
    REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
    REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
    REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
    REPLACE(REPLACE(@name
    , '_x0020_', ' '), '_x0021_', '!'), '_x0022_', '"'), '_x0023_', '#'), '_x0024_', '$')
    , '_x0025_', '%'), '_x0026_', '&'), '_x0027_', ''''), '_x0028_', '('), '_x0029_', ')')
    , '_x002A_', '*'), '_x002B_', '+'), '_x002C_', ','), '_x002D_', '-'), '_x002E_', '.')
    , '_x002F_', '/'), '_x003A_', ':'), '_x003B_', ';'), '_x003C_', '<'), '_x003D_', '=')
    , '_x003E_', '>'), '_x003F_', '?'), '_x0040_', '@'), '_x005B_', '['), '_x005C_', '\')
    , '_x005D_', ']'), '_x005E_', '^'), '_x0060_', '`'), '_x007B_', '{'), '_x007C_', '|')
    , '_x007D_', '}'), '_x007E_', '~')

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Generated event handlers for administrators
-- =============================================

CREATE VIEW [logs].[view_administrator_handlers]
AS

SELECT
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , CAST(NULL AS nvarchar(128)) AS COLUMN_NAME
    , 'Actions' AS EVENT_NAME
    , v.handler_schema AS HANDLER_SCHEMA
    , v.handler_name AS HANDLER_NAME
    , v.handler_type AS HANDLER_TYPE
    , CAST(NULL AS nvarchar(max)) AS HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , v.menu_order AS MENU_ORDER
    , CAST(1 AS bit) AS EDIT_PARAMETERS
FROM
    INFORMATION_SCHEMA.TABLES t
    CROSS JOIN (VALUES
        (190, 'logs', 'MenuSeparator190', 'MENUSEPARATOR')
        ) v(menu_order, handler_schema, handler_name, handler_type)
WHERE
    t.TABLE_TYPE = 'BASE TABLE'
    AND NOT t.TABLE_SCHEMA IN ('logs')
    AND EXISTS(SELECT TOP 1 r.ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES r WHERE
            r.ROUTINE_SCHEMA = 'logs' AND r.ROUTINE_NAME = 'xl_actions_create_change_tracking_triggers' AND r.ROUTINE_TYPE = 'PROCEDURE')

UNION ALL
SELECT
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , CAST(NULL AS nvarchar(128)) AS COLUMN_NAME
    , 'Actions' AS EVENT_NAME
    , v.handler_schema AS HANDLER_SCHEMA
    , v.handler_name AS HANDLER_NAME
    , v.handler_type AS HANDLER_TYPE
    , v.handler_code AS HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , v.menu_order AS MENU_ORDER
    , CAST(1 AS bit) AS EDIT_PARAMETERS
FROM
    INFORMATION_SCHEMA.TABLES t
    CROSS JOIN (VALUES
        (191, 'logs', 'Create Change Tracking Triggers', 'CODE', 'EXEC logs.xl_actions_create_change_tracking_triggers NULL, @TableName, @execute_script, @data_language'),
        (192, 'logs', 'Drop Change Tracking Triggers', 'CODE', 'EXEC logs.xl_actions_drop_change_tracking_triggers NULL, @TableName, @execute_script, @data_language')
        ) v(menu_order, handler_schema, handler_name, handler_type, handler_code)
WHERE
    t.TABLE_TYPE = 'BASE TABLE'
    AND NOT t.TABLE_SCHEMA IN ('logs')
    AND EXISTS(SELECT TOP 1 r.ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES r WHERE
            r.ROUTINE_SCHEMA = 'logs' AND r.ROUTINE_NAME = 'xl_actions_create_change_tracking_triggers' AND r.ROUTINE_TYPE = 'PROCEDURE')


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Database objects with change tracking
-- =============================================

CREATE VIEW [logs].[view_captured_objects]
AS

SELECT
    t.[object_id]
    , s.name AS [schema]
    , t.name AS name
    , r1.name AS insert_trigger
    , r2.name AS update_trigger
    , r3.name AS delete_trigger
    , p.name AS select_handler
    , CASE WHEN r1.name IS NOT NULL THEN 1 ELSE 0 END AS has_log_triggers
FROM
    sys.tables t
    INNER JOIN sys.schemas s ON s.[schema_id] = t.[schema_id]
    LEFT OUTER JOIN sys.triggers r1 ON r1.parent_id = t.[object_id] AND r1.name = 'trigger_' + t.name + '_log_insert'
    LEFT OUTER JOIN sys.triggers r2 ON r2.parent_id = t.[object_id] AND r2.name = 'trigger_' + t.name + '_log_update'
    LEFT OUTER JOIN sys.triggers r3 ON r3.parent_id = t.[object_id] AND r3.name = 'trigger_' + t.name + '_log_delete'
    LEFT OUTER JOIN sys.procedures p ON p.name = 'xl_log_' + t.name AND p.[schema_id] = t.[schema_id]
WHERE
    NOT t.name LIKE 'sys%'
    AND NOT s.name IN ('logs')


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Change Tracking Framework objects to connect in Excel
-- =============================================

CREATE VIEW [logs].[view_query_list]
AS

SELECT
    s.name AS TABLE_SCHEMA
    , o.name AS TABLE_NAME
    , CASE o.[type] WHEN 'U' THEN 'BASE TABLE' WHEN 'V' THEN 'VIEW' WHEN 'P' THEN 'PROCEDURE' WHEN 'IF' THEN 'FUNCTION' ELSE o.type_desc END AS TABLE_TYPE
    , CAST(NULL AS nvarchar(max)) AS TABLE_CODE
    , CAST(NULL AS nvarchar(max)) AS INSERT_PROCEDURE
    , CAST(NULL AS nvarchar(max)) AS UPDATE_PROCEDURE
    , CAST(NULL AS nvarchar(max)) AS DELETE_PROCEDURE
    , CAST(NULL AS nvarchar(50)) AS PROCEDURE_TYPE
FROM
    sys.objects o
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
    o.name IN ('view_captured_objects', 'base_tables')
    AND s.name IN ('logs')


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Generated object translation
-- =============================================

CREATE VIEW [logs].[view_translations]
AS

SELECT
    h.HANDLER_SCHEMA AS TABLE_SCHEMA
    , h.HANDLER_NAME AS TABLE_NAME
    , t.COLUMN_NAME
    , t.LANGUAGE_NAME
    , t.TRANSLATED_NAME
    , t.TRANSLATED_DESC
    , t.TRANSLATED_COMMENT
FROM
    logs.handlers h
    CROSS JOIN logs.translations t
WHERE
    t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME =  'xl_actions_select_records' AND t.COLUMN_NAME IS NULL
    AND NOT h.TABLE_SCHEMA = 'logs'
    AND h.TARGET_WORKSHEET = '_taskpane'
    AND h.TABLE_SCHEMA = h.HANDLER_SCHEMA

UNION ALL
SELECT
    h.TABLE_SCHEMA
    , v.name AS TABLE_NAME
    , t.COLUMN_NAME
    , t.LANGUAGE_NAME
    , t.TRANSLATED_NAME
    , t.TRANSLATED_DESC
    , t.TRANSLATED_COMMENT
FROM
    (
        SELECT
            DISTINCT
            h.TABLE_SCHEMA
        FROM
            logs.handlers h
        WHERE
            h.TABLE_SCHEMA = h.HANDLER_SCHEMA
            AND h.TARGET_WORKSHEET = '_taskpane'
            AND h.EVENT_NAME = 'ContextMenu'
            AND h.COLUMN_NAME IS NULL
            AND NOT h.TABLE_SCHEMA = 'logs'
    ) h
    CROSS JOIN (VALUES
        ('All Changes'),
        ('Inserted Rows'),
        ('Updated Rows'),
        ('Deleted Rows')
        ) v(name)
    INNER JOIN logs.translations t ON  t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME = v.name AND t.COLUMN_NAME IS NULL

UNION ALL
SELECT
    h.HANDLER_SCHEMA AS TABLE_SCHEMA
    , h.HANDLER_NAME AS TABLE_NAME
    , t.COLUMN_NAME
    , t.LANGUAGE_NAME
    , t.TRANSLATED_NAME
    , t.TRANSLATED_DESC
    , t.TRANSLATED_COMMENT
FROM
    logs.handlers h
    CROSS JOIN logs.translations t
WHERE
    t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME =  'xl_actions_select_records'
    AND h.TABLE_SCHEMA = h.HANDLER_SCHEMA
    AND h.TARGET_WORKSHEET = '_taskpane'
    AND h.COLUMN_NAME IS NULL
    AND NOT h.TABLE_SCHEMA = 'logs'

UNION ALL
SELECT
    h.TABLE_SCHEMA
    , v.name AS TABLE_NAME
    , t.COLUMN_NAME
    , t.LANGUAGE_NAME
    , t.TRANSLATED_NAME
    , t.TRANSLATED_DESC
    , t.TRANSLATED_COMMENT
FROM
    (
        SELECT
            DISTINCT
            h.TABLE_SCHEMA
        FROM
            logs.handlers h
        WHERE
            h.TABLE_SCHEMA = h.HANDLER_SCHEMA
            AND h.TARGET_WORKSHEET = '_taskpane'
            AND h.EVENT_NAME = 'ContextMenu'
            AND h.COLUMN_NAME IS NULL
            AND NOT h.TABLE_SCHEMA = 'logs'
    ) h
    CROSS JOIN (VALUES
        ('All changes'),
        ('Inserted rows'),
        ('Updated rows'),
        ('Deleted rows')
        ) v(name)
    CROSS JOIN logs.translations t
WHERE
    t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME =  'xl_actions_select_records'
    AND NOT h.TABLE_SCHEMA = 'logs'


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Generated event handlers
-- =============================================

CREATE VIEW [logs].[view_user_handlers]
AS


-- ContextMenu All Changes, Inserted Rows, Updated Rows, Deleted Rows for the underlying ListObject
-- menu_orders: 111, 112, 113, 114
SELECT
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , t.COLUMN_NAME
    , t.EVENT_NAME
    , t.HANDLER_SCHEMA
    , v.handler_name AS HANDLER_NAME
    , 'CODE' AS HANDLER_TYPE
    , 'EXEC ' + QUOTENAME(t.HANDLER_SCHEMA) +  '.' + QUOTENAME(t.HANDLER_NAME) + ' @change_type=' + v.params + ', @data_language=@data_language' AS HANDLER_CODE
    , t.TARGET_WORKSHEET
    , v.menu_order AS MENU_ORDER
    , t.EDIT_PARAMETERS
FROM
    logs.handlers t
    CROSS JOIN (VALUES
        (111, 'All Changes', 'NULL'),
        (112, 'Inserted Rows', '1'),
        (113, 'Updated Rows', '3'),
        (114, 'Deleted Rows', '2')
        ) v(menu_order, handler_name, params)
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'

UNION ALL

-- Menu separators for the underlying ListObject
-- menu_orders: 100 (before Change Log), 110 (before All Changes)
SELECT
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , t.COLUMN_NAME
    , t.EVENT_NAME
    , CAST(NULL AS nvarchar) AS HANDLER_SCHEMA
    , v.handler_name AS HANDLER_NAME
    , 'MENUSEPARATOR' AS HANDLER_TYPE
    , CAST(NULL AS nvarchar) AS HANDLER_CODE
    , CAST(NULL AS nvarchar) AS TARGET_WORKSHEET
    , v.menu_order AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    logs.handlers t
    CROSS JOIN (VALUES
        (100, 'MenuSeparator100'),
        (110, 'MenuSeparator110')
        ) v(menu_order, handler_name)
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'

UNION ALL

-- Source ContextMenu handlers for the inherited objects
-- Complete copies except TABLE_SCHEMA and TABLE_NAME
SELECT
    bt.[OBJECT_SCHEMA] AS TABLE_SCHEMA
    , bt.[OBJECT_NAME] AS TABLE_NAME
    , t.COLUMN_NAME
    , t.EVENT_NAME
    , t.HANDLER_SCHEMA
    , t.HANDLER_NAME
    , t.HANDLER_TYPE
    , t.HANDLER_CODE
    , t.TARGET_WORKSHEET
    , t.MENU_ORDER
    , t.EDIT_PARAMETERS
FROM
    logs.handlers t
    INNER JOIN logs.base_tables bt ON bt.BASE_TABLE_SCHEMA = t.TABLE_SCHEMA AND bt.BASE_TABLE_NAME = t.TABLE_NAME
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'

UNION ALL

-- ContextMenu All Changes, Inserted Rows, Updated Rows, Deleted Rows for the inherited objects
-- menu_orders: 111, 112, 113, 114
SELECT
    bt.[OBJECT_SCHEMA] AS TABLE_SCHEMA
    , bt.[OBJECT_NAME] AS TABLE_NAME
    , t.COLUMN_NAME
    , t.EVENT_NAME
    , t.HANDLER_SCHEMA
    , v.handler_name AS HANDLER_NAME
    , 'CODE' AS HANDLER_TYPE
    , 'EXEC ' + QUOTENAME(t.HANDLER_SCHEMA) +  '.' + QUOTENAME(t.HANDLER_NAME) + ' @change_type=' + v.params + ', @data_language=@data_language' AS HANDLER_CODE
    , t.TARGET_WORKSHEET
    , v.menu_order AS MENU_ORDER
    , t.EDIT_PARAMETERS
FROM
    logs.handlers t
    INNER JOIN logs.base_tables bt ON bt.BASE_TABLE_SCHEMA = t.TABLE_SCHEMA AND bt.BASE_TABLE_NAME = t.TABLE_NAME
    CROSS JOIN (VALUES
        (111, 'All Changes', 'NULL'),
        (112, 'Inserted Rows', '1'),
        (113, 'Updated Rows', '3'),
        (114, 'Deleted Rows', '2')
        ) v(menu_order, handler_name, params)
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'

UNION ALL

-- Menu separators for inherited objects
-- menu_orders: 100 (before Change Log), 110 (before All Changes)
SELECT
    bt.[OBJECT_SCHEMA] AS TABLE_SCHEMA
    , bt.[OBJECT_NAME] AS TABLE_NAME
    , t.COLUMN_NAME
    , t.EVENT_NAME
    , CAST(NULL AS nvarchar) AS HANDLER_SCHEMA
    , v.handler_name AS HANDLER_NAME
    , 'MENUSEPARATOR' AS HANDLER_TYPE
    , CAST(NULL AS nvarchar) AS HANDLER_CODE
    , CAST(NULL AS nvarchar) AS TARGET_WORKSHEET
    , v.menu_order AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    logs.handlers t
    INNER JOIN logs.base_tables bt ON bt.BASE_TABLE_SCHEMA = t.TABLE_SCHEMA AND bt.BASE_TABLE_NAME = t.TABLE_NAME
    CROSS JOIN (VALUES
        (100, 'MenuSeparator100'),
        (110, 'MenuSeparator110')
        ) v(menu_order, handler_name)
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'

UNION ALL

-- SelectionChange for Change Log, All Changes, Inserted Rows, Updated Rows, Deleted Rows in task panes
-- Shows record details
SELECT
    t.TABLE_SCHEMA
    , COALESCE(v.name, t.HANDLER_NAME) AS TABLE_NAME
    , t.COLUMN_NAME
    , 'SelectionChange' AS EVENT_NAME
    , 'logs' AS HANDLER_SCHEMA
    , 'xl_actions_select_record' AS HANDLER_NAME
    , 'PROCEDURE' AS HANDLER_TYPE
    , CAST(NULL AS nvarchar) AS HANDLER_CODE
    , t.TARGET_WORKSHEET
    , CAST(NULL AS int) AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    logs.handlers t
    CROSS JOIN (VALUES
        (NULL),
        ('All Changes'),
        ('Inserted Rows'),
        ('Updated Rows'),
        ('Deleted Rows')
        ) v(name)
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'

UNION ALL

-- ContextMenu for Change Log, All Changes, Inserted Rows, Updated Rows, Deleted Rows in task panes
-- Shows a record from the foreign key table
-- menu_order: 21
SELECT
    t.TABLE_SCHEMA
    , COALESCE(v.name, t.HANDLER_NAME) AS TABLE_NAME
    , t.COLUMN_NAME
    , 'ContextMenu' AS EVENT_NAME
    , 'logs' AS HANDLER_SCHEMA
    , 'xl_actions_select_lookup_id' AS HANDLER_NAME
    , 'PROCEDURE' AS HANDLER_TYPE
    , CAST(NULL AS nvarchar) AS HANDLER_CODE
    , t.TARGET_WORKSHEET
    , 21 AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    logs.handlers t
    CROSS JOIN (VALUES
        (NULL),
        ('All Changes'),
        ('Inserted Rows'),
        ('Updated Rows'),
        ('Deleted Rows')
        ) v(name)
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'

UNION ALL

-- ContextMenu for Change Log, All Changes, Inserted Rows, Updated Rows, Deleted Rows in task panes
-- Restore current record and Restore previous record
-- menu_order: 31, 32
SELECT
    t.TABLE_SCHEMA
    , COALESCE(v.name, t.HANDLER_NAME) AS TABLE_NAME
    , t.COLUMN_NAME
    , 'ContextMenu' AS EVENT_NAME
    , 'logs' AS HANDLER_SCHEMA
    , p.handler_name AS HANDLER_NAME
    , 'PROCEDURE' AS HANDLER_TYPE
    , CAST(NULL AS nvarchar) AS HANDLER_CODE
    , '_reload' AS TARGET_WORKSHEET
    , p.menu_order AS MENU_ORDER
    , CAST(1 AS bit) AS EDIT_PARAMETERS
FROM
    logs.handlers t
    CROSS JOIN (VALUES
        (NULL),
        ('All Changes'),
        ('Inserted Rows'),
        ('Updated Rows'),
        ('Deleted Rows')
        ) v(name)
    CROSS JOIN (VALUES
        (31, 'xl_actions_restore_current_record'),
        (32, 'xl_actions_restore_previous_record')
        ) p(menu_order, handler_name)
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'

UNION ALL

-- Menu separators for Change Log, All Changes, Inserted Rows, Updated Rows, Deleted Rows in task panes
-- menu_orders: 20 (before Lookup ID), 30 (before Restore current record)
SELECT
    t.TABLE_SCHEMA
    , COALESCE(v.name, t.HANDLER_NAME) AS TABLE_NAME
    , t.COLUMN_NAME
    , t.EVENT_NAME
    , CAST(NULL AS nvarchar) AS HANDLER_SCHEMA
    , p.handler_name AS HANDLER_NAME
    , 'MENUSEPARATOR' AS HANDLER_TYPE
    , CAST(NULL AS nvarchar) AS HANDLER_CODE
    , CAST(NULL AS nvarchar) AS TARGET_WORKSHEET
    , p.menu_order AS MENU_ORDER
    , CAST(NULL AS bit) AS EDIT_PARAMETERS
FROM
    logs.handlers t
    CROSS JOIN (VALUES
        (NULL),
        ('All Changes'),
        ('Inserted Rows'),
        ('Updated Rows'),
        ('Deleted Rows')
        ) v(name)
    CROSS JOIN (VALUES
        (20, 'MenuSeparator20'),
        (30, 'MenuSeparator30')
        ) p(menu_order, handler_name)
WHERE
    t.TABLE_SCHEMA = t.HANDLER_SCHEMA
    AND t.TARGET_WORKSHEET = '_taskpane'
    AND t.EVENT_NAME = 'ContextMenu'
    AND t.COLUMN_NAME IS NULL
    AND NOT t.TABLE_SCHEMA = 'logs'


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Selects translations
-- =============================================

CREATE PROCEDURE [logs].[usp_translations]
    @field nvarchar(128) = NULL
AS
BEGIN

SET NOCOUNT ON;

IF @field IS NULL
    SET @field = 'TRANSLATED_NAME'
ELSE IF @field NOT IN ('TRANSLATED_NAME', 'TRANSLATED_DESC', 'TRANSLATED_COMMENT')
    BEGIN
    DECLARE @message nvarchar(max) = N'Invalid column name: %s' + CHAR(13) + CHAR(10)
         + 'Use TRANSLATED_NAME, TRANSLATED_DESC, or TRANSLATED_COMMENT'
    RAISERROR(@message, 11, 0, @field);
    RETURN
    END

DECLARE @sql nvarchar(max)
DECLARE @languages nvarchar(max)

SELECT @languages = STUFF((
    SELECT
        t.name
    FROM
        (
        SELECT
            DISTINCT
            ', [' + t.LANGUAGE_NAME + ']' AS name
            , CASE
                WHEN t.LANGUAGE_NAME = 'en' THEN '1'
                WHEN t.LANGUAGE_NAME = 'fr' THEN '2'
                WHEN t.LANGUAGE_NAME = 'it' THEN '3'
                WHEN t.LANGUAGE_NAME = 'es' THEN '4'
                ELSE t.LANGUAGE_NAME
                END AS sort_order
        FROM
            logs.translations t
        ) t
    ORDER BY
        t.sort_order
    FOR XML PATH(''), TYPE).value('.', 'nvarchar(MAX)'), 1, 2, '')

IF @languages IS NULL SET @languages = '[en]'

SET @sql = 'SELECT
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , t.COLUMN_NAME AS [COLUMN]
    , ' + @languages + '
FROM
    (
        SELECT
            t.TABLE_SCHEMA
            , t.TABLE_NAME
            , t.COLUMN_NAME
            , t.LANGUAGE_NAME
            , t.' + @field + ' AS name
        FROM
            logs.translations t
    ) s PIVOT (
        MAX(name) FOR LANGUAGE_NAME IN (' + @languages + ')
    ) t
ORDER BY
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , t.COLUMN_NAME'

-- PRINT @sql

EXEC (@sql)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Cell change event handler for usp_translations
--
-- @column_name is a column name of the edited cell
-- @cell_value is a new cell value
-- @field is a value of the field parameter of the usp_translations
-- @TABLE_SCHEMA, @TABLE_NAME, and @COLUMN are values of the Excel table columns
-- =============================================

CREATE PROCEDURE [logs].[usp_translations_change]
    @column_name nvarchar(128) = NULL
    , @cell_value nvarchar(max) = NULL
    , @TABLE_SCHEMA nvarchar(128) = NULL
    , @TABLE_NAME nvarchar(128) = NULL
    , @COLUMN nvarchar(128) = NULL
    , @field nvarchar(128) = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @message nvarchar(max)

IF NOT EXISTS(SELECT TOP 1 ID FROM logs.translations WHERE LANGUAGE_NAME = @column_name)
    RETURN

IF @field IS NULL SET @field = 'TRANSLATED_NAME'

IF @field = 'TRANSLATED_NAME'
    UPDATE logs.translations SET TRANSLATED_NAME = @cell_value
        WHERE COALESCE(TABLE_SCHEMA, '') = COALESCE(@TABLE_SCHEMA, '') AND COALESCE(TABLE_NAME, '') = COALESCE(@TABLE_NAME, '')
            AND COALESCE(COLUMN_NAME, '') = COALESCE(@COLUMN, '') AND LANGUAGE_NAME = @column_name
ELSE IF @field = 'TRANSLATED_DESC'
    UPDATE logs.translations SET TRANSLATED_DESC = @cell_value
        WHERE COALESCE(TABLE_SCHEMA, '') = COALESCE(@TABLE_SCHEMA, '') AND COALESCE(TABLE_NAME, '') = COALESCE(@TABLE_NAME, '')
            AND COALESCE(COLUMN_NAME, '') = COALESCE(@COLUMN, '') AND LANGUAGE_NAME = @column_name
ELSE IF @field = 'TRANSLATED_COMMENT'
    UPDATE logs.translations SET TRANSLATED_COMMENT = @cell_value
        WHERE COALESCE(TABLE_SCHEMA, '') = COALESCE(@TABLE_SCHEMA, '') AND COALESCE(TABLE_NAME, '') = COALESCE(@TABLE_NAME, '')
            AND COALESCE(COLUMN_NAME, '') = COALESCE(@COLUMN, '') AND LANGUAGE_NAME = @column_name
ELSE
    BEGIN
    SET @message = N'Invalid column name: %s' + CHAR(13) + CHAR(10)
         + 'Use TRANSLATED_NAME, TRANSLATED_DESC, or TRANSLATED_COMMENT'
    RAISERROR(@message, 11, 0, @field);
    RETURN
    END

IF @@ROWCOUNT > 0 RETURN

IF @cell_value IS NULL RETURN

IF @field = 'TRANSLATED_NAME'
    INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME)
        VALUES (@TABLE_SCHEMA, @TABLE_NAME, @COLUMN, @column_name, @cell_value)
ELSE IF @field = 'TRANSLATED_DESC'
    INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_DESC)
        VALUES (@TABLE_SCHEMA, @TABLE_NAME, @COLUMN, @column_name, @cell_value)
ELSE IF @field = 'TRANSLATED_COMMENT'
    INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_COMMENT)
        VALUES (@TABLE_SCHEMA, @TABLE_NAME, @COLUMN, @column_name, @cell_value)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure clears table logs
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_clear_logs]
    @schema nvarchar(128) = NULL
    , @name nvarchar(128) = NULL
    , @confirm bit = 0
    , @data_language varchar(10) = NULL
AS
BEGIN

IF NOT @confirm = 1
    RETURN

DECLARE @obj int = (SELECT [object_id] FROM sys.tables WHERE name = @name AND [schema_id] = SCHEMA_ID(@schema))

DECLARE @message nvarchar(max)

IF @obj IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'Invalid table', @data_language)
    RAISERROR(@message, 11, 0)
    RETURN
    END

DELETE FROM logs.change_logs WHERE [object_id] = @obj

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure creates triggers to log changes
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_create_change_tracking_triggers]
    @schema nvarchar(128) = NULL
    , @name nvarchar(128) = NULL
    , @execute_script bit = 0
    , @data_language varchar(10) = NULL
WITH EXECUTE AS CALLER
AS
BEGIN

DECLARE @message nvarchar(max)

IF @schema IS NULL AND @name IS NOT NULL AND CHARINDEX('.', @name) > 1
    BEGIN
    SET @schema = LEFT(@name, CHARINDEX('.', @name) - 1)
    SET @name = SUBSTRING(@name, CHARINDEX('.', @name) + 1, LEN(@name))
    END

IF @schema IS NULL OR @name IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'Specify all parameters of the procedure', @data_language)
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF LEFT(@schema, 1) = '[' AND RIGHT(@schema, 1) = ']'
    SET @schema = REPLACE(SUBSTRING(@schema, 2, LEN(@schema) - 2), ']]', ']')

IF LEFT(@name, 1) = '[' AND RIGHT(@name, 1) = ']'
    SET @name = REPLACE(SUBSTRING(@name, 2, LEN(@name) - 2), ']]', ']')

IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = @schema)
    BEGIN
    SET @message = logs.get_translated_string(N'Invalid schema', @data_language)
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = @name)
    BEGIN
    SET @message = logs.get_translated_string(N'Invalid table name', @data_language)
    RAISERROR(@message, 11, 0);
    RETURN
    END

SET NOCOUNT ON

DECLARE @sql1 nvarchar(MAX), @sql2 nvarchar(MAX), @sql3 nvarchar(MAX),
        @sql4 nvarchar(MAX), @sql5 nvarchar(MAX), @sql6 nvarchar(MAX),
        @sql7 nvarchar(MAX), @sql8 nvarchar(MAX), @sql9 nvarchar(MAX),
        @sql10 nvarchar(MAX)

SELECT
    @sql1 ='
IF OBJECT_ID(N'''  + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_insert'), '''', '''''') + ''') IS NOT NULL
    DROP TRIGGER ' +         QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_insert')
    , @sql2 = '
IF OBJECT_ID(N'''  + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_update'), '''', '''''') + ''') IS NOT NULL
    DROP TRIGGER ' +         QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_update')
    , @sql3 = '
IF OBJECT_ID(N'''  + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_delete'), '''', '''''') + ''') IS NOT NULL
    DROP TRIGGER ' +         QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_delete')
    , @sql4 = '
CREATE' + ' TRIGGER ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_insert') + '
    ON ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME(t.name) + '
    WITH EXECUTE AS ''log_app''
    AFTER INSERT
AS
BEGIN

SET NOCOUNT ON

DECLARE @username nvarchar(max)

EXECUTE AS CALLER

SELECT @username = USER_NAME()

REVERT

INSERT INTO logs.change_logs (object_id, ' + t.key_field + ', inserted, deleted, change_type, change_date, change_user)
SELECT
    ' + CAST(t.[object_id] AS nvarchar) + '
    , ' + CASE WHEN t.identity_name IS NOT NULL THEN 'inserted.' + QUOTENAME(t.identity_name)
               ELSE '(SELECT ' + t.keys + ' FROM inserted FOR XML RAW)' END + '
    , (SELECT * FROM inserted FOR XML RAW)
    , NULL
    , 1
    , GETDATE()
    , @username
FROM
    inserted

END'
    , @sql5 = '
CREATE' + ' TRIGGER ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_update') + '
    ON ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME(t.name) + '
    WITH EXECUTE AS ''log_app''
    AFTER UPDATE
AS
BEGIN

SET NOCOUNT ON

DECLARE @username nvarchar(max)

EXECUTE AS CALLER

SELECT @username = USER_NAME()

REVERT

INSERT INTO logs.change_logs (object_id, ' + t.key_field + ', inserted, deleted, change_type, change_date, change_user)
SELECT
    ' + CAST(t.[object_id] AS nvarchar) + '
    , ' + CASE WHEN t.identity_name IS NOT NULL THEN 'inserted.' + QUOTENAME(t.identity_name)
               ELSE '(SELECT ' + t.keys + ' FROM inserted FOR XML RAW)' END + '
    , (SELECT * FROM inserted FOR XML RAW)
    , (SELECT * FROM deleted FOR XML RAW)
    , 3
    , GETDATE()
    , @username
FROM
    inserted

END'
    , @sql6 = '
CREATE' + ' TRIGGER ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_delete') + '
    ON ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME(t.name) + '
    WITH EXECUTE AS ''log_app''
    AFTER DELETE
AS
BEGIN

SET NOCOUNT ON

DECLARE @username nvarchar(max)

EXECUTE AS CALLER

SELECT @username = USER_NAME()

REVERT

INSERT INTO logs.change_logs (object_id, ' + t.key_field + ', inserted, deleted, change_type, change_date, change_user)
SELECT
    ' + CAST(t.[object_id] AS nvarchar) + '
    , ' + CASE WHEN t.identity_name IS NOT NULL THEN 'deleted.' + QUOTENAME(t.identity_name)
               ELSE '(SELECT ' + t.keys + ' FROM deleted FOR XML RAW)' END + '
    , NULL
    , (SELECT * FROM deleted FOR XML RAW)
    , 2
    , GETDATE()
    , @username
FROM
    deleted

END'
    , @sql7 = '
IF OBJECT_ID(N'''    + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME('xl_log_' + t.name), '''', '''''') + ''') IS NOT NULL
    DROP PROCEDURE ' +         QUOTENAME(t.[schema]) + '.' + QUOTENAME('xl_log_' + t.name)
    , @sql8 = '
CREATE' + ' PROCEDURE ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME('xl_log_' + t.name) + '
    ' + CASE WHEN t.identity_name IS NOT NULL THEN '@' + logs.get_escaped_parameter_name(t.identity_name) + ' int = NULL' + CHAR(13) + CHAR(10)
             ELSE t.params
             END +
'    , @change_type tinyint = NULL
    , @data_language varchar(10) = NULL
AS
BEGIN
' + CASE WHEN t.identity_name IS NOT NULL THEN '' ELSE '
DECLARE @keys nvarchar(445)

SET @keys = (SELECT ' + t.set_keys + ' FOR XML RAW)
' END + '
EXEC logs.xl_actions_select_records N''' + REPLACE(t.[schema], '''','''''') + ''', N''' + REPLACE(t.name, '''','''''') +''', '
    + CASE WHEN t.identity_name IS NOT NULL THEN '@' + logs.get_escaped_parameter_name(t.identity_name) +', NULL' ELSE 'NULL, @keys' END + ', @change_type, @data_language

END'
    , @sql9 = 'EXEC logs.xl_import_handlers '
            +   'N''' + REPLACE(t.[schema], '''','''''') + ''''
            + ', N''' + REPLACE(t.name, '''','''''') + ''''
            + ', NULL'
            + ', N''ContextMenu'''
            + ', N''' + REPLACE(t.[schema], '''','''''') + ''''
            + ', N''xl_log_' + REPLACE(t.name, '''','''''') + ''''
            + ', N''PROCEDURE'''
            + ', NULL'
            + ', N''_taskpane'''
            + ', 101'
            + ', NULL'
    , @sql10 = 'INSERT INTO logs.tables (object_id, TABLE_SCHEMA, TABLE_NAME) VALUES ('
            + CAST(t.[object_id] AS nvarchar(128))
            + ', N''' + REPLACE(t.[schema], '''','''''') + ''''
            + ', N''' + REPLACE(t.name, '''','''''') + ''''
            + ')'
FROM
    (
        SELECT
            t.[object_id]
            , OBJECT_SCHEMA_NAME(t.[object_id]) AS [schema]
            , t.name AS name
            , c.name AS identity_name
            , CASE WHEN c.name IS NOT NULL THEN 'id' ELSE 'keys' END AS key_field
            , (SELECT STUFF((
                    SELECT
                        ', ' + QUOTENAME(c.name) + ''
                    FROM
                        sys.index_columns ic
                        INNER JOIN sys.columns c ON c.[object_id] = ic.[object_id] AND c.column_id = ic.column_id
                    WHERE
                        c.[object_id] = k.parent_object_id
                        AND ic.index_id = k.unique_index_id
                    ORDER BY
                        ic.index_id
                    FOR XML PATH(''), TYPE
                    ).value('.', 'nvarchar(MAX)'), 1, 2, '')
                FROM
                    sys.key_constraints k
                WHERE
                    k.parent_object_id = t.[object_id]
                    AND k.[type] = 'PK'
                    ) AS keys
            , (SELECT STUFF((
                    SELECT
                        '    , @' + logs.get_escaped_parameter_name(c.name) + ' ' + tp.name
                        + CASE WHEN tp.name IN ('varchar', 'char', 'varbinary', 'binary', 'text')
                                THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length AS VARCHAR(5)) END + ')'
                                WHEN tp.name IN ('nvarchar', 'nchar', 'ntext')
                                THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length / 2 AS VARCHAR(5)) END + ')'
                                WHEN tp.name IN ('datetime2', 'time2', 'datetimeoffset')
                                THEN '(' + CAST(c.scale AS VARCHAR(5)) + ')'
                                WHEN tp.name = 'decimal'
                                THEN '(' + CAST(c.[precision] AS VARCHAR(5)) + ',' + CAST(c.scale AS VARCHAR(5)) + ')'
                            ELSE ''
                        END
                        + ' = NULL' + CHAR(13) + CHAR(10)
                    FROM
                        sys.index_columns ic
                        INNER JOIN sys.columns c ON c.[object_id] = ic.[object_id] AND c.column_id = ic.column_id
                        INNER JOIN sys.types tp ON c.user_type_id = tp.user_type_id
                    WHERE
                        c.[object_id] = k.parent_object_id
                        AND ic.index_id = k.unique_index_id
                    ORDER BY
                        ic.index_id
                    FOR XML PATH(''), TYPE
                    ).value('.', 'nvarchar(MAX)'), 1, 6, '')
                FROM
                    sys.key_constraints k
                WHERE
                    k.parent_object_id = t.[object_id]
                    AND k.[type] = 'PK'
                    ) AS params
            , (SELECT STUFF((
                    SELECT
                        ', @' + logs.get_escaped_parameter_name(c.name) + ' AS ' + QUOTENAME(c.name) + ''
                    FROM
                        sys.index_columns ic
                        INNER JOIN sys.columns c ON c.[object_id] = ic.[object_id] AND c.column_id = ic.column_id
                    WHERE
                        c.[object_id] = k.parent_object_id
                        AND ic.index_id = k.unique_index_id
                    ORDER BY
                        ic.index_id
                    FOR XML PATH(''), TYPE
                    ).value('.', 'nvarchar(MAX)'), 1, 2, '')
                FROM
                    sys.key_constraints k
                WHERE
                    k.parent_object_id = t.[object_id]
                    AND k.[type] = 'PK'
                    ) AS set_keys
        FROM
            sys.tables t
            LEFT OUTER JOIN sys.columns c ON c.[object_id] = t.[object_id] AND c.is_identity = 1
        WHERE
            t.[schema_id] = SCHEMA_ID(@schema)
            AND t.name = @name
    ) t

DECLARE @br nvarchar(100) = ';' + CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10) + 'GO' + CHAR(13) + CHAR(10)

-- PRINT @sql1 + @br + @sql2 + @br + @sql3 + @br + @sql4 + @br + @sql5 + @br + @sql6 + @br + @sql7 + @br + @sql8 + @br + @sql9 + @br

IF @execute_script = 1
    BEGIN
    EXEC (@sql1)
    EXEC (@sql2)
    EXEC (@sql3)
    EXEC (@sql4)
    EXEC (@sql5)
    EXEC (@sql6)
    EXEC (@sql7)
    EXEC (@sql8)
    EXEC (@sql9)
    EXEC (@sql10)
    END
ELSE
    BEGIN
    RAISERROR(@sql1, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql2, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql3, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql4, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql5, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql6, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql7, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql8, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql9, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql10, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    --PRINT @sql1
    --PRINT 'GO'
    --PRINT @sql2
    --PRINT 'GO'
    --PRINT @sql3
    --PRINT 'GO'
    --PRINT @sql4
    --PRINT 'GO'
    --PRINT @sql5
    --PRINT 'GO'
    --PRINT @sql6
    --PRINT 'GO'
    --PRINT @sql7
    --PRINT 'GO'
    --PRINT @sql8
    --PRINT 'GO'
    --PRINT @sql9
    --PRINT 'GO'
    --PRINT @sql10
    --PRINT 'GO'
    END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure drops all triggers to log changes
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_drop_all_change_tracking_triggers]
    @schema nvarchar(max) = NULL
    , @execute_script bit = 0
WITH EXECUTE AS CALLER
AS
BEGIN

SET NOCOUNT ON


DECLARE @sql nvarchar(max)

SET @sql = 'DELETE FROM logs.change_logs' + CASE WHEN @schema IS NULL THEN ''
        ELSE ' WHERE OBJECT_SCHEMA_NAME([object_id]) '
            + CASE WHEN CHARINDEX('%', @schema)>0 THEN 'LIKE' ELSE '=' END
            + ' ''' + REPLACE(@schema, '''', '''''') + '''' END + ';' + CHAR(13) + CHAR(10)
         + 'DELETE FROM logs.handlers'
         + CASE WHEN @schema IS NULL THEN ' WHERE NOT TABLE_SCHEMA = ''logs'''
        ELSE ' WHERE TABLE_SCHEMA '
            + CASE WHEN CHARINDEX('%', @schema)>0 THEN 'LIKE' ELSE '=' END
            + ' ''' + REPLACE(@schema, '''', '''''') + '''' END
            + ';' + CHAR(13) + CHAR(10)

SELECT
    @sql = @sql + 'DROP TRIGGER ' + QUOTENAME(s.name) + '.' + QUOTENAME(r.name) + ';' + CHAR(13) + CHAR(10)
FROM
    sys.tables t
    INNER JOIN sys.schemas s ON s.[schema_id] = t.[schema_id]
    INNER JOIN sys.triggers r ON r.parent_id = t.[object_id]
WHERE
    NOT t.name LIKE 'sys%'
    AND r.name LIKE 'trigger_' + t.name + '_log_%'
    AND s.name LIKE COALESCE(@schema, s.name)

SELECT
    @sql = @sql + 'DROP PROCEDURE ' + QUOTENAME(s.name) + '.' + QUOTENAME(p.name) + ';' + CHAR(13) + CHAR(10)
FROM
    sys.tables t
    INNER JOIN sys.schemas s ON s.[schema_id] = t.[schema_id]
    INNER JOIN sys.procedures p ON p.name = 'xl_log_' + t.name AND p.[schema_id] = t.[schema_id]
WHERE
    NOT t.name LIKE 'sys%'
    AND s.name LIKE COALESCE(@schema, s.name)

SET @sql = @sql + 'DELETE FROM logs.tables' + CASE WHEN @schema IS NULL THEN ''
        ELSE ' WHERE OBJECT_SCHEMA_NAME([object_id]) '
            + CASE WHEN CHARINDEX('%', @schema)>0 THEN 'LIKE' ELSE '=' END
            + ' ''' + REPLACE(@schema, '''', '''''') + ''''
            + ' OR TABLE_SCHEMA '
            + CASE WHEN CHARINDEX('%', @schema)>0 THEN 'LIKE' ELSE '=' END
            + ' ''' + REPLACE(@schema, '''', '''''') + '''' END
            + ';' + CHAR(13) + CHAR(10)

-- PRINT @sql

IF @execute_script = 1
    EXEC (@sql)
ELSE
    PRINT @sql

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure drops triggers to log changes
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_drop_change_tracking_triggers]
    @schema nvarchar(128) = NULL
    , @name nvarchar(128) = NULL
    , @execute_script bit = 0
    , @data_language varchar(10) = NULL
WITH EXECUTE AS CALLER
AS
BEGIN

DECLARE @message nvarchar(max)

IF @schema IS NULL AND @name IS NOT NULL AND CHARINDEX('.', @name) > 1
    BEGIN
    SET @schema = LEFT(@name, CHARINDEX('.', @name) - 1)
    SET @name = SUBSTRING(@name, CHARINDEX('.', @name) + 1, LEN(@name))
    END

IF @schema IS NULL OR @name IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'Specify all parameters of the procedure', @data_language)
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF LEFT(@schema, 1) = '[' AND RIGHT(@schema, 1) = ']'
    SET @schema = REPLACE(SUBSTRING(@schema, 2, LEN(@schema) - 2), ']]', ']')

IF LEFT(@name, 1) = '[' AND RIGHT(@name, 1) = ']'
    SET @name = REPLACE(SUBSTRING(@name, 2, LEN(@name) - 2), ']]', ']')

IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = @schema)
    BEGIN
    SET @message = logs.get_translated_string(N'Invalid schema', @data_language)
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = @name)
    BEGIN
    SET @message = logs.get_translated_string(N'Invalid table name', @data_language)
    RAISERROR(@message, 11, 0);
    RETURN
    END

SET NOCOUNT ON

DECLARE @obj int
DECLARE @sql1 nvarchar(MAX), @sql2 nvarchar(MAX), @sql3 nvarchar(MAX),
        @sql7 nvarchar(MAX), @sql8 nvarchar(MAX), @sql9 nvarchar(MAX),
        @sql10 nvarchar(MAX)

SELECT
    @obj = t.[object_id]
    , @sql1 ='
IF OBJECT_ID(N'''  + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_insert'), '''', '''''') + ''') IS NOT NULL
    DROP TRIGGER ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_insert')
    , @sql2 = '
IF OBJECT_ID(N'''  + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_update'), '''', '''''') + ''') IS NOT NULL
    DROP TRIGGER ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_update')
    , @sql3 = '
IF OBJECT_ID(N'''  + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_delete'), '''', '''''') + ''') IS NOT NULL
    DROP TRIGGER ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME('trigger_' + t.name + '_log_delete')
    , @sql7 = '
IF OBJECT_ID(N'''    + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME('xl_log_' + t.name), '''', '''''') + ''') IS NOT NULL
    DROP PROCEDURE ' + QUOTENAME(t.[schema]) + '.' + QUOTENAME('xl_log_' + t.name)
    , @sql8 = '
DELETE FROM logs.change_logs WHERE [object_id] = OBJECT_ID(N''' + REPLACE(QUOTENAME(t.[schema]) + '.' + QUOTENAME(t.name), '''', '''''') + ''')'
    , @sql9 = '
DELETE FROM logs.handlers WHERE TABLE_SCHEMA = N''' + REPLACE(t.[schema], '''','''''') + ''' AND TABLE_NAME = N''' + REPLACE(t.name, '''','''''') + ''''
    , @sql10 = 'DELETE FROM logs.tables WHERE '
            +   'TABLE_SCHEMA = N''' + REPLACE(t.[schema], '''','''''') + ''''
            + ' AND TABLE_NAME = N''' + REPLACE(t.name, '''','''''') + ''''
FROM
    (
        SELECT
            [object_id]
            , OBJECT_SCHEMA_NAME(t.[object_id]) AS [schema]
            , t.name AS name
        FROM
            sys.tables t
        WHERE
            t.[schema_id] = SCHEMA_ID(@schema)
            AND t.name = @name
    ) t

DECLARE @br nvarchar(100) = ';' + CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10) + 'GO' + CHAR(13) + CHAR(10)

-- PRINT @sql1 + @br + @sql2 + @br + @sql3 + @br + @sql7 + @br + @sql8 + @br + @sql9 + @br

IF @execute_script = 1
    BEGIN
    EXEC (@sql1)
    EXEC (@sql2)
    EXEC (@sql3)
    EXEC (@sql7)
    EXEC (@sql8)
    EXEC (@sql9)
    EXEC (@sql10)
    END
ELSE
    BEGIN
    RAISERROR(@sql1, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql2, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql3, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql7, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql8, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql9, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    RAISERROR(@sql10, 0, 1) WITH NOWAIT
    RAISERROR('GO',  0, 1) WITH NOWAIT
    --PRINT @sql1
    --PRINT 'GO'
    --PRINT @sql2
    --PRINT 'GO'
    --PRINT @sql3
    --PRINT 'GO'
    --PRINT @sql7
    --PRINT 'GO'
    --PRINT @sql8
    --PRINT 'GO'
    --PRINT @sql9
    --PRINT 'GO'
    --PRINT @sql10
    --PRINT 'GO'
    END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure restores a record from the log
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_restore_record]
    @change_id int = NULL
    , @restore_previous bit = 0
    , @confirm bit = 0
    , @data_language varchar(10) = NULL
WITH EXECUTE AS OWNER
AS
BEGIN

BEGIN -- Change data --

SET NOCOUNT ON

DECLARE @message nvarchar(max)

IF @change_id IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'Specify @change_id', @data_language)
    RAISERROR(@message, 11, 0)
    RETURN
    END

IF @restore_previous IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'Specify @restore_previous', @data_language)
    RAISERROR(@message, 11, 0)
    RETURN
    END

DECLARE @source_schema nvarchar(128)
DECLARE @source_name nvarchar(128)
DECLARE @obj int
DECLARE @s1 xml
DECLARE @h1 int

SELECT
    @source_schema = t.TABLE_SCHEMA
    , @source_name = t.TABLE_NAME
    , @obj = c.[object_id]
    , @s1 = CASE WHEN @restore_previous = 1 THEN COALESCE(c.deleted, c.inserted) ELSE COALESCE(c.inserted, c.deleted) END
FROM
    logs.change_logs c
    LEFT OUTER JOIN logs.tables t ON t.[object_id] = c.[object_id]
WHERE
    c.change_id = @change_id

-- PRINT CAST(@s1 AS nvarchar(max))

IF @obj IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'change_id %i not found', @data_language)
    RAISERROR(@message, 11, 0, @change_id)
    RETURN
    END

END

BEGIN -- Object data --

EXECUTE AS CALLER

DECLARE @obj_schema nvarchar(128)
DECLARE @obj_name nvarchar(128)

SELECT @obj_schema = SCHEMA_NAME([schema_id]), @obj_name = name FROM sys.objects WHERE [object_id] = @obj

IF @obj_name IS NULL
    BEGIN
    SELECT @obj_schema = SCHEMA_NAME([schema_id]), @obj_name = name, @obj = [object_id] FROM sys.objects
        WHERE name = @source_name AND [schema_id] = SCHEMA_ID(@source_schema)
    END

IF @obj_name IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'Table ''%s.%s'' not found', @data_language)
    RAISERROR(@message, 11, 0, @source_schema, @source_name)
    RETURN
    END

IF HAS_PERMS_BY_NAME(QUOTENAME(@obj_schema) + '.' + QUOTENAME(@obj_name), 'OBJECT', 'UPDATE') = 0
    BEGIN
    SET @message = logs.get_translated_string(N'You have no the UPDATE permission on ''%s.%s''', @data_language)
    RAISERROR(@message, 11, 0, @obj_schema, @obj_name)
    RETURN
    END

END

BEGIN -- SQL --

EXEC sp_xml_preparedocument @h1 OUTPUT, @s1

DECLARE @set nvarchar(max) = ''
DECLARE @where nvarchar(max) = ''
DECLARE @obj_names nvarchar(max) = ''
DECLARE @values nvarchar(max) = ''
DECLARE @is_identity int

SELECT
    @set = @set + CASE WHEN is_pk = 1 THEN '' ELSE ', ' + t.name + ' = ' + t.value END
    , @where = @where + CASE WHEN is_pk = 0 THEN '' ELSE ' AND ' + t.name + ' = ' + t.value END
    , @obj_names = @obj_names + ', ' + t.name
    , @values = @values + ', ' + t.value
    , @is_identity = @is_identity + is_identity
FROM
    (
    SELECT
        QUOTENAME(c.name) AS name
        , CASE WHEN t1.[text] IS NULL THEN 'NULL'
            WHEN tp.name IN ('nvarchar', 'nchar', 'ntext', 'varchar', 'char', 'varbinary', 'binary', 'text', 'uniqueidentifier', 'xml', 'timestamp')
                THEN 'N''' + REPLACE(CAST(t1.[text] AS nvarchar), '''', '''''') + ''''
            WHEN tp.name IN ('datetime', 'datetime2', 'smalldatetime', 'date', 'time', 'datetimeoffset') THEN '''' + CAST(t1.[text] AS nvarchar) + ''''
            ELSE CAST(t1.[text] AS nvarchar)
            END AS value
        , CASE WHEN ic.column_id IS NOT NULL THEN 1 ELSE 0 END AS is_pk
        , c.is_identity
    FROM
        sys.columns c
        LEFT OUTER JOIN (
            SELECT
                logs.get_unescaped_parameter_name(t2.localname) AS name
                , t1.[text]
            FROM
                OPENXML (@h1, '/row', 0) t1
                INNER JOIN OPENXML (@h1, '/row', 0) t2 ON t2.id = t1.parentid
            WHERE
                t2.nodetype = 2
        ) t1 ON t1.name = c.name
        LEFT OUTER JOIN sys.indexes i ON i.[object_id] = c.[object_id] AND i.is_primary_key = 1
        LEFT OUTER JOIN sys.index_columns ic ON ic.[object_id] = c.[object_id] AND ic.column_id = c.column_id AND ic.index_id = i.index_id
        INNER JOIN sys.types tp ON tp.user_type_id = c.user_type_id
    WHERE
        c.[object_id] = @obj
    ) t

EXEC sp_xml_removedocument @h1

SET @set = SUBSTRING(@set, 3, LEN(@set))
SET @where = SUBSTRING(@where, 6, LEN(@where))
SET @obj_names = SUBSTRING(@obj_names, 3, LEN(@obj_names))
SET @values = SUBSTRING(@values, 3, LEN(@values))

DECLARE @sql nvarchar(max)

-- PRINT COALESCE(@set, '@set = NULL')
-- PRINT COALESCE(@where,  '@where = NULL')
-- PRINT COALESCE(@obj_names,  '@obj_names = NULL')
-- PRINT COALESCE(@values, '@values = NULL')

SET @sql =
      'UPDATE ' + QUOTENAME(@obj_schema) + '.' + QUOTENAME(@obj_name) + CHAR(13) + CHAR(10)
    + 'SET' + CHAR(13) + CHAR(10)
    + '    ' + @set + CHAR(13) + CHAR(10)
    + 'WHERE ' + CHAR(13) + CHAR(10)
    + '    ' + @where + CHAR(13) + CHAR(10)
    + CHAR(13) + CHAR(10)
    + 'IF @@ROWCOUNT = 0' + CHAR(13) + CHAR(10)
    + CASE WHEN @is_identity = 0 THEN '' ELSE
      '    BEGIN' + CHAR(13) + CHAR(10)
    + '    SET IDENTITY_INSERT ' + QUOTENAME(@obj_schema) + '.' + QUOTENAME(@obj_name) + ' ON' + CHAR(13) + CHAR(10)
        END
    + '    INSERT INTO ' + QUOTENAME(@obj_schema) + '.' + QUOTENAME(@obj_name) + CHAR(13) + CHAR(10)
    + '       (' + @obj_names + ')' + CHAR(13) + CHAR(10)
    + '    VALUES ' + CHAR(13) + CHAR(10)
    + '       (' + @values + ')' + CHAR(13) + CHAR(10)
    + CASE WHEN @is_identity = 0 THEN '' ELSE
      '    SET IDENTITY_INSERT ' + QUOTENAME(@obj_schema) + '.' + QUOTENAME(@obj_name) + ' OFF' + CHAR(13) + CHAR(10)
    + '    END' + CHAR(13) + CHAR(10)
        END

END

BEGIN -- EXEC --

-- PRINT @sql

IF @confirm = 1
    BEGIN
    EXEC (@sql)

    REVERT
    END
ELSE
    BEGIN
    SET @message = logs.get_translated_string(N'Set @confirm = 1 to restore the record.', @data_language) + CHAR(13) + CHAR(10)
    + CHAR(13) + CHAR(10)
    + logs.get_translated_string(N'SQL code to restore the record:', @data_language) + CHAR(13) + CHAR(10)
    + CHAR(13) + CHAR(10)

    PRINT @message + @sql
    END

END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure restores a current record from the log
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_restore_current_record]
    @change_id int = NULL
    , @confirm bit = 0
AS
BEGIN

EXEC logs.xl_actions_restore_record @change_id, 0, @confirm

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure restores a previous record from the log
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_restore_previous_record]
    @change_id int = NULL
    , @confirm bit = 0
AS
BEGIN

EXEC logs.xl_actions_restore_record @change_id, 1, @confirm

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure shows a foreign key record
--
-- @column_name - the active cell column name in a table
-- @name        - a value of the name column (use it for xl_actions_select_record)
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_select_lookup_id]
    @change_id int = NULL
    , @column_name nvarchar(128) = NULL
    , @name nvarchar(128) = NULL
    , @data_language varchar(10) = NULL
WITH EXECUTE AS OWNER
AS
BEGIN

BEGIN -- Change data --

SET NOCOUNT ON

DECLARE @message nvarchar(max)

IF COALESCE(@column_name, 'value') = 'value' AND @name IS NOT NULL
    SET @column_name = @name

DECLARE @source_schema nvarchar(128)
DECLARE @source_name nvarchar(128)
DECLARE @obj int
DECLARE @s1 xml, @s2 xml
DECLARE @h1 int, @h2 int

SELECT
    @source_schema = t.TABLE_SCHEMA
    , @source_name = t.TABLE_NAME
    , @obj = c.[object_id]
    , @s1 = c.inserted
    , @s2 = c.deleted
FROM
    logs.change_logs c
    LEFT OUTER JOIN logs.tables t ON t.[object_id] = c.[object_id]
WHERE
    c.change_id = @change_id

IF @obj IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'change_id %i not found', @data_language)
    RAISERROR(@message, 11, 0, @change_id)
    RETURN
    END

END

BEGIN -- Object data --

EXECUTE AS CALLER

DECLARE @obj_schema nvarchar(128)
DECLARE @obj_name nvarchar(128)

SELECT @obj_schema = SCHEMA_NAME([schema_id]), @obj_name = name FROM sys.objects WHERE [object_id] = @obj

IF @obj_name IS NULL AND @source_name IS NOT NULL
    BEGIN
    SELECT @obj_schema = SCHEMA_NAME([schema_id]), @obj_name = name, @obj = [object_id] FROM sys.objects
        WHERE name = @source_name AND [schema_id] = SCHEMA_ID(@source_schema)
    END

IF @obj_name IS NULL
    BEGIN
    SET @message = logs.get_translated_string(N'Table ''%s.%s'' not found', @data_language)
    RAISERROR(@message, 11, 0, @source_schema, @source_name)
    RETURN
    END

IF HAS_PERMS_BY_NAME(QUOTENAME(@obj_schema) + '.' + QUOTENAME(@obj_name), 'OBJECT', 'SELECT') = 0
    BEGIN
    SET @message = logs.get_translated_string(N'You have no the SELECT permission on ''%s.%s''', @data_language)
    RAISERROR(@message, 11, 0, @obj_schema, @obj_name)
    RETURN
    END

END

BEGIN -- Column data --

DECLARE @table_schema nvarchar(128)
DECLARE @table_name nvarchar(128)
DECLARE @key_column_name nvarchar(128)
DECLARE @key_column_type tinyint

SELECT
    @table_schema = ps.name
    , @table_name = po.name
    , @key_column_name = pc.name
    , @key_column_type = CASE
        WHEN tp.name IN ('nvarchar', 'nchar', 'ntext', 'varchar', 'char', 'varbinary', 'binary', 'text', 'uniqueidentifier', 'xml', 'timestamp') THEN 1
        WHEN tp.name IN ('datetime', 'datetime2', 'smalldatetime', 'date', 'time', 'datetimeoffset') THEN 1
        ELSE 0
        END
FROM
    sys.foreign_key_columns kc
    INNER JOIN sys.objects fo ON fo.[object_id] = kc.parent_object_id
    INNER JOIN sys.objects po ON po.[object_id] = kc.referenced_object_id
    INNER JOIN sys.schemas fs ON fs.[schema_id] = fo.[schema_id]
    INNER JOIN sys.schemas ps ON ps.[schema_id] = po.[schema_id]
    INNER JOIN sys.columns fc ON fc.[object_id] = kc.parent_object_id AND fc.column_id = kc.parent_column_id
    INNER JOIN sys.columns pc ON pc.[object_id] = kc.referenced_object_id AND pc.column_id = kc.referenced_column_id
    INNER JOIN sys.foreign_keys fk ON fk.[object_id] = kc.constraint_object_id
    INNER JOIN sys.index_columns ic ON ic.[object_id] = kc.referenced_object_id AND ic.column_id = kc.referenced_column_id AND ic.index_id = fk.key_index_id
    INNER JOIN sys.types tp ON tp.user_type_id = pc.user_type_id
WHERE
    fo.[object_id] = @obj
    AND fc.name = @column_name

IF @table_name IS NULL
    BEGIN
    RAISERROR('Column ''%s'' has no referenced tables', 11, 0, @column_name)
    --SELECT NULL AS [none]
    RETURN
    END

DECLARE @key1 nvarchar(128)
DECLARE @key2 nvarchar(128)

IF @s1 IS NOT NULL
    BEGIN
    EXEC sp_xml_preparedocument @h1 OUTPUT, @s1

    SELECT
        @key1 = t1.[text]
    FROM
        OPENXML (@h1, '/row', 0) t1
        INNER JOIN OPENXML (@h1, '/row', 0) t2 ON t2.id = t1.parentid
    WHERE
        t2.nodetype = 2
        AND t2.localname = @column_name

    EXEC sp_xml_removedocument @h1
    END

IF @s2 IS NOT NULL
    BEGIN
    EXEC sp_xml_preparedocument @h2 OUTPUT, @s2

    SELECT
        @key2 = t1.[text]
    FROM
        OPENXML (@h2, '/row', 0) t1
        INNER JOIN OPENXML (@h2, '/row', 0) t2 ON t2.id = t1.parentid
    WHERE
        t2.nodetype = 2
        AND t2.localname = @column_name

    EXEC sp_xml_removedocument @h2
    END

END

BEGIN -- SQL --

DECLARE @sql nvarchar(max)

SET @sql = 'SELECT * FROM ' + QUOTENAME(@table_schema) + '.' + QUOTENAME(@table_name) + ' WHERE ' + QUOTENAME(@key_column_name)
    + CASE WHEN @key2 IS NULL THEN ' = @id1' WHEN @key1 IS NULL THEN ' = @id2' ELSE ' = @id1 OR ' + QUOTENAME(@key_column_name) + ' = @id2' END

-- PRINT @sql

END

BEGIN -- EXEC --

IF @key_column_type = 1
    EXEC sys.sp_executesql @stmt = @sql, @params = N'@id1 nvarchar(128), @id2 nvarchar(128)', @id1 = @key1, @id2 = @key2
ELSE
    BEGIN
    DECLARE @id1 nvarchar(128), @id2 nvarchar(128)
    SET @id1 = CAST(@key1 AS int)
    SET @id2 = CAST(@key2 AS int)
    EXEC sys.sp_executesql @stmt = @sql, @params = N'@id1 int, @id2 int', @id1 = @id1, @id2 = @id2
    END

REVERT

END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure shows a log record
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_select_record]
    @change_id int = NULL
    , @data_language varchar(10) = NULL
WITH EXECUTE AS OWNER
AS
BEGIN

BEGIN -- Change data --

SET NOCOUNT ON

DECLARE @message nvarchar(max)

IF @data_language IS NULL SET @data_language = 'en'

DECLARE @source_schema nvarchar(128)
DECLARE @source_name nvarchar(128)
DECLARE @obj int
DECLARE @s1 xml, @s2 xml
DECLARE @h1 int, @h2 int
DECLARE @d datetime
DECLARE @u nvarchar(128)

SELECT
    @source_schema = t.TABLE_SCHEMA
    , @source_name = t.TABLE_NAME
    , @obj = c.[object_id]
    , @s1 = c.inserted
    , @s2 = c.deleted
    , @d = c.change_date
    , @u = c.change_user
FROM
    logs.change_logs c
    LEFT OUTER JOIN logs.tables t ON t.[object_id] = c.[object_id]
WHERE
    c.change_id = @change_id

IF @obj IS NULL
    BEGIN
    -- SET @message = logs.get_translated_string(N'change_id %i not found', @data_language)
    -- RAISERROR(@message, 11, 0, @change_id)
    RETURN
    END

-- PRINT CAST(@s1 AS nvarchar(max))
-- PRINT CAST(@s2 AS nvarchar(max))

END

BEGIN -- Object data --

EXECUTE AS CALLER

DECLARE @obj_schema nvarchar(128)
DECLARE @obj_name nvarchar(128) = NULL

SELECT @obj_schema = SCHEMA_NAME([schema_id]), @obj_name = name FROM sys.objects WHERE [object_id] = @obj

IF @obj_name IS NULL
    BEGIN
    SELECT @obj_schema = SCHEMA_NAME([schema_id]), @obj_name = name, @obj = [object_id]  FROM sys.objects
        WHERE name = @source_name AND [schema_id] = SCHEMA_ID(@source_schema)
    END

IF @obj_name IS NULL
    BEGIN
    -- SET @message = logs.get_translated_string(N'Table ''%s.%s'' not found', @data_language)
    -- RAISERROR(@message, 11, 0, @source_schema, @source_name)
    RETURN
    END

IF HAS_PERMS_BY_NAME(QUOTENAME(@obj_schema) + '.' + QUOTENAME(@obj_name), 'OBJECT', 'SELECT') = 0
    BEGIN
    -- SET @message = logs.get_translated_string(N'You have no the SELECT permission on ''%s.%s''', @data_language)
    -- RAISERROR(@message, 11, 0, @obj_schema, @obj_name)
    RETURN
    END

END

BEGIN -- SELECT --

IF @s1 IS NULL AND @s2 IS NULL
    SELECT NULL AS name, NULL AS value
ELSE IF @s1 IS NULL
    BEGIN
    EXEC sp_xml_preparedocument @h2 OUTPUT, @s2

    SELECT t.name, t.value FROM (
    SELECT
        c.column_id
        , COALESCE(c.name, t2.localname) AS name
        , t1.[text] AS value
    FROM
        OPENXML (@h2, '/row', 0) t1
        INNER JOIN OPENXML (@h2, '/row', 0) t2 ON t2.id = t1.parentid
        INNER JOIN sys.columns c ON c.[object_id] = @obj AND c.name = logs.get_unescaped_parameter_name(t2.localname)
    WHERE
        t2.nodetype = 2
    UNION ALL
    SELECT
        r.column_id
        , COALESCE(t.TRANSLATED_NAME, r.name) AS name
        , COALESCE(v.TRANSLATED_NAME, r.value) AS value
    FROM
        (VALUES (1000, 'change_type', 'deleted')) r(column_id, name, value)
        LEFT OUTER JOIN logs.translations t ON
            t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME = 'xl_actions_select_records' AND t.LANGUAGE_NAME = @data_language AND t.COLUMN_NAME = r.name
        LEFT OUTER JOIN logs.translations v ON
            v.TABLE_SCHEMA = 'logs' AND v.TABLE_NAME = 'change_type' AND v.LANGUAGE_NAME = @data_language AND v.COLUMN_NAME = r.value
    UNION ALL
    SELECT
        r.column_id
        , COALESCE(t.TRANSLATED_NAME, r.name) AS name
        , r.value
    FROM
        (VALUES (1001, 'change_date', CAST(@d AS nvarchar)), (1002, 'change_user', @u)) r(column_id, name, value)
        LEFT OUTER JOIN logs.translations t ON
            t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME = 'xl_actions_select_records' AND t.LANGUAGE_NAME = @data_language AND t.COLUMN_NAME = r.name
    ) t
    ORDER BY
        t.column_id

    EXEC sp_xml_removedocument @h2
    END
ELSE IF @s2 IS NULL
    BEGIN
    EXEC sp_xml_preparedocument @h1 OUTPUT, @s1

    SELECT t.name, t.value FROM (
    SELECT
        c.column_id
        , COALESCE(c.name, t2.localname) AS name
        , t1.[text] AS value
    FROM
        OPENXML (@h1, '/row', 0) t1
        INNER JOIN OPENXML (@h1, '/row', 0) t2 ON t2.id = t1.parentid
        INNER JOIN sys.columns c ON c.[object_id] = @obj AND c.name = logs.get_unescaped_parameter_name(t2.localname)
    WHERE
        t2.nodetype = 2
    UNION ALL
    SELECT
        r.column_id
        , COALESCE(t.TRANSLATED_NAME, r.name) AS name
        , COALESCE(v.TRANSLATED_NAME, r.value) AS value
    FROM
        (VALUES (1000, 'change_type', 'inserted')) r(column_id, name, value)
        LEFT OUTER JOIN logs.translations t ON
            t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME = 'xl_actions_select_records' AND t.LANGUAGE_NAME = @data_language AND t.COLUMN_NAME = r.name
        LEFT OUTER JOIN logs.translations v ON
            v.TABLE_SCHEMA = 'logs' AND v.TABLE_NAME = 'change_type' AND v.LANGUAGE_NAME = @data_language AND v.COLUMN_NAME = r.value
    UNION ALL
    SELECT
        r.column_id
        , COALESCE(t.TRANSLATED_NAME, r.name) AS name
        , r.value
    FROM
        (VALUES (1001, 'change_date', CAST(@d AS nvarchar)), (1002, 'change_user', @u)) r(column_id, name, value)
        LEFT OUTER JOIN logs.translations t ON
            t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME = 'xl_actions_select_records' AND t.LANGUAGE_NAME = @data_language AND t.COLUMN_NAME = r.name
    ) t
    ORDER BY
        t.column_id

    EXEC sp_xml_removedocument @h1
    END
ELSE
    BEGIN
    EXEC sp_xml_preparedocument @h1 OUTPUT, @s1
    EXEC sp_xml_preparedocument @h2 OUTPUT, @s2

    -- PRINT @h1
    -- PRINT @h2

    SELECT t.name, t.value FROM (
    SELECT
        c.column_id
        , COALESCE(c.name, t2.localname) AS name
        , CASE
            WHEN t1.[text] IS NULL AND t3.[text] IS NULL THEN NULL
            WHEN t1.[text] IS NULL THEN '[' + CAST(t3.[text] AS nvarchar(max)) + ']'
            WHEN t3.[text] IS NULL THEN CAST(t1.[text] AS nvarchar(max)) + ' []'
            WHEN CAST(t1.[text] AS nvarchar(max)) = CAST(t3.[text] AS nvarchar(max)) THEN CAST(t1.[text] AS nvarchar(max))
            ELSE CAST(t1.[text] AS nvarchar(max)) + ' [' + CAST(t3.[text] AS nvarchar(max)) + ']' END AS value
    FROM
        OPENXML (@h1, '/row', 0) t1
        INNER JOIN OPENXML (@h1, '/row', 0) t2 ON t2.id = t1.parentid
        LEFT OUTER JOIN OPENXML (@h2, '/row', 0) t4 ON t4.localname = t2.localname
        LEFT OUTER JOIN OPENXML (@h2, '/row', 0) t3 ON t3.parentid = t4.id
        INNER JOIN sys.columns c ON c.[object_id] = @obj AND c.name = logs.get_unescaped_parameter_name(t2.localname)
    WHERE
        t2.nodetype = 2
    UNION ALL
    SELECT
        r.column_id
        , COALESCE(t.TRANSLATED_NAME, r.name) AS name
        , COALESCE(v.TRANSLATED_NAME, r.value) AS value
    FROM
        (VALUES (1000, 'change_type', 'updated')) r(column_id, name, value)
        LEFT OUTER JOIN logs.translations t ON
            t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME = 'xl_actions_select_records' AND t.LANGUAGE_NAME = @data_language AND t.COLUMN_NAME = r.name
        LEFT OUTER JOIN logs.translations v ON
            v.TABLE_SCHEMA = 'logs' AND v.TABLE_NAME = 'change_type' AND v.LANGUAGE_NAME = @data_language AND v.COLUMN_NAME = r.value
    UNION ALL
    SELECT
        r.column_id
        , COALESCE(t.TRANSLATED_NAME, r.name) AS name
        , r.value
    FROM
        (VALUES (1001, 'change_date', CAST(@d AS nvarchar)), (1002, 'change_user', @u)) r(column_id, name, value)
        LEFT OUTER JOIN logs.translations t ON
            t.TABLE_SCHEMA = 'logs' AND t.TABLE_NAME = 'xl_actions_select_records' AND t.LANGUAGE_NAME = @data_language AND t.COLUMN_NAME = r.name
    ) t
    ORDER BY
        t.column_id

    EXEC sp_xml_removedocument @h1
    EXEC sp_xml_removedocument @h2
    END

REVERT

END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure shows log records
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_select_records]
    @schema nvarchar(128) = NULL
    , @name nvarchar(128) = NULL
    , @id int = NULL
    , @keys nvarchar(445) = NULL
    , @change_type tinyint = NULL
    , @data_language varchar(10) = NULL
WITH EXECUTE AS OWNER
AS
BEGIN

BEGIN -- Object data --

SET NOCOUNT ON

DECLARE @message nvarchar(max)

DECLARE @obj int

EXECUTE AS CALLER

SELECT @obj = [object_id] FROM sys.objects WHERE name = @name AND schema_id = SCHEMA_ID(@schema)

IF @obj IS NULL
    BEGIN
    SELECT @schema + '.' + @name + ' ' + logs.get_translated_string(N'not found', @data_language) AS [Message]
    RETURN
    END

IF HAS_PERMS_BY_NAME(QUOTENAME(@schema) + '.' + QUOTENAME(@name), 'OBJECT', 'SELECT') = 0
    BEGIN
    SELECT @schema + '.' + @name + ' ' + logs.get_translated_string(N'has no permission', @data_language) AS [Message]
    RETURN
    END

END

BEGIN -- Columns --

DECLARE @sql nvarchar(max)

SET @sql = 'SELECT @columns = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE((SELECT ' + CHAR(13) + CHAR(10)
    + '      ' + STUFF((
        SELECT
            '    , ''' + LOWER(tp.name) +
                CASE WHEN tp.name IN ('varchar', 'char', 'varbinary', 'binary', 'text')
                        THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length AS VARCHAR(5)) END + ')'
                        WHEN tp.name IN ('nvarchar', 'nchar', 'ntext')
                        THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length / 2 AS VARCHAR(5)) END + ')'
                        WHEN tp.name IN ('datetime2', 'time2', 'datetimeoffset')
                        THEN '(' + CAST(c.scale AS VARCHAR(5)) + ')'
                        WHEN tp.name = 'decimal'
                        THEN '(' + CAST(c.[precision] AS VARCHAR(5)) + ',' + CAST(c.scale AS VARCHAR(5)) + ')'
                    ELSE ''
                END
            + ''''') AS [' + REPLACE(REPLACE(c.name, '''', ''''''), ']', ']]') + ']'''
            + ' AS [' + REPLACE(c.name, ']', ']]') + ']' + CHAR(13) + CHAR(10)
        FROM
            sys.columns c
            INNER JOIN sys.types tp ON c.user_type_id = tp.user_type_id
        WHERE
            c.[object_id] = (SELECT [object_id] FROM sys.objects WHERE name = @name
                AND schema_id = SCHEMA_ID(@schema))
        ORDER BY
            c.column_id
        FOR XML PATH(''), TYPE).value('.', 'nvarchar(MAX)'), 1, 6, '')
        + '    FOR XML RAW)' + CHAR(13) + CHAR(10)
        + '    , ''&quot;'', ''"'')'
        + '    , ''&lt;'', ''<'')'
        + '    , ''&gt;'', ''>'')'
        + '    , ''&amp;'', ''&'')'
        + '    , ''="'', '''''', '''''')' + CHAR(13) + CHAR(10)
        + '    , ''"/>'', '''')' + CHAR(13) + CHAR(10)
        + '    , ''" '', CHAR(13) + CHAR(10) + CHAR(9) + '', t.r.value(''''row[1]/@'')' + CHAR(13) + CHAR(10)
        + '    , ''<row '', CHAR(9) + ''t.r.value(''''row[1]/@'')' + CHAR(13) + CHAR(10)

-- PRINT @sql

DECLARE @columns nvarchar(max)

EXEC sp_executesql @sql, N'@columns nvarchar(max) OUTPUT', @columns = @columns OUTPUT

END

BEGIN -- SQL --

SET @sql = 'SELECT' + CHAR(13) + CHAR(10) + @columns + '
    , t.change_type
    , t.change_date
    , t.change_user
    , t.change_id
FROM
    (
    SELECT
        COALESCE(t.inserted, t.deleted) AS r
        , ' + (
        SELECT
            'CASE WHEN t.deleted IS NULL THEN N''' + COALESCE(t1.TRANSLATED_NAME, r.inserted)
            + ''' WHEN t.inserted IS NULL THEN N''' + COALESCE(t3.TRANSLATED_NAME, r.deleted)
            + ''' ELSE N''' + COALESCE(t2.TRANSLATED_NAME, r.updated) + ''' END AS change_type'
        FROM
            (VALUES ('inserted', 'updated', 'deleted')) r(inserted, updated, deleted)
            LEFT OUTER JOIN logs.translations t1 ON
                t1.TABLE_SCHEMA = 'logs' AND t1.TABLE_NAME = 'change_type' AND t1.LANGUAGE_NAME = @data_language AND t1.COLUMN_NAME = r.inserted
            LEFT OUTER JOIN logs.translations t2 ON
                t2.TABLE_SCHEMA = 'logs' AND t2.TABLE_NAME = 'change_type' AND t2.LANGUAGE_NAME = @data_language AND t2.COLUMN_NAME = r.updated
            LEFT OUTER JOIN logs.translations t3 ON
                t3.TABLE_SCHEMA = 'logs' AND t3.TABLE_NAME = 'change_type' AND t3.LANGUAGE_NAME = @data_language AND t3.COLUMN_NAME = r.deleted
        ) + '
        , t.change_date
        , t.change_user
        , t.change_id
    FROM
        logs.change_logs t
    WHERE
        t.[object_id] = ' + CAST(@obj AS nvarchar) + '
        ' + CASE WHEN @keys IS NOT NULL AND @keys <> '<row/>' THEN 'AND t.keys = @keys' WHEN @id IS NOT NULL THEN 'AND t.id = ' + CAST(@id AS nvarchar) ELSE '' END +  '
        ' + CASE WHEN @change_type IS NULL THEN '' ELSE 'AND t.change_type = ' + CAST(@change_type AS nvarchar) + '' END + '
    ) t
'

-- PRINT COALESCE(@sql, 'NULL')

END

BEGIN -- EXEC --

REVERT

IF @sql IS NULL
    BEGIN
    SELECT logs.get_translated_string(N'No records', @data_language) AS [Message]
    END
ELSE IF @keys IS NOT NULL AND @keys <> '<row/>'
    EXEC sys.sp_executesql @sql, N'@keys nvarchar(445)', @keys
ELSE
    EXEC (@sql)

END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: The procedure sets permissions of the application roles
-- =============================================

CREATE PROCEDURE [logs].[xl_actions_set_role_permissions]
WITH EXECUTE AS OWNER
AS
BEGIN

SET NOCOUNT ON

GRANT SELECT, EXECUTE, VIEW DEFINITION ON SCHEMA::logs TO log_admins;

GRANT SELECT, VIEW DEFINITION ON logs.handlers              TO log_users;
GRANT SELECT, VIEW DEFINITION ON logs.translations          TO log_users;

GRANT SELECT ON logs.view_translations                      TO log_users;
GRANT SELECT ON logs.view_user_handlers                     TO log_users;

GRANT EXECUTE ON logs.xl_actions_restore_current_record     TO log_users;
GRANT EXECUTE ON logs.xl_actions_restore_previous_record    TO log_users;
GRANT EXECUTE ON logs.xl_actions_restore_record             TO log_users;
GRANT EXECUTE ON logs.xl_actions_select_lookup_id           TO log_users;
GRANT EXECUTE ON logs.xl_actions_select_record              TO log_users;
GRANT EXECUTE ON logs.xl_actions_select_records             TO log_users;

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Exports Change Tracking Framework settings
-- =============================================

CREATE PROCEDURE [logs].[xl_export_settings]
    @part tinyint = NULL
    , @sort_by_names bit = 0
    , @schema nvarchar(128) = NULL
    , @language varchar(10) = NULL
AS
BEGIN

SET NOCOUNT ON;

DECLARE @app_schema_only bit = 0

IF @schema = 'x'
    BEGIN
    SET @schema = NULL
    SET @app_schema_only = 1
    END

SELECT
    t.command
FROM
    (
SELECT
    2 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY TABLE_SCHEMA, TABLE_NAME, EVENT_NAME, CASE WHEN COLUMN_NAME IS NULL THEN 0 ELSE 1 END, COLUMN_NAME, HANDLER_SCHEMA, HANDLER_NAME, MENU_ORDER) ELSE ID END AS ID
    , 'INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES ('
           + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN COLUMN_NAME               IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(COLUMN_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN EVENT_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(EVENT_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN HANDLER_SCHEMA            IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(HANDLER_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN HANDLER_NAME              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(HANDLER_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN HANDLER_TYPE              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(HANDLER_TYPE, '''', '''''') + '''' END
    + ', ' + CASE WHEN HANDLER_CODE              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(HANDLER_CODE, '''', '''''') + '''' END
    + ', ' + CASE WHEN TARGET_WORKSHEET          IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TARGET_WORKSHEET, '''', '''''') + '''' END
    + ', ' + CASE WHEN MENU_ORDER                IS NULL THEN 'NULL' ELSE CAST(MENU_ORDER AS nvarchar(128))  END
    + ', ' + CASE WHEN EDIT_PARAMETERS           IS NULL THEN 'NULL' ELSE CAST(EDIT_PARAMETERS AS nvarchar(128))  END
    + ');' AS command
FROM
    logs.handlers
WHERE
    TABLE_SCHEMA = COALESCE(@schema, TABLE_SCHEMA)
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('logs'))

UNION ALL SELECT 3 AS part, -1 AS ID, '' AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            logs.handlers
        WHERE
            TABLE_SCHEMA = COALESCE(@schema, TABLE_SCHEMA)
            AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('logs'))
    )

UNION ALL
SELECT
    3 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY LANGUAGE_NAME, CASE WHEN COLUMN_NAME IS NULL THEN 0 ELSE 1 END, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME) ELSE ID END AS ID
    , 'INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES ('
           + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN COLUMN_NAME               IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(COLUMN_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN LANGUAGE_NAME             IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(LANGUAGE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TRANSLATED_NAME           IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TRANSLATED_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TRANSLATED_DESC           IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TRANSLATED_DESC, '''', '''''') + '''' END
    + ', ' + CASE WHEN TRANSLATED_COMMENT        IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TRANSLATED_COMMENT, '''', '''''') + '''' END
    + ');' AS command
FROM
    logs.translations
WHERE
    TABLE_SCHEMA = COALESCE(@schema, TABLE_SCHEMA)
    AND COALESCE(LANGUAGE_NAME, '') = COALESCE(@language, LANGUAGE_NAME, '')
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('logs'))

UNION ALL SELECT 4 AS part, -1 AS ID, '' AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            logs.translations
        WHERE
            TABLE_SCHEMA = COALESCE(@schema, TABLE_SCHEMA)
            AND COALESCE(LANGUAGE_NAME, '') = COALESCE(@language, LANGUAGE_NAME, '')
            AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('logs'))
    )

UNION ALL
SELECT
    4 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY TABLE_SCHEMA, TABLE_NAME) ELSE ID END AS ID
    , 'INSERT INTO logs.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES ('
           + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_EXCEL_FORMAT_XML    IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(CAST(TABLE_EXCEL_FORMAT_XML AS nvarchar(max)), '''', '''''') + '''' END
    + ');' AS command
FROM
    logs.formats
WHERE
    TABLE_SCHEMA = COALESCE(@schema, TABLE_SCHEMA)
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('logs'))

UNION ALL SELECT 5 AS part, -1 AS ID, '' AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            logs.formats
        WHERE
            TABLE_SCHEMA = COALESCE(@schema, TABLE_SCHEMA)
            AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('logs'))
    )

UNION ALL
SELECT
    5 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY NAME) ELSE ID END AS ID
    , 'INSERT INTO logs.workbooks (NAME, TEMPLATE, DEFINITION, TABLE_SCHEMA) VALUES ('
           + CASE WHEN NAME                      IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TEMPLATE                  IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TEMPLATE, '''', '''''') + '''' END
    + ', ' + CASE WHEN DEFINITION                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(DEFINITION, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + ');' AS command
FROM
    logs.workbooks
WHERE
    TABLE_SCHEMA = COALESCE(@schema, TABLE_SCHEMA)
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('logs'))

UNION ALL SELECT 6 AS part, -1 AS ID, '' AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            logs.workbooks
        WHERE
            (@app_schema_only = 0 OR @schema IN ('logs'))
    )

UNION ALL
SELECT
    6 AS part
    , ROW_NUMBER() OVER (ORDER BY s.name, t.name) AS ID
    , 'EXEC logs.xl_actions_create_triggers N''' + s.name + ''', N''' + t.name + ''' , 1;' AS command
FROM
    sys.tables t
    INNER JOIN sys.schemas s ON s.[schema_id] = t.[schema_id]
    LEFT OUTER JOIN sys.triggers r1 ON r1.parent_id = t.[object_id] AND r1.name = 'trigger_' + t.name + '_log_insert'
    LEFT OUTER JOIN sys.triggers r2 ON r2.parent_id = t.[object_id] AND r2.name = 'trigger_' + t.name + '_log_update'
    LEFT OUTER JOIN sys.triggers r3 ON r3.parent_id = t.[object_id] AND r3.name = 'trigger_' + t.name + '_log_delete'
    LEFT OUTER JOIN sys.procedures p ON p.name = 'xl_log_' + t.name AND p.[schema_id] = t.[schema_id]
WHERE
    NOT t.name LIKE 'sys%'
    AND r1.name IS NOT NULL
    AND r2.name IS NOT NULL
    AND r3.name IS NOT NULL
    AND s.name = COALESCE(@schema, s.name)
    AND (@app_schema_only = 0 OR NOT s.name IN ('logs'))

UNION ALL SELECT 7 AS part, -1 AS ID, '' AS command
    WHERE EXISTS(
        SELECT
            t.[object_id]
        FROM
            sys.tables t
            INNER JOIN sys.schemas s ON s.[schema_id] = t.[schema_id]
            LEFT OUTER JOIN sys.triggers r1 ON r1.parent_id = t.[object_id] AND r1.name = 'trigger_' + t.name + '_log_insert'
            LEFT OUTER JOIN sys.triggers r2 ON r2.parent_id = t.[object_id] AND r2.name = 'trigger_' + t.name + '_log_update'
            LEFT OUTER JOIN sys.triggers r3 ON r3.parent_id = t.[object_id] AND r3.name = 'trigger_' + t.name + '_log_delete'
            LEFT OUTER JOIN sys.procedures p ON p.name = 'xl_log_' + t.name AND p.[schema_id] = t.[schema_id]
        WHERE
            NOT t.name LIKE 'sys%'
            AND r1.name IS NOT NULL
            AND r2.name IS NOT NULL
            AND r3.name IS NOT NULL
            AND s.name = COALESCE(@schema, s.name)
            AND (@app_schema_only = 0 OR NOT s.name IN ('logs'))
    )

UNION ALL
SELECT
    7 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY OBJECT_SCHEMA, [OBJECT_NAME], BASE_TABLE_SCHEMA, BASE_TABLE_NAME) ELSE ID END AS ID
    , 'INSERT INTO logs.base_tables (OBJECT_SCHEMA, OBJECT_NAME, BASE_TABLE_SCHEMA, BASE_TABLE_NAME) VALUES ('
           + CASE WHEN OBJECT_SCHEMA            IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(OBJECT_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN [OBJECT_NAME]            IS NULL THEN 'NULL' ELSE 'N''' + REPLACE([OBJECT_NAME], '''', '''''') + '''' END
    + ', ' + CASE WHEN BASE_TABLE_SCHEMA        IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(BASE_TABLE_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN BASE_TABLE_NAME          IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(BASE_TABLE_NAME, '''', '''''') + '''' END
    + ');' AS command
FROM
    logs.base_tables
WHERE
    (@app_schema_only = 0 OR NOT @schema IN ('logs'))

UNION ALL SELECT 8 AS part, -1 AS ID, '' AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            logs.base_tables
        WHERE
            (@app_schema_only = 0 OR NOT @schema IN ('logs'))
    )

    ) t
WHERE
    t.part = COALESCE(@part, t.part)
    AND (@part IS NULL OR NOT t.command = '')
ORDER BY
    t.part
    , t.ID

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     4.0, 2022-07-05
-- Description: Imports Change Tracking Framework handlers
-- =============================================

CREATE PROCEDURE [logs].[xl_import_handlers]
    @TABLE_SCHEMA nvarchar(20) = NULL
    , @TABLE_NAME nvarchar(128) = NULL
    , @COLUMN_NAME nvarchar(128) = NULL
    , @EVENT_NAME varchar(25) = NULL
    , @HANDLER_SCHEMA nvarchar(20) = NULL
    , @HANDLER_NAME nvarchar(128) = NULL
    , @HANDLER_TYPE nvarchar(25) = NULL
    , @HANDLER_CODE nvarchar(max) = NULL
    , @TARGET_WORKSHEET nvarchar(128) = NULL
    , @MENU_ORDER int = NULL
    , @EDIT_PARAMETERS bit = NULL
AS
BEGIN

SET NOCOUNT ON;

UPDATE logs.handlers
SET
    HANDLER_CODE = @HANDLER_CODE
    , TARGET_WORKSHEET = @TARGET_WORKSHEET
    , MENU_ORDER = @MENU_ORDER
    , EDIT_PARAMETERS = @EDIT_PARAMETERS
WHERE
    TABLE_SCHEMA = @TABLE_SCHEMA
    AND TABLE_NAME = @TABLE_NAME
    AND COALESCE(COLUMN_NAME, '') = COALESCE(@COLUMN_NAME, '')
    AND EVENT_NAME = @EVENT_NAME
    AND COALESCE(HANDLER_SCHEMA, '') = COALESCE(@HANDLER_SCHEMA, '')
    AND COALESCE(HANDLER_NAME, '') = COALESCE(@HANDLER_NAME, '')
    AND COALESCE(HANDLER_TYPE, '') = COALESCE(@HANDLER_TYPE, '')

IF @@ROWCOUNT = 0
    INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS)
        VALUES (@TABLE_SCHEMA, @TABLE_NAME, @COLUMN_NAME, @EVENT_NAME, @HANDLER_SCHEMA, @HANDLER_NAME, @HANDLER_TYPE, @HANDLER_CODE, @TARGET_WORKSHEET, @MENU_ORDER, @EDIT_PARAMETERS)

END


GO

INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'handlers', N'HANDLER_CODE', N'DoNotConvertFormulas', NULL, NULL, N'ATTRIBUTE', NULL, NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'handlers', N'EVENT_NAME', N'ValidationList', NULL, NULL, N'VALUES', N'Actions, AddHyperlinks, AddStateColumn, BitColumn, Change, ContextMenu, ConvertFormulas, DataTypeBit, DataTypeBoolean, DataTypeDate, DataTypeDateTime, DataTypeDateTimeOffset, DataTypeDouble, DataTypeInt, DataTypeGuid, DataTypeString, DataTypeTime, DataTypeTimeSpan, DefaultListObject, DefaultValue, DependsOn, DoNotAddChangeHandler, DoNotAddDependsOn, DoNotAddManyToMany, DoNotAddValidation, DoNotChange, DoNotConvertFormulas, DoNotKeepComments, DoNotKeepFormulas, DoNotSave, DoNotSelect, DoNotSort, DoNotTranslate, DoubleClick, DynamicColumns, Format, Formula, FormulaValue, Information, JsonForm, KeepFormulas, KeepComments, License, ManyToMany, ParameterValues, ProtectRows, RegEx, SelectionChange, SelectionList, SelectPeriod, SyncParameter, UpdateChangedCellsOnly, UpdateEntireRow, ValidationList', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'handlers', N'HANDLER_TYPE', N'ValidationList', NULL, NULL, N'VALUES', N'TABLE, VIEW, PROCEDURE, FUNCTION, CODE, HTTP, TEXT, MACRO, CMD, VALUES, RANGE, REFRESH, MENUSEPARATOR, PDF, REPORT, SHOWSHEETS, HIDESHEETS, SELECTSHEET, ATTRIBUTE', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'objects', N'TABLE_TYPE', N'ValidationList', NULL, NULL, N'VALUES', N'TABLE, VIEW, PROCEDURE, FUNCTION, CODE, HTTP, TEXT, HIDDEN', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'base_tables', NULL, N'Actions', N'logs', N'E-book - Change Tracking Framework - Attaching Context Menus', N'HTTP', N'https://www.savetodb.com/change-tracking-framework/chapter-02.htm#_Toc497911553', NULL, 41, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'Actions', N'logs', N'MenuSeparator40', N'MENUSEPARATOR', NULL, NULL, 40, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'Actions', N'logs', N'E-book - Change Tracking Framework - Creating Change Tracking Triggers', N'HTTP', N'https://www.savetodb.com/change-tracking-framework/chapter-02.htm#_Toc497911551', NULL, 41, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'Actions', N'logs', N'xl_actions_drop_all_change_tracking_triggers', N'PROCEDURE', NULL, N'_reload', 11, 1);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'ContextMenu', N'logs', N'MenuSeparator20', N'MENUSEPARATOR', NULL, NULL, 20, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'ContextMenu', N'logs', N'xl_actions_clear_logs', N'PROCEDURE', NULL, N'_reload', 21, 1);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'ContextMenu', N'logs', N'xl_actions_create_change_tracking_triggers', N'PROCEDURE', NULL, N'_reload', 11, 1);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'ContextMenu', N'logs', N'xl_actions_drop_change_tracking_triggers', N'PROCEDURE', NULL, N'_reload', 12, 1);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'xl_actions_select_record', NULL, N'ContextMenu', N'logs', N'MenuSeparator20', N'MENUSEPARATOR', NULL, NULL, 20, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'xl_actions_select_record', NULL, N'ContextMenu', N'logs', N'MenuSeparator30', N'MENUSEPARATOR', NULL, NULL, 30, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'xl_actions_select_record', NULL, N'ContextMenu', N'logs', N'xl_actions_restore_current_record', N'PROCEDURE', NULL, N'_taskpane', 31, 1);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'xl_actions_select_record', NULL, N'ContextMenu', N'logs', N'xl_actions_restore_previous_record', N'PROCEDURE', NULL, N'_taskpane', 32, 1);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'xl_actions_select_record', NULL, N'ContextMenu', N'logs', N'xl_actions_select_lookup_id', N'PROCEDURE', NULL, N'_taskpane', 21, 0);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'base_tables', NULL, N'Actions', N'logs', N'MenuSeparator40', N'MENUSEPARATOR', NULL, NULL, 40, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'translations', NULL, N'Actions', N'logs', N'MenuSeparator40', N'MENUSEPARATOR', NULL, NULL, 40, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'translations', NULL, N'Actions', N'logs', N'E-book - Change Tracking Framework - Translations', N'HTTP', N'https://www.savetodb.com/change-tracking-framework-online.htm#_Toc497911559', NULL, 41, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'usp_translations', N'field', N'ParameterValues', NULL, NULL, N'VALUES', N'TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'workbooks', NULL, N'Actions', N'logs', N'SaveToDB Online Help - Configuring Application Workbooks', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-workbooks.htm', NULL, 42, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'workbooks', NULL, N'Actions', N'logs', N'MenuSeparator40', N'MENUSEPARATOR', NULL, NULL, 40, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'workbooks', NULL, N'Actions', N'logs', N'SaveToDB Online Help - Getting Workbook Definition', N'HTTP', N'https://www.savetodb.com/savetodb/dialog-box-workbook-information.htm', NULL, 43, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'workbooks', NULL, N'Actions', N'logs', N'SaveToDB Online Help - Application Workbooks Wizard', N'HTTP', N'https://www.savetodb.com/savetodb/wizard-application-workbooks.htm', NULL, 41, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'handlers', NULL, N'Actions', N'logs', N'MenuSeparator40', N'MENUSEPARATOR', NULL, NULL, 40, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'handlers', NULL, N'Actions', N'logs', N'SaveToDB Online Help - Configuring Event Handlers', N'HTTP', N'https://www.savetodb.com/dev-guide/context-actions.htm', NULL, 41, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'base_tables', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'base_tables', NULL, N'Actions', N'logs', N'Online Database Help - logs.base_tables', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-tables.htm#logs.base_tables', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'change_logs', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'change_logs', NULL, N'Actions', N'logs', N'Online Database Help - logs.change_logs', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-tables.htm#logs.change_logs', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'formats', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'formats', NULL, N'Actions', N'logs', N'Online Database Help - logs.formats', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-tables.htm#logs.formats', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'handlers', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'handlers', NULL, N'Actions', N'logs', N'Online Database Help - logs.handlers', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-tables.htm#logs.handlers', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'tables', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'tables', NULL, N'Actions', N'logs', N'Online Database Help - logs.tables', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-tables.htm#logs.tables', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'translations', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'translations', NULL, N'Actions', N'logs', N'Online Database Help - logs.translations', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-tables.htm#logs.translations', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'usp_translations', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'usp_translations', NULL, N'Actions', N'logs', N'Online Database Help - logs.usp_translations', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-procedures.htm#logs.usp_translations', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_administrator_handlers', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_administrator_handlers', NULL, N'Actions', N'logs', N'Online Database Help - logs.view_administrator_handlers', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-views.htm#logs.view_administrator_handlers', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_captured_objects', NULL, N'Actions', N'logs', N'Online Database Help - logs.view_captured_objects', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-views.htm#logs.view_captured_objects', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_query_list', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_query_list', NULL, N'Actions', N'logs', N'Online Database Help - logs.view_query_list', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-views.htm#logs.view_query_list', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_translations', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_translations', NULL, N'Actions', N'logs', N'Online Database Help - logs.view_translations', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-views.htm#logs.view_translations', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_user_handlers', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'view_user_handlers', NULL, N'Actions', N'logs', N'Online Database Help - logs.view_user_handlers', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-views.htm#logs.view_user_handlers', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'workbooks', NULL, N'Actions', N'logs', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'workbooks', NULL, N'Actions', N'logs', N'Online Database Help - logs.workbooks', N'HTTP', N'https://www.savetodb.com/help/change-tracking-framework-tables.htm#logs.workbooks', NULL, 91, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'base_tables', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'uPh4IfdgyDeKDe2t+UuS04EvfiaOfwwOmMrECQLUPx4Lv7pZ+VY1DsSAWahVckNfmLuCwcSHRMRkPVUBim6xExnnplYmDPL3H70PFGdWjQ6gRBUnVDpk16USSRzRzZN6RxHRbQKgiH7zoTlMIc8I8G6XPJMN72hmLL5CAri3YwY=', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'formats', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'EjdSU6TXqF66uNqVFE1NGeaJabhQ0nwwaIdYKcoPIDrIDRjM6U1s961DK7N0RicU3JNvGOrdAtdwbJ+6/LI75kgnapidzUtXBrx3rgTP4Y/IpW7vdISubEtKH2YimsaIdDsRm8InZo0t6mfZX4+ByrlJzsqBr2nZ44BrdTxEU/g=', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'handlers', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'mv4Mf5p3DsUqBjAyIcqV/QD1pRL7RQ/pJ3Hq0jxEw0dsvy5mlL4wMNHbAmruYUt0A3ta6OlG+x3h1wOO7yc1p1vzOjCM5uv127m1UuIFxCAR1ssbw7CcbUa2mekQTXAzZBDQHGyt/gTXX9nlAlX189hBK7+h37T9fX28wjpYDjM=', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'translations', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'EFglQsAqZ9PymMEje3CG+I+iPcEJupAblPbHdaFIxUHYnGG6SG92z1JML4U6EBeitdyUCQU6/Xxde4bDPdRYHASs2r4ugjJSfBve4tRdB+4Incth1bbGQ6EScIcCRJkZBGReNyd9oXly5GM3HRoLQLKuetLQ3ADHFCcJ64vKuJA=', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'usp_translations', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'K8PB5gwgoSm/ZA8hzm9g0sI3+caYn1s53a9GCkaWVIWJKjUfmyyF383Nzax5FLbvhYs6EoQp+6+kWt6B3KYlRP3hbL7DJ/S77HoCX/bGp5Xxdx2Hu1i5ggTXcM9PoSiwE3drSbzpEg5DDJSn+n/F+jQGz0iiEThKhhe3TYR/Dro=', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'workbooks', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'z3MFkOBHK9KXk5PSpMnd4Aq5x9VqArNqaV8NvcWQWVBQp783v85kdtYIL2PL2SlGcYsLYQL2PPcYXmSLIszu+pygp7rmtIn433KPQ+GVPLBNDXsCwkVGEEtXFcyB/LDPH8E1XWGi/eo6AStyDMaOgRVgd361ro8TB3ttcct5+4E=', NULL, NULL, NULL);
INSERT INTO logs.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'logs', N'change_tracking_framework', N'version', N'Information', NULL, NULL, N'ATTRIBUTE', N'4.2', NULL, NULL, NULL);

INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'All Changes', NULL, N'en', N'All Changes', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Deleted Rows', NULL, N'en', N'Deleted Rows', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Inserted Rows', NULL, N'en', N'Inserted Rows', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Updated Rows', NULL, N'en', N'Updated Rows', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_clear_logs', NULL, N'en', N'Clear Logs', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_create_change_tracking_triggers', NULL, N'en', N'Create Change Tracking Triggers', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_all_change_tracking_triggers', NULL, N'en', N'Drop All Change Tracking Triggers', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_change_tracking_triggers', NULL, N'en', N'Drop Change Tracking Triggers', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_current_record', NULL, N'en', N'Restore Current Record', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_previous_record', NULL, N'en', N'Restore Previous Record', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_lookup_id', NULL, N'en', N'ID Lookup', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', NULL, N'en', N'Row Changes', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', NULL, N'en', N'Change Log', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'deleted', N'en', N'deleted', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'inserted', N'en', N'inserted', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'updated', N'en', N'updated', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'name', N'en', N'Name', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'value', N'en', N'Value', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_date', N'en', N'Change Date', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_id', N'en', N'Change ID', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_type', N'en', N'Change Action', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_user', N'en', N'Change User', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'All Changes', NULL, N'fr', N'Tous les changements', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Deleted Rows', NULL, N'fr', N'Lignes supprimées', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Inserted Rows', NULL, N'fr', N'Lignes insérées', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Updated Rows', NULL, N'fr', N'Lignes modifiées', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_clear_logs', NULL, N'fr', N'Effacer les journaux', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_create_change_tracking_triggers', NULL, N'fr', N'Créer des déclencheurs de suivi des modifications', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_all_change_tracking_triggers', NULL, N'fr', N'Supprimer tous les déclencheurs de suivi des modifications', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_change_tracking_triggers', NULL, N'fr', N'Supprimer les déclencheurs de suivi des modifications', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_current_record', NULL, N'fr', N'Restaurer l''enregistrement en cours', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_previous_record', NULL, N'fr', N'Restaurer l''enregistrement précédent', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_lookup_id', NULL, N'fr', N'Recherche d''identifiant', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', NULL, N'fr', N'Changements de ligne', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', NULL, N'fr', N'Modifier le journal', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'deleted', N'fr', N'supprimé', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'inserted', N'fr', N'inséré', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'updated', N'fr', N'actualisé', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'name', N'fr', N'Nom', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'value', N'fr', N'Valeur', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_date', N'fr', N'Date de modification', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_id', N'fr', N'ID de modification', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_type', N'fr', N'Action de modification', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_user', N'fr', N'Utilisateur de modification', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'All Changes', NULL, N'it', N'Tutte le modifiche', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Deleted Rows', NULL, N'it', N'Righe eliminate', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Inserted Rows', NULL, N'it', N'Righe inserite', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Updated Rows', NULL, N'it', N'Righe aggiornate', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_clear_logs', NULL, N'it', N'Cancella i registri', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_create_change_tracking_triggers', NULL, N'it', N'Crea i trigger di tracciamento delle modifiche', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_all_change_tracking_triggers', NULL, N'it', N'Elimina tutti i trigger di tracciamento delle modifiche', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_change_tracking_triggers', NULL, N'it', N'Elimina i trigger di tracciamento delle modifiche', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_current_record', NULL, N'it', N'Ripristina il record corrente', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_previous_record', NULL, N'it', N'Ripristina il record precedente', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_lookup_id', NULL, N'it', N'Ricerca ID', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', NULL, N'it', N'Modifiche di riga', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', NULL, N'it', N'Cambia il registro', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'deleted', N'it', N'cancellato', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'inserted', N'it', N'inserito', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'updated', N'it', N'aggiornato', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'name', N'it', N'Nome', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'value', N'it', N'Valore', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_date', N'it', N'Data di modifica', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_id', N'it', N'ID di modifica', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_type', N'it', N'Azioni di modifica', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_user', N'it', N'Utente di modifica', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'All Changes', NULL, N'es', N'Todos los cambios', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Deleted Rows', NULL, N'es', N'Filas eliminadas', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Inserted Rows', NULL, N'es', N'Filas insertadas', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Updated Rows', NULL, N'es', N'Filas actualizadas', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_clear_logs', NULL, N'es', N'Borrar registros', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_create_change_tracking_triggers', NULL, N'es', N'Crear desencadenantes de seguimiento de cambios', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_all_change_tracking_triggers', NULL, N'es', N'Eliminar todos los desencadenantes de seguimiento de cambios', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_change_tracking_triggers', NULL, N'es', N'Eliminar desencadenantes de seguimiento de cambios', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_current_record', NULL, N'es', N'Restaurar registro actual', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_previous_record', NULL, N'es', N'Restaurar registro anterior', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_lookup_id', NULL, N'es', N'Búsqueda de identificación', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', NULL, N'es', N'Cambios de fila', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', NULL, N'es', N'Cambio de registro', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'deleted', N'es', N'borrado', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'inserted', N'es', N'insertado', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'updated', N'es', N'actualizado', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'name', N'es', N'Nombre', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'value', N'es', N'Valor', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_date', N'es', N'Fecha de modificación', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_id', N'es', N'ID de modificación', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_type', N'es', N'Tipo de modificación', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_user', N'es', N'Usuario de modificación', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'All Changes', NULL, N'ru', N'Все изменения', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Deleted Rows', NULL, N'ru', N'Удаленные строки', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Inserted Rows', NULL, N'ru', N'Добавленные строки', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Updated Rows', NULL, N'ru', N'Обновленные строки', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_clear_logs', NULL, N'ru', N'Очистить лог', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_create_change_tracking_triggers', NULL, N'ru', N'Создать триггеры отслеживания изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_all_change_tracking_triggers', NULL, N'ru', N'Удалить все триггеры отслеживания изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_change_tracking_triggers', NULL, N'ru', N'Удалить триггеры отслеживания изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_current_record', NULL, N'ru', N'Восстановить текущую запись', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_previous_record', NULL, N'ru', N'Восстановить предыдущую запись', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_lookup_id', NULL, N'ru', N'Поиск ID', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', NULL, N'ru', N'Изменения строки', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', NULL, N'ru', N'Журнал изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'deleted', N'ru', N'удаление', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'inserted', N'ru', N'вставка', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'updated', N'ru', N'изменение', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'name', N'ru', N'Имя', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'value', N'ru', N'Значение', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_date', N'ru', N'Дата изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_id', N'ru', N'ID изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_type', N'ru', N'Тип изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_user', N'ru', N'Пользователь изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'All Changes', NULL, N'de', N'Alle Änderungen', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Deleted Rows', NULL, N'de', N'Gelöschte Zeilen', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Inserted Rows', NULL, N'de', N'Eingefügte Zeilen', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Updated Rows', NULL, N'de', N'Aktualisierte Zeilen', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_clear_logs', NULL, N'de', N'Klare Protokolle', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_create_change_tracking_triggers', NULL, N'de', N'Erstellen Sie Änderungsverfolgungs Trigger', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_all_change_tracking_triggers', NULL, N'de', N'Entfernen Sie alle Änderungsverfolgungs Trigger', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_drop_change_tracking_triggers', NULL, N'de', N'Entfernen Sie Änderungsverfolgungs Trigger', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_current_record', NULL, N'de', N'Aktuellen Datensatz Wiederherstellen', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_restore_previous_record', NULL, N'de', N'Vorherigen Datensatz Wiederherstellen', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_lookup_id', NULL, N'de', N'ID-Suche', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', NULL, N'de', N'Zeilenwechsel', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', NULL, N'de', N'Änderungsprotokoll', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'deleted', N'de', N'gelöscht', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'inserted', N'de', N'eingefügt', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'change_type', N'updated', N'de', N'aktualisiert', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'name', N'de', N'Name', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_record', N'value', N'de', N'Wert', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_date', N'de', N'Datum ändern', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_id', N'de', N'ID ändern', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_type', N'de', N'Typ ändern', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'xl_actions_select_records', N'change_user', N'de', N'Benutzer ändern', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Create Change Tracking Triggers', NULL, N'en', N'Create Change Tracking Triggers', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Drop Change Tracking Triggers', NULL, N'en', N'Drop Change Tracking Triggers', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Create Change Tracking Triggers', NULL, N'fr', N'Créer des déclencheurs de suivi des modifications', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Create Change Tracking Triggers', NULL, N'it', N'Crea i trigger di tracciamento delle modifiche', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Create Change Tracking Triggers', NULL, N'es', N'Crear desencadenantes de seguimiento de cambios', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Create Change Tracking Triggers', NULL, N'de', N'Erstellen Sie Änderungsverfolgungs Trigger', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Create Change Tracking Triggers', NULL, N'ru', N'Создать триггеры отслеживания изменений', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Drop Change Tracking Triggers', NULL, N'fr', N'Supprimer les déclencheurs de suivi des modifications', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Drop Change Tracking Triggers', NULL, N'it', N'Elimina i trigger di tracciamento delle modifiche', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Drop Change Tracking Triggers', NULL, N'es', N'Eliminar desencadenantes de seguimiento de cambios', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Drop Change Tracking Triggers', NULL, N'de', N'Entfernen Sie Änderungsverfolgungs Trigger', NULL, NULL);
INSERT INTO logs.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES (N'logs', N'Drop Change Tracking Triggers', NULL, N'ru', N'Удалить триггеры отслеживания изменений', NULL, NULL);

INSERT INTO logs.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'logs', N'base_tables', N'<table name="logs.base_tables"><columnFormats><column name="" property="ListObjectName" value="attached_handlers" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="ID" property="Address" value="$C$4" type="String"/><column name="ID" property="NumberFormat" value="General" type="String"/><column name="ID" property="Validation.Type" value="1" type="Double"/><column name="ID" property="Validation.Operator" value="1" type="Double"/><column name="ID" property="Validation.Formula1" value="-2147483648" type="String"/><column name="ID" property="Validation.Formula2" value="2147483647" type="String"/><column name="ID" property="Validation.AlertStyle" value="1" type="Double"/><column name="ID" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="ID" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="ID" property="Validation.ShowInput" value="True" type="Boolean"/><column name="ID" property="Validation.ShowError" value="True" type="Boolean"/><column name="OBJECT_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="OBJECT_SCHEMA" property="Address" value="$D$4" type="String"/><column name="OBJECT_SCHEMA" property="ColumnWidth" value="17.71" type="Double"/><column name="OBJECT_SCHEMA" property="NumberFormat" value="General" type="String"/><column name="OBJECT_SCHEMA" property="Validation.Type" value="6" type="Double"/><column name="OBJECT_SCHEMA" property="Validation.Operator" value="8" type="Double"/><column name="OBJECT_SCHEMA" property="Validation.Formula1" value="128" type="String"/><column name="OBJECT_SCHEMA" property="Validation.AlertStyle" value="1" type="Double"/><column name="OBJECT_SCHEMA" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="OBJECT_SCHEMA" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="OBJECT_SCHEMA" property="Validation.ShowInput" value="True" type="Boolean"/><column name="OBJECT_SCHEMA" property="Validation.ShowError" value="True" type="Boolean"/><column name="OBJECT_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="OBJECT_NAME" property="Address" value="$E$4" type="String"/><column name="OBJECT_NAME" property="ColumnWidth" value="19.29" type="Double"/><column name="OBJECT_NAME" property="NumberFormat" value="General" type="String"/><column name="OBJECT_NAME" property="Validation.Type" value="6" type="Double"/><column name="OBJECT_NAME" property="Validation.Operator" value="8" type="Double"/><column name="OBJECT_NAME" property="Validation.Formula1" value="128" type="String"/><column name="OBJECT_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="OBJECT_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="OBJECT_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="OBJECT_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="OBJECT_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="BASE_TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="BASE_TABLE_SCHEMA" property="Address" value="$F$4" type="String"/><column name="BASE_TABLE_SCHEMA" property="ColumnWidth" value="22.14" type="Double"/><column name="BASE_TABLE_SCHEMA" property="NumberFormat" value="General" type="String"/><column name="BASE_TABLE_SCHEMA" property="Validation.Type" value="6" type="Double"/><column name="BASE_TABLE_SCHEMA" property="Validation.Operator" value="8" type="Double"/><column name="BASE_TABLE_SCHEMA" property="Validation.Formula1" value="128" type="String"/><column name="BASE_TABLE_SCHEMA" property="Validation.AlertStyle" value="1" type="Double"/><column name="BASE_TABLE_SCHEMA" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="BASE_TABLE_SCHEMA" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="BASE_TABLE_SCHEMA" property="Validation.ShowInput" value="True" type="Boolean"/><column name="BASE_TABLE_SCHEMA" property="Validation.ShowError" value="True" type="Boolean"/><column name="BASE_TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="BASE_TABLE_NAME" property="Address" value="$G$4" type="String"/><column name="BASE_TABLE_NAME" property="ColumnWidth" value="20.14" type="Double"/><column name="BASE_TABLE_NAME" property="NumberFormat" value="General" type="String"/><column name="BASE_TABLE_NAME" property="Validation.Type" value="6" type="Double"/><column name="BASE_TABLE_NAME" property="Validation.Operator" value="8" type="Double"/><column name="BASE_TABLE_NAME" property="Validation.Formula1" value="128" type="String"/><column name="BASE_TABLE_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="BASE_TABLE_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="BASE_TABLE_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="BASE_TABLE_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="BASE_TABLE_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="" property="Tab.Color" value="5287936" type="Double"/><column name="" property="Tab.Color" value="5287936" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
INSERT INTO logs.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'logs', N'view_captured_objects', N'<table name="logs.view_captured_objects"><columnFormats><column name="" property="ListObjectName" value="tracked_objects" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="object_id" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="object_id" property="Address" value="$C$4" type="String"/><column name="object_id" property="NumberFormat" value="General" type="String"/><column name="schema" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="schema" property="Address" value="$D$4" type="String"/><column name="schema" property="ColumnWidth" value="9.29" type="Double"/><column name="schema" property="NumberFormat" value="General" type="String"/><column name="name" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="name" property="Address" value="$E$4" type="String"/><column name="name" property="ColumnWidth" value="21.14" type="Double"/><column name="name" property="NumberFormat" value="General" type="String"/><column name="insert_trigger" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="insert_trigger" property="Address" value="$F$4" type="String"/><column name="insert_trigger" property="ColumnWidth" value="26" type="Double"/><column name="insert_trigger" property="NumberFormat" value="General" type="String"/><column name="update_trigger" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="update_trigger" property="Address" value="$G$4" type="String"/><column name="update_trigger" property="ColumnWidth" value="27.14" type="Double"/><column name="update_trigger" property="NumberFormat" value="General" type="String"/><column name="delete_trigger" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="delete_trigger" property="Address" value="$H$4" type="String"/><column name="delete_trigger" property="ColumnWidth" value="26.71" type="Double"/><column name="delete_trigger" property="NumberFormat" value="General" type="String"/><column name="select_handler" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="select_handler" property="Address" value="$I$4" type="String"/><column name="select_handler" property="ColumnWidth" value="15.86" type="Double"/><column name="select_handler" property="NumberFormat" value="General" type="String"/><column name="has_log_triggers" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="has_log_triggers" property="Address" value="$J$4" type="String"/><column name="has_log_triggers" property="ColumnWidth" value="17.14" type="Double"/><column name="has_log_triggers" property="NumberFormat" value="General" type="String"/><column name="object_id" property="FormatConditions(1).ColumnsCount" value="3" type="Double"/><column name="object_id" property="FormatConditions(1).AppliesTo.Address" value="$C$4:$E$38" type="String"/><column name="object_id" property="FormatConditions(1).Type" value="2" type="Double"/><column name="object_id" property="FormatConditions(1).Priority" value="1" type="Double"/><column name="object_id" property="FormatConditions(1).Formula1" value="=ISBLANK(C4)" type="String"/><column name="object_id" property="FormatConditions(1).Interior.Color" value="65535" type="Double"/><column name="object_id" property="FormatConditions(1).Interior.Color" value="65535" type="Double"/><column name="SortFields(1)" property="KeyfieldName" value="schema" type="String"/><column name="SortFields(1)" property="SortOn" value="0" type="Double"/><column name="SortFields(1)" property="Order" value="1" type="Double"/><column name="SortFields(1)" property="DataOption" value="0" type="Double"/><column name="SortFields(2)" property="KeyfieldName" value="name" type="String"/><column name="SortFields(2)" property="SortOn" value="0" type="Double"/><column name="SortFields(2)" property="Order" value="1" type="Double"/><column name="SortFields(2)" property="DataOption" value="0" type="Double"/><column name="" property="Tab.Color" value="5287936" type="Double"/><column name="" property="Tab.Color" value="5287936" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
INSERT INTO logs.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'logs', N'handlers', N'<table name="logs.handlers"><columnFormats><column name="" property="ListObjectName" value="handlers" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="ID" property="Address" value="$C$4" type="String"/><column name="ID" property="NumberFormat" value="General" type="String"/><column name="ID" property="Validation.Type" value="1" type="Double"/><column name="ID" property="Validation.Operator" value="1" type="Double"/><column name="ID" property="Validation.Formula1" value="-2147483648" type="String"/><column name="ID" property="Validation.Formula2" value="2147483647" type="String"/><column name="ID" property="Validation.AlertStyle" value="1" type="Double"/><column name="ID" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="ID" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="ID" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="ID" property="Validation.ErrorMessage" value="The column requires values of the int datatype." type="String"/><column name="ID" property="Validation.ShowInput" value="True" type="Boolean"/><column name="ID" property="Validation.ShowError" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_SCHEMA" property="Address" value="$D$4" type="String"/><column name="TABLE_SCHEMA" property="ColumnWidth" value="4.86" type="Double"/><column name="TABLE_SCHEMA" property="NumberFormat" value="General" type="String"/><column name="TABLE_SCHEMA" property="Validation.Type" value="6" type="Double"/><column name="TABLE_SCHEMA" property="Validation.Operator" value="8" type="Double"/><column name="TABLE_SCHEMA" property="Validation.Formula1" value="20" type="String"/><column name="TABLE_SCHEMA" property="Validation.AlertStyle" value="1" type="Double"/><column name="TABLE_SCHEMA" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="TABLE_SCHEMA" property="Validation.ErrorMessage" value="The column requires values of the nvarchar(20) datatype." type="String"/><column name="TABLE_SCHEMA" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="Validation.ShowError" value="True" type="Boolean"/><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_NAME" property="Address" value="$E$4" type="String"/><column name="TABLE_NAME" property="ColumnWidth" value="22.57" type="Double"/><column name="TABLE_NAME" property="NumberFormat" value="General" type="String"/><column name="TABLE_NAME" property="Validation.Type" value="6" type="Double"/><column name="TABLE_NAME" property="Validation.Operator" value="8" type="Double"/><column name="TABLE_NAME" property="Validation.Formula1" value="128" type="String"/><column name="TABLE_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="TABLE_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TABLE_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TABLE_NAME" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="TABLE_NAME" property="Validation.ErrorMessage" value="The column requires values of the nvarchar(128) datatype." type="String"/><column name="TABLE_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TABLE_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="COLUMN_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="COLUMN_NAME" property="Address" value="$F$4" type="String"/><column name="COLUMN_NAME" property="ColumnWidth" value="17.43" type="Double"/><column name="COLUMN_NAME" property="NumberFormat" value="General" type="String"/><column name="COLUMN_NAME" property="Validation.Type" value="6" type="Double"/><column name="COLUMN_NAME" property="Validation.Operator" value="8" type="Double"/><column name="COLUMN_NAME" property="Validation.Formula1" value="128" type="String"/><column name="COLUMN_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="COLUMN_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="COLUMN_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="COLUMN_NAME" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="COLUMN_NAME" property="Validation.ErrorMessage" value="The column requires values of the nvarchar(128) datatype." type="String"/><column name="COLUMN_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="COLUMN_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="EVENT_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="EVENT_NAME" property="Address" value="$G$4" type="String"/><column name="EVENT_NAME" property="ColumnWidth" value="21.57" type="Double"/><column name="EVENT_NAME" property="NumberFormat" value="General" type="String"/><column name="EVENT_NAME" property="Validation.Type" value="3" type="Double"/><column name="EVENT_NAME" property="Validation.Operator" value="1" type="Double"/><column name="EVENT_NAME" property="Validation.Formula1" value="=INDIRECT(&quot;vl_logs_handlers_event_name&quot;)" type="String"/><column name="EVENT_NAME" property="Validation.Formula2" value="" type="String"/><column name="EVENT_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="EVENT_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="EVENT_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="EVENT_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="EVENT_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="HANDLER_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="HANDLER_SCHEMA" property="Address" value="$H$4" type="String"/><column name="HANDLER_SCHEMA" property="ColumnWidth" value="5" type="Double"/><column name="HANDLER_SCHEMA" property="NumberFormat" value="General" type="String"/><column name="HANDLER_SCHEMA" property="Validation.Type" value="6" type="Double"/><column name="HANDLER_SCHEMA" property="Validation.Operator" value="8" type="Double"/><column name="HANDLER_SCHEMA" property="Validation.Formula1" value="20" type="String"/><column name="HANDLER_SCHEMA" property="Validation.AlertStyle" value="1" type="Double"/><column name="HANDLER_SCHEMA" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="HANDLER_SCHEMA" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="HANDLER_SCHEMA" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="HANDLER_SCHEMA" property="Validation.ErrorMessage" value="The column requires values of the nvarchar(20) datatype." type="String"/><column name="HANDLER_SCHEMA" property="Validation.ShowInput" value="True" type="Boolean"/><column name="HANDLER_SCHEMA" property="Validation.ShowError" value="True" type="Boolean"/><column name="HANDLER_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="HANDLER_NAME" property="Address" value="$I$4" type="String"/><column name="HANDLER_NAME" property="ColumnWidth" value="65.71" type="Double"/><column name="HANDLER_NAME" property="NumberFormat" value="General" type="String"/><column name="HANDLER_NAME" property="Validation.Type" value="6" type="Double"/><column name="HANDLER_NAME" property="Validation.Operator" value="8" type="Double"/><column name="HANDLER_NAME" property="Validation.Formula1" value="128" type="String"/><column name="HANDLER_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="HANDLER_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="HANDLER_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="HANDLER_NAME" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="HANDLER_NAME" property="Validation.ErrorMessage" value="The column requires values of the nvarchar(128) datatype." type="String"/><column name="HANDLER_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="HANDLER_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="HANDLER_TYPE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="HANDLER_TYPE" property="Address" value="$J$4" type="String"/><column name="HANDLER_TYPE" property="ColumnWidth" value="16.29" type="Double"/><column name="HANDLER_TYPE" property="NumberFormat" value="General" type="String"/><column name="HANDLER_TYPE" property="Validation.Type" value="3" type="Double"/><column name="HANDLER_TYPE" property="Validation.Operator" value="1" type="Double"/><column name="HANDLER_TYPE" property="Validation.Formula1" value="TABLE; VIEW; PROCEDURE; FUNCTION; CODE; HTTP; TEXT; MACRO; CMD; VALUES; RANGE; REFRESH; MENUSEPARATOR; PDF; REPORT" type="String"/><column name="HANDLER_TYPE" property="Validation.Formula2" value="" type="String"/><column name="HANDLER_TYPE" property="Validation.AlertStyle" value="1" type="Double"/><column name="HANDLER_TYPE" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="HANDLER_TYPE" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="HANDLER_TYPE" property="Validation.ShowInput" value="True" type="Boolean"/><column name="HANDLER_TYPE" property="Validation.ShowError" value="True" type="Boolean"/><column name="HANDLER_CODE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="HANDLER_CODE" property="Address" value="$K$4" type="String"/><column name="HANDLER_CODE" property="ColumnWidth" value="20.86" type="Double"/><column name="HANDLER_CODE" property="NumberFormat" value="General" type="String"/><column name="TARGET_WORKSHEET" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TARGET_WORKSHEET" property="Address" value="$L$4" type="String"/><column name="TARGET_WORKSHEET" property="ColumnWidth" value="21.71" type="Double"/><column name="TARGET_WORKSHEET" property="NumberFormat" value="General" type="String"/><column name="TARGET_WORKSHEET" property="Validation.Type" value="6" type="Double"/><column name="TARGET_WORKSHEET" property="Validation.Operator" value="8" type="Double"/><column name="TARGET_WORKSHEET" property="Validation.Formula1" value="128" type="String"/><column name="TARGET_WORKSHEET" property="Validation.AlertStyle" value="1" type="Double"/><column name="TARGET_WORKSHEET" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TARGET_WORKSHEET" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TARGET_WORKSHEET" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="TARGET_WORKSHEET" property="Validation.ErrorMessage" value="The column requires values of the nvarchar(128) datatype." type="String"/><column name="TARGET_WORKSHEET" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TARGET_WORKSHEET" property="Validation.ShowError" value="True" type="Boolean"/><column name="MENU_ORDER" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="MENU_ORDER" property="Address" value="$M$4" type="String"/><column name="MENU_ORDER" property="ColumnWidth" value="15.43" type="Double"/><column name="MENU_ORDER" property="NumberFormat" value="General" type="String"/><column name="MENU_ORDER" property="Validation.Type" value="1" type="Double"/><column name="MENU_ORDER" property="Validation.Operator" value="1" type="Double"/><column name="MENU_ORDER" property="Validation.Formula1" value="-2147483648" type="String"/><column name="MENU_ORDER" property="Validation.Formula2" value="2147483647" type="String"/><column name="MENU_ORDER" property="Validation.AlertStyle" value="1" type="Double"/><column name="MENU_ORDER" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="MENU_ORDER" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="MENU_ORDER" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="MENU_ORDER" property="Validation.ErrorMessage" value="The column requires values of the int datatype." type="String"/><column name="MENU_ORDER" property="Validation.ShowInput" value="True" type="Boolean"/><column name="MENU_ORDER" property="Validation.ShowError" value="True" type="Boolean"/><column name="EDIT_PARAMETERS" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="EDIT_PARAMETERS" property="Address" value="$N$4" type="String"/><column name="EDIT_PARAMETERS" property="ColumnWidth" value="19.57" type="Double"/><column name="EDIT_PARAMETERS" property="NumberFormat" value="General" type="String"/><column name="EDIT_PARAMETERS" property="HorizontalAlignment" value="-4108" type="Double"/><column name="EDIT_PARAMETERS" property="Font.Size" value="10" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).AppliesTo.Address" value="$N$4:$N$31" type="String"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).Type" value="6" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).Priority" value="3" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).ShowIconOnly" value="True" type="Boolean"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconSet.ID" value="8" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconCriteria(1).Type" value="3" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconCriteria(1).Operator" value="7" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconCriteria(2).Type" value="0" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconCriteria(2).Value" value="0.5" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconCriteria(2).Operator" value="7" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconCriteria(3).Type" value="0" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconCriteria(3).Value" value="1" type="Double"/><column name="EDIT_PARAMETERS" property="FormatConditions(1).IconCriteria(3).Operator" value="7" type="Double"/><column name="SortFields(1)" property="KeyfieldName" value="EVENT_NAME" type="String"/><column name="SortFields(1)" property="SortOn" value="0" type="Double"/><column name="SortFields(1)" property="Order" value="1" type="Double"/><column name="SortFields(1)" property="DataOption" value="0" type="Double"/><column name="SortFields(2)" property="KeyfieldName" value="TABLE_SCHEMA" type="String"/><column name="SortFields(2)" property="SortOn" value="0" type="Double"/><column name="SortFields(2)" property="Order" value="1" type="Double"/><column name="SortFields(2)" property="DataOption" value="0" type="Double"/><column name="SortFields(3)" property="KeyfieldName" value="TABLE_NAME" type="String"/><column name="SortFields(3)" property="SortOn" value="0" type="Double"/><column name="SortFields(3)" property="Order" value="1" type="Double"/><column name="SortFields(3)" property="DataOption" value="0" type="Double"/><column name="SortFields(4)" property="KeyfieldName" value="MENU_ORDER" type="String"/><column name="SortFields(4)" property="SortOn" value="0" type="Double"/><column name="SortFields(4)" property="Order" value="1" type="Double"/><column name="SortFields(4)" property="DataOption" value="0" type="Double"/><column name="SortFields(5)" property="KeyfieldName" value="HANDLER_SCHEMA" type="String"/><column name="SortFields(5)" property="SortOn" value="0" type="Double"/><column name="SortFields(5)" property="Order" value="1" type="Double"/><column name="SortFields(5)" property="DataOption" value="0" type="Double"/><column name="SortFields(6)" property="KeyfieldName" value="HANDLER_NAME" type="String"/><column name="SortFields(6)" property="SortOn" value="0" type="Double"/><column name="SortFields(6)" property="Order" value="1" type="Double"/><column name="SortFields(6)" property="DataOption" value="0" type="Double"/><column name="" property="Tab.Color" value="6299648" type="Double"/><column name="" property="Tab.Color" value="6299648" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
INSERT INTO logs.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'logs', N'workbooks', N'<table name="logs.workbooks"><columnFormats><column name="" property="ListObjectName" value="workbooks" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="_RowNum" property="VerticalAlignment" value="-4160" type="Double"/><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="ID" property="Address" value="$C$4" type="String"/><column name="ID" property="NumberFormat" value="General" type="String"/><column name="ID" property="VerticalAlignment" value="-4160" type="Double"/><column name="ID" property="Validation.Type" value="1" type="Double"/><column name="ID" property="Validation.Operator" value="1" type="Double"/><column name="ID" property="Validation.Formula1" value="-2147483648" type="String"/><column name="ID" property="Validation.Formula2" value="2147483647" type="String"/><column name="ID" property="Validation.AlertStyle" value="1" type="Double"/><column name="ID" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="ID" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="ID" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="ID" property="Validation.ErrorMessage" value="The column requires values of the int datatype." type="String"/><column name="ID" property="Validation.ShowInput" value="True" type="Boolean"/><column name="ID" property="Validation.ShowError" value="True" type="Boolean"/><column name="NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="NAME" property="Address" value="$D$4" type="String"/><column name="NAME" property="ColumnWidth" value="27.86" type="Double"/><column name="NAME" property="NumberFormat" value="General" type="String"/><column name="NAME" property="VerticalAlignment" value="-4160" type="Double"/><column name="NAME" property="Validation.Type" value="6" type="Double"/><column name="NAME" property="Validation.Operator" value="8" type="Double"/><column name="NAME" property="Validation.Formula1" value="128" type="String"/><column name="NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="NAME" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="NAME" property="Validation.ErrorMessage" value="The column requires values of the nvarchar(128) datatype." type="String"/><column name="NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="TEMPLATE" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TEMPLATE" property="Address" value="$E$4" type="String"/><column name="TEMPLATE" property="ColumnWidth" value="15.71" type="Double"/><column name="TEMPLATE" property="NumberFormat" value="General" type="String"/><column name="TEMPLATE" property="VerticalAlignment" value="-4160" type="Double"/><column name="TEMPLATE" property="Validation.Type" value="6" type="Double"/><column name="TEMPLATE" property="Validation.Operator" value="8" type="Double"/><column name="TEMPLATE" property="Validation.Formula1" value="255" type="String"/><column name="TEMPLATE" property="Validation.AlertStyle" value="1" type="Double"/><column name="TEMPLATE" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TEMPLATE" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TEMPLATE" property="Validation.ErrorTitle" value="Datatype Control" type="String"/><column name="TEMPLATE" property="Validation.ErrorMessage" value="The column requires values of the nvarchar(255) datatype." type="String"/><column name="TEMPLATE" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TEMPLATE" property="Validation.ShowError" value="True" type="Boolean"/><column name="DEFINITION" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="DEFINITION" property="Address" value="$F$4" type="String"/><column name="DEFINITION" property="ColumnWidth" value="151.14" type="Double"/><column name="DEFINITION" property="NumberFormat" value="General" type="String"/><column name="DEFINITION" property="VerticalAlignment" value="-4160" type="Double"/><column name="DEFINITION" property="WrapText" value="True" type="Boolean"/><column name="" property="Tab.Color" value="6299648" type="Double"/><column name="" property="Tab.Color" value="6299648" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
INSERT INTO logs.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'logs', N'usp_translations', N'<table name="logs.usp_translations"><columnFormats><column name="" property="ListObjectName" value="all_translations" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_SCHEMA" property="Address" value="$C$4" type="String"/><column name="TABLE_SCHEMA" property="ColumnWidth" value="4.14" type="Double"/><column name="TABLE_SCHEMA" property="NumberFormat" value="General" type="String"/><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_NAME" property="Address" value="$D$4" type="String"/><column name="TABLE_NAME" property="ColumnWidth" value="40.57" type="Double"/><column name="TABLE_NAME" property="NumberFormat" value="General" type="String"/><column name="COLUMN" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="COLUMN" property="Address" value="$E$4" type="String"/><column name="COLUMN" property="ColumnWidth" value="11.57" type="Double"/><column name="COLUMN" property="NumberFormat" value="General" type="String"/><column name="en" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="en" property="Address" value="$F$4" type="String"/><column name="en" property="ColumnWidth" value="30.29" type="Double"/><column name="en" property="NumberFormat" value="General" type="String"/><column name="fr" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="fr" property="Address" value="$G$4" type="String"/><column name="fr" property="ColumnWidth" value="54.71" type="Double"/><column name="fr" property="NumberFormat" value="General" type="String"/><column name="it" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="it" property="Address" value="$H$4" type="String"/><column name="it" property="ColumnWidth" value="48.86" type="Double"/><column name="it" property="NumberFormat" value="General" type="String"/><column name="es" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="es" property="Address" value="$I$4" type="String"/><column name="es" property="ColumnWidth" value="57.14" type="Double"/><column name="es" property="NumberFormat" value="General" type="String"/><column name="de" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="de" property="Address" value="$J$4" type="String"/><column name="de" property="ColumnWidth" value="44.14" type="Double"/><column name="de" property="NumberFormat" value="General" type="String"/><column name="ru" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="ru" property="Address" value="$K$4" type="String"/><column name="ru" property="ColumnWidth" value="45.14" type="Double"/><column name="ru" property="NumberFormat" value="General" type="String"/><column name="" property="Tab.Color" value="6299648" type="Double"/><column name="" property="Tab.Color" value="6299648" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/></columnFormats></table>');
INSERT INTO logs.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES (N'logs', N'translations', N'<table name="logs.translations"><columnFormats><column name="" property="ListObjectName" value="translations" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String"/><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean"/><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean"/><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean"/><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="_RowNum" property="Address" value="$B$4" type="String"/><column name="_RowNum" property="NumberFormat" value="General" type="String"/><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="ID" property="Address" value="$C$4" type="String"/><column name="ID" property="NumberFormat" value="General" type="String"/><column name="ID" property="Validation.Type" value="1" type="Double"/><column name="ID" property="Validation.Operator" value="1" type="Double"/><column name="ID" property="Validation.Formula1" value="-2147483648" type="String"/><column name="ID" property="Validation.Formula2" value="2147483647" type="String"/><column name="ID" property="Validation.AlertStyle" value="1" type="Double"/><column name="ID" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="ID" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="ID" property="Validation.ShowInput" value="True" type="Boolean"/><column name="ID" property="Validation.ShowError" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_SCHEMA" property="Address" value="$D$4" type="String"/><column name="TABLE_SCHEMA" property="ColumnWidth" value="9.43" type="Double"/><column name="TABLE_SCHEMA" property="NumberFormat" value="General" type="String"/><column name="TABLE_SCHEMA" property="Validation.Type" value="6" type="Double"/><column name="TABLE_SCHEMA" property="Validation.Operator" value="8" type="Double"/><column name="TABLE_SCHEMA" property="Validation.Formula1" value="128" type="String"/><column name="TABLE_SCHEMA" property="Validation.AlertStyle" value="1" type="Double"/><column name="TABLE_SCHEMA" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="Validation.ShowError" value="True" type="Boolean"/><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_NAME" property="Address" value="$E$4" type="String"/><column name="TABLE_NAME" property="ColumnWidth" value="41.57" type="Double"/><column name="TABLE_NAME" property="NumberFormat" value="General" type="String"/><column name="TABLE_NAME" property="Validation.Type" value="6" type="Double"/><column name="TABLE_NAME" property="Validation.Operator" value="8" type="Double"/><column name="TABLE_NAME" property="Validation.Formula1" value="128" type="String"/><column name="TABLE_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="TABLE_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TABLE_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TABLE_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TABLE_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="COLUMN_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="COLUMN_NAME" property="Address" value="$F$4" type="String"/><column name="COLUMN_NAME" property="ColumnWidth" value="25.57" type="Double"/><column name="COLUMN_NAME" property="NumberFormat" value="General" type="String"/><column name="COLUMN_NAME" property="Validation.Type" value="6" type="Double"/><column name="COLUMN_NAME" property="Validation.Operator" value="8" type="Double"/><column name="COLUMN_NAME" property="Validation.Formula1" value="128" type="String"/><column name="COLUMN_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="COLUMN_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="COLUMN_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="COLUMN_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="COLUMN_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="LANGUAGE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="LANGUAGE_NAME" property="Address" value="$G$4" type="String"/><column name="LANGUAGE_NAME" property="ColumnWidth" value="19.57" type="Double"/><column name="LANGUAGE_NAME" property="NumberFormat" value="General" type="String"/><column name="LANGUAGE_NAME" property="Validation.Type" value="6" type="Double"/><column name="LANGUAGE_NAME" property="Validation.Operator" value="8" type="Double"/><column name="LANGUAGE_NAME" property="Validation.Formula1" value="2" type="String"/><column name="LANGUAGE_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="LANGUAGE_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="LANGUAGE_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="LANGUAGE_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="LANGUAGE_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="TRANSLATED_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_NAME" property="Address" value="$H$4" type="String"/><column name="TRANSLATED_NAME" property="ColumnWidth" value="30.29" type="Double"/><column name="TRANSLATED_NAME" property="NumberFormat" value="@" type="String"/><column name="TRANSLATED_NAME" property="Validation.Type" value="6" type="Double"/><column name="TRANSLATED_NAME" property="Validation.Operator" value="8" type="Double"/><column name="TRANSLATED_NAME" property="Validation.Formula1" value="128" type="String"/><column name="TRANSLATED_NAME" property="Validation.AlertStyle" value="1" type="Double"/><column name="TRANSLATED_NAME" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TRANSLATED_NAME" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TRANSLATED_NAME" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TRANSLATED_NAME" property="Validation.ShowError" value="True" type="Boolean"/><column name="TRANSLATED_DESC" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_DESC" property="Address" value="$I$4" type="String"/><column name="TRANSLATED_DESC" property="ColumnWidth" value="19.57" type="Double"/><column name="TRANSLATED_DESC" property="NumberFormat" value="@" type="String"/><column name="TRANSLATED_DESC" property="Validation.Type" value="6" type="Double"/><column name="TRANSLATED_DESC" property="Validation.Operator" value="8" type="Double"/><column name="TRANSLATED_DESC" property="Validation.Formula1" value="1024" type="String"/><column name="TRANSLATED_DESC" property="Validation.AlertStyle" value="1" type="Double"/><column name="TRANSLATED_DESC" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TRANSLATED_DESC" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TRANSLATED_DESC" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TRANSLATED_DESC" property="Validation.ShowError" value="True" type="Boolean"/><column name="TRANSLATED_COMMENT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_COMMENT" property="Address" value="$J$4" type="String"/><column name="TRANSLATED_COMMENT" property="ColumnWidth" value="25" type="Double"/><column name="TRANSLATED_COMMENT" property="NumberFormat" value="@" type="String"/><column name="TRANSLATED_COMMENT" property="Validation.Type" value="6" type="Double"/><column name="TRANSLATED_COMMENT" property="Validation.Operator" value="8" type="Double"/><column name="TRANSLATED_COMMENT" property="Validation.Formula1" value="2000" type="String"/><column name="TRANSLATED_COMMENT" property="Validation.AlertStyle" value="1" type="Double"/><column name="TRANSLATED_COMMENT" property="Validation.IgnoreBlank" value="True" type="Boolean"/><column name="TRANSLATED_COMMENT" property="Validation.InCellDropdown" value="True" type="Boolean"/><column name="TRANSLATED_COMMENT" property="Validation.ShowInput" value="True" type="Boolean"/><column name="TRANSLATED_COMMENT" property="Validation.ShowError" value="True" type="Boolean"/><column name="SortFields(1)" property="KeyfieldName" value="COLUMN_NAME" type="String"/><column name="SortFields(1)" property="SortOn" value="0" type="Double"/><column name="SortFields(1)" property="Order" value="1" type="Double"/><column name="SortFields(1)" property="DataOption" value="0" type="Double"/><column name="SortFields(2)" property="KeyfieldName" value="LANGUAGE_NAME" type="String"/><column name="SortFields(2)" property="SortOn" value="0" type="Double"/><column name="SortFields(2)" property="Order" value="1" type="Double"/><column name="SortFields(2)" property="DataOption" value="0" type="Double"/><column name="SortFields(3)" property="KeyfieldName" value="TABLE_SCHEMA" type="String"/><column name="SortFields(3)" property="SortOn" value="0" type="Double"/><column name="SortFields(3)" property="Order" value="1" type="Double"/><column name="SortFields(3)" property="DataOption" value="0" type="Double"/><column name="SortFields(4)" property="KeyfieldName" value="TABLE_NAME" type="String"/><column name="SortFields(4)" property="SortOn" value="0" type="Double"/><column name="SortFields(4)" property="Order" value="1" type="Double"/><column name="SortFields(4)" property="DataOption" value="0" type="Double"/><column name="" property="Tab.Color" value="6299648" type="Double"/><column name="" property="Tab.Color" value="6299648" type="Double"/><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean"/><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean"/><column name="" property="ActiveWindow.Split" value="True" type="Boolean"/><column name="" property="ActiveWindow.SplitRow" value="0" type="Double"/><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double"/><column name="" property="PageSetup.Orientation" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double"/><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double"/><column name="" property="PageSetup.PaperSize" value="9" type="Double"/></columnFormats><views><view name="By ID"><column name="" property="ListObjectName" value="translations" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="COLUMN_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="LANGUAGE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_DESC" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_COMMENT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="SortFields(1)" property="KeyfieldName" value="ID" type="String"/><column name="SortFields(1)" property="SortOn" value="0" type="Double"/><column name="SortFields(1)" property="Order" value="1" type="Double"/><column name="SortFields(1)" property="DataOption" value="0" type="Double"/></view><view name="By Object"><column name="" property="ListObjectName" value="translations" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="COLUMN_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="LANGUAGE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_DESC" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_COMMENT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="SortFields(1)" property="KeyfieldName" value="LANGUAGE_NAME" type="String"/><column name="SortFields(1)" property="SortOn" value="0" type="Double"/><column name="SortFields(1)" property="Order" value="1" type="Double"/><column name="SortFields(1)" property="DataOption" value="0" type="Double"/><column name="SortFields(2)" property="KeyfieldName" value="TABLE_SCHEMA" type="String"/><column name="SortFields(2)" property="SortOn" value="0" type="Double"/><column name="SortFields(2)" property="Order" value="1" type="Double"/><column name="SortFields(2)" property="DataOption" value="0" type="Double"/><column name="SortFields(3)" property="KeyfieldName" value="TABLE_NAME" type="String"/><column name="SortFields(3)" property="SortOn" value="0" type="Double"/><column name="SortFields(3)" property="Order" value="1" type="Double"/><column name="SortFields(3)" property="DataOption" value="0" type="Double"/><column name="SortFields(4)" property="KeyfieldName" value="COLUMN_NAME" type="String"/><column name="SortFields(4)" property="SortOn" value="0" type="Double"/><column name="SortFields(4)" property="Order" value="1" type="Double"/><column name="SortFields(4)" property="DataOption" value="0" type="Double"/></view><view name="By Column"><column name="" property="ListObjectName" value="translations" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="COLUMN_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="LANGUAGE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_DESC" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_COMMENT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="SortFields(1)" property="KeyfieldName" value="LANGUAGE_NAME" type="String"/><column name="SortFields(1)" property="SortOn" value="0" type="Double"/><column name="SortFields(1)" property="Order" value="1" type="Double"/><column name="SortFields(1)" property="DataOption" value="0" type="Double"/><column name="SortFields(2)" property="KeyfieldName" value="COLUMN_NAME" type="String"/><column name="SortFields(2)" property="SortOn" value="0" type="Double"/><column name="SortFields(2)" property="Order" value="1" type="Double"/><column name="SortFields(2)" property="DataOption" value="0" type="Double"/><column name="SortFields(3)" property="KeyfieldName" value="TABLE_SCHEMA" type="String"/><column name="SortFields(3)" property="SortOn" value="0" type="Double"/><column name="SortFields(3)" property="Order" value="1" type="Double"/><column name="SortFields(3)" property="DataOption" value="0" type="Double"/><column name="SortFields(4)" property="KeyfieldName" value="TABLE_NAME" type="String"/><column name="SortFields(4)" property="SortOn" value="0" type="Double"/><column name="SortFields(4)" property="Order" value="1" type="Double"/><column name="SortFields(4)" property="DataOption" value="0" type="Double"/></view><view name="By Column + Language"><column name="" property="ListObjectName" value="translations" type="String"/><column name="" property="ShowTotals" value="False" type="Boolean"/><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean"/><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="COLUMN_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="LANGUAGE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_NAME" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_DESC" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="TRANSLATED_COMMENT" property="EntireColumn.Hidden" value="False" type="Boolean"/><column name="SortFields(1)" property="KeyfieldName" value="COLUMN_NAME" type="String"/><column name="SortFields(1)" property="SortOn" value="0" type="Double"/><column name="SortFields(1)" property="Order" value="1" type="Double"/><column name="SortFields(1)" property="DataOption" value="0" type="Double"/><column name="SortFields(2)" property="KeyfieldName" value="LANGUAGE_NAME" type="String"/><column name="SortFields(2)" property="SortOn" value="0" type="Double"/><column name="SortFields(2)" property="Order" value="1" type="Double"/><column name="SortFields(2)" property="DataOption" value="0" type="Double"/><column name="SortFields(3)" property="KeyfieldName" value="TABLE_SCHEMA" type="String"/><column name="SortFields(3)" property="SortOn" value="0" type="Double"/><column name="SortFields(3)" property="Order" value="1" type="Double"/><column name="SortFields(3)" property="DataOption" value="0" type="Double"/><column name="SortFields(4)" property="KeyfieldName" value="TABLE_NAME" type="String"/><column name="SortFields(4)" property="SortOn" value="0" type="Double"/><column name="SortFields(4)" property="Order" value="1" type="Double"/><column name="SortFields(4)" property="DataOption" value="0" type="Double"/></view></views></table>');

INSERT INTO logs.workbooks (NAME, TEMPLATE, DEFINITION, TABLE_SCHEMA) VALUES (N'change_tracking_admin.xlsx', NULL, N'tracked_objects=logs.view_captured_objects,logs.view_query_list,False,$B$3,,{"Parameters":{"schema":null,"has_log_triggers":null},"WorkbookLanguage":"en"}
attached_handlers=logs.base_tables,logs.view_query_list,False,$B$3,,{"Parameters":{"OBJECT_SCHEMA":null,"BASE_TABLE_SCHEMA":null},"WorkbookLanguage":"en"}
handlers=logs.handlers,(Default),False,$B$3,,{"Parameters":{"TABLE_SCHEMA":null,"TABLE_NAME":null,"EVENT_NAME":null},"WorkbookLanguage":"en"}
translations=logs.translations,(Default),False,$B$3,,{"Parameters":{"TABLE_SCHEMA":"logs","LANGUAGE_NAME":"en"},"WorkbookLanguage":"en"}
translation_pivot=logs.usp_translations,(Default),False,$B$3,,{"Parameters":{"field":"TRANSLATED_NAME"},"WorkbookLanguage":"en"}
workbooks=logs.workbooks,(Default),False,$B$3,,{"Parameters":{},"WorkbookLanguage":"en"}
', N'logs');

GO

EXEC logs.xl_actions_set_role_permissions;
GO

print 'Change Tracking Framework installed';
