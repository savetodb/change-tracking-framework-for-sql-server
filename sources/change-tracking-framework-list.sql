-- =============================================
-- Change Tracking Framework for Microsoft SQL Server
-- Version 4.2, January 9, 2023
--
-- Copyright 2017-2023 Gartle LLC
--
-- License: MIT
-- =============================================

SET NOCOUNT ON

SELECT
    CAST(s.name AS nchar(15)) AS [SCHEMA]
    , CAST(o.name AS nchar(50)) AS [NAME]
    , CASE o.[type]
        WHEN 'P'  THEN 'procedure'
        WHEN 'IF' THEN 'function'
        WHEN 'FN' THEN 'function'
        WHEN 'TF' THEN 'function'
        WHEN 'V'  THEN 'view'
        WHEN 'U'  THEN 'table'
        ELSE o.[type_desc] END AS [TYPE]
FROM
    sys.objects o
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
    o.[type] IN ('U', 'V', 'P', 'IF', 'FN', 'TF')
    AND s.name IN ('logs')
    AND o.is_ms_shipped = 0
    AND NOT (s.name = 'dbo' AND (o.name LIKE 'sp_%' OR o.name LIKE 'fn_%' OR o.name LIKE 'sys%'))
ORDER BY
    CASE o.[type]
        WHEN 'P'  THEN 3
        WHEN 'IF' THEN 4
        WHEN 'FN' THEN 5
        WHEN 'TF' THEN 6
        WHEN 'V'  THEN 2
        WHEN 'U'  THEN 1
        ELSE 7 END
    , s.name
    , o.name
