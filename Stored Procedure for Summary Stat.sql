USE [SuperStoreDB]
GO

/****** Object:  StoredProcedure [dbo].[sp_SummaryStatistics]    Script Date: 21-Feb-2024 4:34:48 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- You could use this stored procedure on any table
-- to get summary statistics 
-- =============================================
CREATE PROCEDURE [dbo].[sp_SummaryStatistics] 
	-- Add the parameters for the stored procedure here
     @tableName NVARCHAR(MAX)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
	SET NOCOUNT ON;

DECLARE @sql NVARCHAR(MAX);
DECLARE @sqlQuery1 NVARCHAR(MAX);

-- Initialize @sqlQuery1 variable
SET @sqlQuery1 = '';

-- Get all numeric columns in the table and calculate statistics
SELECT @sqlQuery1 = @sqlQuery1 +
    'SELECT ''' + COLUMN_NAME + ''' AS ColumnName, ''N_Rows'' AS Statistic, COUNT(' + QUOTENAME(COLUMN_NAME) + ') AS Value FROM ' + QUOTENAME(@tableName) + ' WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT ''' + COLUMN_NAME + ''' AS ColumnName, ''Min'' AS Statistic, ROUND(MIN(' + QUOTENAME(COLUMN_NAME) + '),2) AS Value FROM ' + QUOTENAME(@tableName) + ' WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT ''' + COLUMN_NAME + ''' AS ColumnName, ''Max'' AS Statistic, ROUND(MAX(' + QUOTENAME(COLUMN_NAME) + '),2) AS Value FROM ' + QUOTENAME(@tableName) + ' WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT ''' + COLUMN_NAME + ''' AS ColumnName, ''Avg'' AS Statistic, ROUND(AVG(' + QUOTENAME(COLUMN_NAME) + '),2) AS Value FROM ' + QUOTENAME(@tableName) + ' WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT ''' + COLUMN_NAME + ''' AS ColumnName, ''StdDev'' AS Statistic, ROUND(STDEV(' + QUOTENAME(COLUMN_NAME) + '),2) AS Value FROM ' + QUOTENAME(@tableName) + ' WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT ''' + COLUMN_NAME + ''' AS ColumnName, ''Coff_Variance'' AS Statistic, ROUND((STDEV(' + QUOTENAME(COLUMN_NAME) + ')/AVG(' + QUOTENAME(COLUMN_NAME) + ')),2) AS Value FROM ' + QUOTENAME(@tableName) + ' WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT ''' + COLUMN_NAME + ''' AS ColumnName, ''NA_Count'' AS Statistic, ROUND(COUNT(*) - COUNT(' + QUOTENAME(COLUMN_NAME) + '),2) AS Value FROM ' + QUOTENAME(@tableName) + ' WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NULL UNION ALL ' +
    'SELECT ''' + COLUMN_NAME + ''' AS ColumnName, ''Range'' AS Statistic, ROUND(MAX(' + QUOTENAME(COLUMN_NAME) + ') - MIN(' + QUOTENAME(COLUMN_NAME) + '),2) AS Value FROM ' + QUOTENAME(@tableName) + ' WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT DISTINCT ''' + COLUMN_NAME + ''' AS ColumnName, ''1st_Quartile'' AS Statistic, ROUND(PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY ' + QUOTENAME(COLUMN_NAME) + ') OVER (),2) AS Value FROM (SELECT ' + QUOTENAME(COLUMN_NAME) + ' FROM ' + QUOTENAME(@tableName) + ') AS T WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT DISTINCT ''' + COLUMN_NAME + ''' AS ColumnName, ''Median'' AS Statistic, ROUND(PERCENTILE_CONT(0.50) WITHIN GROUP (ORDER BY ' + QUOTENAME(COLUMN_NAME) + ') OVER (),2) AS Value FROM (SELECT ' + QUOTENAME(COLUMN_NAME) + ' FROM ' + QUOTENAME(@tableName) + ') AS T WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' +
    'SELECT DISTINCT ''' + COLUMN_NAME + ''' AS ColumnName, ''3rd_Quartile'' AS Statistic, ROUND(PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY ' + QUOTENAME(COLUMN_NAME) + ') OVER (),2) AS Value FROM (SELECT ' + QUOTENAME(COLUMN_NAME) + ' FROM ' + QUOTENAME(@tableName) + ') AS T WHERE ' + QUOTENAME(COLUMN_NAME) + ' IS NOT NULL UNION ALL ' 
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = @tableName
  AND DATA_TYPE IN ('int', 'smallint', 'tinyint', 'bigint', 'decimal', 'numeric', 'float', 'real');


-- Remove the last "UNION ALL" from @sqlQuery1
SET @sqlQuery1 = LEFT(@sqlQuery1, LEN(@sqlQuery1) - LEN('UNION ALL '));

 
-- Construct the dynamic SQL query
SET @sql = '
SELECT ColumnName, Statistic, Value
FROM (' + @sqlQuery1 + ') AS SummaryStatistics;';


-- Declare a table variable to store the result of the dynamic SQL query
DECLARE @SummaryStatistics TABLE (
    ColumnName NVARCHAR(50),
    Statistic NVARCHAR(50),
    Value FLOAT
);




-- Execute the dynamic SQL query and insert the result into the table variable
INSERT INTO @SummaryStatistics (ColumnName, Statistic, Value)
EXEC sp_executesql @sqlQuery1;

DECLARE @resultTable TABLE(
    ColumnName NVARCHAR(50),
	N_Rows FLOAT,
    Min FLOAT,
    Max FLOAT,
    Avg FLOAT,
    StdDev FLOAT,
    Coff_Variance FLOAT,
    NA_Count INT,
    Range FLOAT,
    Quartile_1 FLOAT,
    Median FLOAT,
    Quartile_3 FLOAT,
	Lower_CI FLOAT,
	Upper_CI FLOAT,
	Skewness FLOAT,
	Skewness_Comment NVARCHAR(255)
);

-- Inserting into resultTable in columnnar form
INSERT INTO @resultTable (ColumnName,N_Rows, Min, Max, Avg, StdDev, Coff_Variance, NA_Count, Range, Quartile_1, Median, Quartile_3)
SELECT
    ColumnName,
	MAX(CASE WHEN Statistic = 'N_Rows' THEN Value END) AS N_Rows,
    MAX(CASE WHEN Statistic = 'Min' THEN Value END) AS Min,
    MAX(CASE WHEN Statistic = 'Max' THEN Value END) AS Max,
    MAX(CASE WHEN Statistic = 'Avg' THEN Value END) AS Avg,
    MAX(CASE WHEN Statistic = 'StdDev' THEN Value END) AS StdDev,
    MAX(CASE WHEN Statistic = 'Coff_Variance' THEN Value END) AS Coff_Variance,
    MAX(CASE WHEN Statistic = 'NA_Count' THEN Value END) AS NA_Count,
    MAX(CASE WHEN Statistic = 'Range' THEN Value END) AS Range,
    MAX(CASE WHEN Statistic = '1st_Quartile' THEN Value END) AS Quartile_1,
    MAX(CASE WHEN Statistic = 'Median' THEN Value END) AS Median,
    MAX(CASE WHEN Statistic = '3rd_Quartile' THEN Value END) AS Quartile_3
FROM @SummaryStatistics
GROUP BY ColumnName;


DECLARE @sqlQuery2 NVARCHAR(MAX) = '';
SELECT @sqlQuery2 = @sqlQuery2 + '
    SELECT 
        ''' + COLUMN_NAME + ''' AS ColumnName,
        AVG(POWER(CAST(ColumnVal AS FLOAT) - AVG_VAL, 3))
        / POWER(STDEV(CAST(ColumnVal AS FLOAT)), 3) AS Skewness
    FROM (
        SELECT 
            CAST([' + COLUMN_NAME + '] AS FLOAT) AS ColumnVal,
            AVG(CAST([' + COLUMN_NAME + '] AS FLOAT)) OVER() AS AVG_VAL
        FROM ' + QUOTENAME(@tableName) + '
    ) AS subquery
    GROUP BY AVG_VAL
    UNION ALL
' 
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = @tableName 
 AND DATA_TYPE IN ('int', 'smallint', 'tinyint', 'bigint', 'decimal', 'numeric', 'float', 'real');
 
SET @sqlQuery2 = LEFT(@sqlQuery2, LEN(@sqlQuery2) - 11); -- Remove the last UNION ALL
  
DECLARE @skewnessTbl TABLE(
    ColumnName NVARCHAR(50),
	Skewness FLOAT
);

 INSERT INTO @skewnessTbl (ColumnName,Skewness)
EXEC sp_executesql @sqlQuery2;

MERGE INTO @resultTable AS target
USING @skewnessTbl AS source
ON target.ColumnName = source.ColumnName
WHEN MATCHED THEN
    UPDATE SET target.Skewness = ROUND(source.Skewness,3)
WHEN NOT MATCHED BY TARGET THEN
    INSERT (ColumnName, Skewness)
    VALUES (source.ColumnName, source.Skewness);

--- Reference: https://www.geeksforgeeks.org/skewness-measures-and-interpretation/
-- Lower and Upper Confidential Interval, and skewness comments
UPDATE @resultTable SET 
Lower_CI = ROUND(Avg - (1.959964 * StdDev / SQRT(N_Rows)),2),
Upper_CI = ROUND(Avg + (1.959964 * StdDev / SQRT(N_Rows)),2),
Skewness_Comment = CASE 
		WHEN Skewness BETWEEN -0.5 AND 0.5 THEN 'Weak or No Skewness'
		WHEN Skewness BETWEEN -1 AND -0.5 THEN 'Moderately left Skewed'
		WHEN Skewness BETWEEN 0.5 AND 1 THEN 'Moderately right Skewed'
		WHEN Skewness <= -1 THEN 'Strongly left Skewed'
		WHEN Skewness >= 1 THEN 'Strongly right Skewed'
        ELSE ''
    END

-- Use the @resultTable temporary table as needed
SELECT * FROM @resultTable;

END
GO

