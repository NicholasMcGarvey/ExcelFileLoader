
-- Repository schema for ExcelLoader (SQL Server)
IF DB_ID('ExcelLoader') IS NULL
BEGIN
  PRINT 'Creating database ExcelLoader...'
  EXEC('CREATE DATABASE ExcelLoader');
END
GO
USE ExcelLoader;
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;

IF SCHEMA_ID('repo') IS NULL EXEC('CREATE SCHEMA repo');
GO

-- Reference tables
IF OBJECT_ID('repo.LoadType','U') IS NULL
CREATE TABLE repo.LoadType (
  LoadTypeId   int IDENTITY PRIMARY KEY,
  Name         varchar(40) UNIQUE NOT NULL
);

IF OBJECT_ID('repo.FileGroup','U') IS NULL
CREATE TABLE repo.FileGroup (
  FileGroupId  int IDENTITY PRIMARY KEY,
  Name         varchar(80) NOT NULL,
  Code         varchar(40) NOT NULL,
  CONSTRAINT UQ_FileGroup_Code UNIQUE (Code)
);

IF OBJECT_ID('repo.Entity','U') IS NULL
CREATE TABLE repo.Entity (
  EntityId     int IDENTITY PRIMARY KEY,
  Name         varchar(120) NOT NULL,
  Code         varchar(80) NOT NULL,
  CONSTRAINT UQ_Entity_Code UNIQUE (Code)
);

IF OBJECT_ID('repo.TargetSchema','U') IS NULL
CREATE TABLE repo.TargetSchema (
  TargetSchemaId int IDENTITY PRIMARY KEY,
  Name           varchar(120) NOT NULL,
  Version        int NOT NULL DEFAULT 1,
  IsActive       bit NOT NULL DEFAULT 1,
  Notes          varchar(4000) NULL,
  CONSTRAINT UQ_TargetSchema UNIQUE (Name, Version)
);

IF OBJECT_ID('repo.TargetTable','U') IS NULL
CREATE TABLE repo.TargetTable (
  TargetTableId  int IDENTITY PRIMARY KEY,
  TargetSchemaId int NOT NULL FOREIGN KEY REFERENCES repo.TargetSchema(TargetSchemaId),
  Name           sysname NOT NULL,
  KeyStrategy    varchar(30) NOT NULL DEFAULT 'UPSERT',
  StageFirst     bit NOT NULL DEFAULT 1,
  UniqueKeyCsv   varchar(4000) NULL
);

IF OBJECT_ID('repo.TargetField','U') IS NULL
CREATE TABLE repo.TargetField (
  TargetFieldId  int IDENTITY PRIMARY KEY,
  TargetTableId  int NOT NULL FOREIGN KEY REFERENCES repo.TargetTable(TargetTableId),
  Name           sysname NOT NULL,
  DataType       varchar(60) NOT NULL,
  IsNullable     bit NOT NULL DEFAULT 1,
  Ordinal        int NOT NULL DEFAULT 0,
  RequiredForKey bit NOT NULL DEFAULT 0
);

IF OBJECT_ID('repo.FileSpec','U') IS NULL
CREATE TABLE repo.FileSpec (
  FileSpecId     int IDENTITY PRIMARY KEY,
  FileGroupId    int NOT NULL FOREIGN KEY REFERENCES repo.FileGroup(FileGroupId),
  Name           varchar(120) NOT NULL,
  Pattern        varchar(400) NOT NULL,
  Parser         varchar(40) NOT NULL DEFAULT 'excel', -- excel,csv,tsv
  SheetHint      varchar(120) NULL,
  HeaderRow      int NOT NULL DEFAULT 1,
  FirstDataRow   int NOT NULL DEFAULT 2,
  TargetSchemaId int NOT NULL FOREIGN KEY REFERENCES repo.TargetSchema(TargetSchemaId),
  Active         bit NOT NULL DEFAULT 1
);

IF OBJECT_ID('repo.FieldMap','U') IS NULL
CREATE TABLE repo.FieldMap (
  FieldMapId     int IDENTITY PRIMARY KEY,
  FileSpecId     int NOT NULL FOREIGN KEY REFERENCES repo.FileSpec(FileSpecId),
  SourceHeader   varchar(256) NOT NULL,
  TargetTableId  int NOT NULL FOREIGN KEY REFERENCES repo.TargetTable(TargetTableId),
  TargetFieldId  int NOT NULL FOREIGN KEY REFERENCES repo.TargetField(TargetFieldId),
  TransformChain varchar(1000) NULL,
  DefaultValue   varchar(400) NULL,
  [Required]     bit NOT NULL DEFAULT 0
);

IF OBJECT_ID('repo.FieldMapOverride','U') IS NULL
CREATE TABLE repo.FieldMapOverride (
  FieldMapOverrideId int IDENTITY PRIMARY KEY,
  FieldMapId         int NOT NULL FOREIGN KEY REFERENCES repo.FieldMap(FieldMapId),
  EntityId           int NULL FOREIGN KEY REFERENCES repo.Entity(EntityId),
  LoadTypeId         int NULL FOREIGN KEY REFERENCES repo.LoadType(LoadTypeId),
  SourceHeader       varchar(256) NULL,
  TargetTableId      int NULL,
  TargetFieldId      int NULL,
  TransformChain     varchar(1000) NULL,
  DefaultValue       varchar(400) NULL,
  [Required]         bit NULL,
  CONSTRAINT UQ_FieldMapOverride UNIQUE (FieldMapId, EntityId, LoadTypeId)
);

IF OBJECT_ID('repo.ValidationRule','U') IS NULL
CREATE TABLE repo.ValidationRule (
  ValidationRuleId int IDENTITY PRIMARY KEY,
  Scope            varchar(20) NOT NULL,   -- ROW,FILE,DB
  Phase            varchar(20) NOT NULL,   -- PRE,POST
  FileSpecId       int NULL,
  TargetTableId    int NULL,
  RuleKind         varchar(50) NOT NULL,
  RuleConfigJson   nvarchar(max) NOT NULL,
  Severity         varchar(10) NOT NULL DEFAULT 'ERROR'
);

IF OBJECT_ID('repo.LoadRun','U') IS NULL
CREATE TABLE repo.LoadRun (
  LoadRunId     bigint IDENTITY PRIMARY KEY,
  StartedUtc    datetime2 NOT NULL DEFAULT SYSUTCDATETIME(),
  CompletedUtc  datetime2 NULL,
  [Status]      varchar(20) NOT NULL DEFAULT 'RUNNING',
  FileSpecId    int NOT NULL FOREIGN KEY REFERENCES repo.FileSpec(FileSpecId),
  EntityId      int NULL FOREIGN KEY REFERENCES repo.Entity(EntityId),
  LoadTypeId    int NULL FOREIGN KEY REFERENCES repo.LoadType(LoadTypeId),
  SourcePath    nvarchar(1024) NOT NULL,
  DetectedSheet varchar(120) NULL,
  RowCountRead  int NULL,
  RowCountLoaded int NULL,
  ErrorCount    int NULL,
  [Message]     nvarchar(max) NULL
);

IF OBJECT_ID('repo.LoadRunError','U') IS NULL
CREATE TABLE repo.LoadRunError (
  LoadRunErrorId bigint IDENTITY PRIMARY KEY,
  LoadRunId      bigint NOT NULL FOREIGN KEY REFERENCES repo.LoadRun(LoadRunId),
  RowNumber      int NULL,
  ColumnName     varchar(128) NULL,
  ErrorCode      varchar(60) NOT NULL,
  ErrorMessage   nvarchar(2000) NOT NULL,
  RawValue       nvarchar(4000) NULL
);

-- Seed
IF NOT EXISTS (SELECT 1 FROM repo.LoadType)
INSERT repo.LoadType(Name) VALUES ('static'),('cyclical'),('daily'),('supplemental');

-- Utility: resolve effective mappings given entity/loadtype
IF OBJECT_ID('repo.fn_GetEffectiveFieldMap','IF') IS NOT NULL
  DROP FUNCTION repo.fn_GetEffectiveFieldMap;
GO
CREATE FUNCTION repo.fn_GetEffectiveFieldMap
(
  @FileSpecId int,
  @EntityId   int = NULL,
  @LoadTypeId int = NULL
)
RETURNS TABLE
AS
RETURN
(
  WITH Ranked AS (
    SELECT
      fm.FieldMapId, fm.SourceHeader, fm.TargetTableId, fm.TargetFieldId,
      fm.TransformChain, fm.DefaultValue, fm.[Required],
      o.SourceHeader AS oSourceHeader, o.TargetTableId AS oTargetTableId,
      o.TargetFieldId AS oTargetFieldId, o.TransformChain AS oTransformChain,
      o.DefaultValue AS oDefaultValue, o.[Required] AS oRequired,
      ROW_NUMBER() OVER (
        PARTITION BY fm.FieldMapId
        ORDER BY
          CASE WHEN o.EntityId IS NOT NULL AND o.LoadTypeId IS NOT NULL THEN 1
               WHEN o.EntityId IS NOT NULL AND o.LoadTypeId IS NULL THEN 2
               WHEN o.EntityId IS NULL AND o.LoadTypeId IS NOT NULL THEN 3
               ELSE 4 END
      ) AS rn
    FROM repo.FieldMap fm
    OUTER APPLY (
      SELECT TOP 1 *
      FROM repo.FieldMapOverride o
      WHERE o.FieldMapId = fm.FieldMapId
        AND (@EntityId IS NULL OR o.EntityId = @EntityId OR o.EntityId IS NULL)
        AND (@LoadTypeId IS NULL OR o.LoadTypeId = @LoadTypeId OR o.LoadTypeId IS NULL)
      ORDER BY
        CASE WHEN o.EntityId IS NOT NULL AND o.LoadTypeId IS NOT NULL THEN 1
             WHEN o.EntityId IS NOT NULL AND o.LoadTypeId IS NULL THEN 2
             WHEN o.EntityId IS NULL AND o.LoadTypeId IS NOT NULL THEN 3
             ELSE 4 END
    ) o
    WHERE fm.FileSpecId = @FileSpecId
  )
  SELECT
    FieldMapId,
    COALESCE(oSourceHeader, SourceHeader) AS SourceHeader,
    COALESCE(oTargetTableId, TargetTableId) AS TargetTableId,
    COALESCE(oTargetFieldId, TargetFieldId) AS TargetFieldId,
    COALESCE(oTransformChain, TransformChain) AS TransformChain,
    COALESCE(oDefaultValue, DefaultValue) AS DefaultValue,
    COALESCE(oRequired, [Required]) AS [Required]
  FROM Ranked
  WHERE rn = 1
);
GO
