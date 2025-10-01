
USE ExcelLoader;
GO

-- Sample FileGroup / Entity / Schema
DECLARE @COL int, @SCHEMA int, @TBL int, @LoanId int, @AsOf int, @UPB int, @FS int;

IF NOT EXISTS (SELECT 1 FROM repo.FileGroup WHERE Code='COL')
INSERT repo.FileGroup(Name,Code) VALUES ('Collateral','COL');

IF NOT EXISTS (SELECT 1 FROM repo.Entity WHERE Code='LenderA')
INSERT repo.Entity(Name,Code) VALUES ('LenderA','LenderA');

SELECT @COL = FileGroupId FROM repo.FileGroup WHERE Code='COL';

INSERT repo.TargetSchema(Name,Version,IsActive,Notes) VALUES ('Collateral.v1',1,1,'Sample collateral schema');
SELECT @SCHEMA = SCOPE_IDENTITY();

INSERT repo.TargetTable(TargetSchemaId,Name,KeyStrategy,StageFirst,UniqueKeyCsv)
VALUES (@SCHEMA,'dbo.Loans','UPSERT',1,'LoanId,AsOfDate');
SELECT @TBL = SCOPE_IDENTITY();

INSERT repo.TargetField(TargetTableId,Name,DataType,IsNullable,Ordinal,RequiredForKey) VALUES
(@TBL,'LoanId','NVARCHAR(50)',0,1,1),
(@TBL,'AsOfDate','DATE',0,2,1),
(@TBL,'UPB','DECIMAL(18,2)',1,3,0);

INSERT repo.FileSpec(FileGroupId,Name,Pattern,Parser,SheetHint,HeaderRow,FirstDataRow,TargetSchemaId,Active)
VALUES (@COL,'Collateral Daily','\\share\\collateral\\*Daily*.xlsx','excel',NULL,1,2,@SCHEMA,1);
SELECT @FS = SCOPE_IDENTITY();

-- Base mappings
INSERT repo.FieldMap(FileSpecId,SourceHeader,TargetTableId,TargetFieldId,TransformChain,[Required])
SELECT @FS,'Loan ID',@TBL, (SELECT TargetFieldId FROM repo.TargetField WHERE TargetTableId=@TBL AND Name='LoanId'), 'trim',1
UNION ALL
SELECT @FS,'As Of',  @TBL, (SELECT TargetFieldId FROM repo.TargetField WHERE TargetTableId=@TBL AND Name='AsOfDate'), 'date(M/d/yyyy)',1
UNION ALL
SELECT @FS,'UPB',    @TBL, (SELECT TargetFieldId FROM repo.TargetField WHERE TargetTableId=@TBL AND Name='UPB'), 'money',0;

-- Override for LenderA: "Current UPB" instead of "UPB"
INSERT repo.FieldMapOverride(FieldMapId,EntityId,LoadTypeId,SourceHeader)
SELECT fm.FieldMapId, e.EntityId, NULL, 'Current UPB'
FROM repo.FieldMap fm
CROSS JOIN repo.Entity e
WHERE fm.FileSpecId=@FS AND fm.SourceHeader='UPB' AND e.Code='LenderA';

PRINT 'Seed complete.';
