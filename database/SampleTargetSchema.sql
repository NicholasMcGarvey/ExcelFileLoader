
USE ExcelLoader;
GO

IF OBJECT_ID('dbo.Loans','U') IS NULL
CREATE TABLE dbo.Loans (
  LoanId       nvarchar(50) NOT NULL,
  AsOfDate     date NOT NULL,
  UPB          decimal(18,2) NULL,
  CreatedUtc   datetime2 NOT NULL DEFAULT SYSUTCDATETIME(),
  ModifiedUtc  datetime2 NULL,
  CONSTRAINT PK_Loans PRIMARY KEY (LoanId, AsOfDate)
);
