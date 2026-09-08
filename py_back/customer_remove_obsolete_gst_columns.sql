/* Remove obsolete GST 2 / GST 3 columns from customer table(s). */

IF OBJECT_ID(N'dbo.CustomerMasterData', N'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTIN1') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTIN1];
    IF COL_LENGTH('dbo.CustomerMasterData', 'TradeNameGST1') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [TradeNameGST1];
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTCertificate1') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTCertificate1];
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTCertificate1_FileName') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTCertificate1_FileName];
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTCertificate1_ContentType') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTCertificate1_ContentType];
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTCertificate1_FileSize') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTCertificate1_FileSize];

    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTIN2') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTIN2];
    IF COL_LENGTH('dbo.CustomerMasterData', 'TradeNameGST2') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [TradeNameGST2];
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTCertificate2') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTCertificate2];
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTCertificate2_FileName') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTCertificate2_FileName];
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTCertificate2_ContentType') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTCertificate2_ContentType];
    IF COL_LENGTH('dbo.CustomerMasterData', 'GSTCertificate2_FileSize') IS NOT NULL ALTER TABLE dbo.CustomerMasterData DROP COLUMN [GSTCertificate2_FileSize];
END;

IF OBJECT_ID(N'dbo.CustomerFormUnified', N'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTIN1') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTIN1];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'TradeNameGST1') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [TradeNameGST1];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTCertificate1') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTCertificate1];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTCertificate1_FileName') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTCertificate1_FileName];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTCertificate1_ContentType') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTCertificate1_ContentType];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTCertificate1_FileSize') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTCertificate1_FileSize];

    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTIN2') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTIN2];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'TradeNameGST2') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [TradeNameGST2];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTCertificate2') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTCertificate2];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTCertificate2_FileName') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTCertificate2_FileName];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTCertificate2_ContentType') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTCertificate2_ContentType];
    IF COL_LENGTH('dbo.CustomerFormUnified', 'GSTCertificate2_FileSize') IS NOT NULL ALTER TABLE dbo.CustomerFormUnified DROP COLUMN [GSTCertificate2_FileSize];
END;
