/*
Cleanup duplicate customer columns in supported customer tables.

What this script does:
1) Ensures canonical customer columns exist.
2) Copies values from duplicate/legacy columns into canonical columns when canonical is empty.
3) Drops duplicate columns.
*/

SET NOCOUNT ON;

DECLARE @targets TABLE (full_name NVARCHAR(300));
INSERT INTO @targets (full_name)
VALUES
    (N'dbo.CustomerMasterData'),
    (N'dbo.CustomerFormUnified'),
    (N'dbo.Customer Master Data');

DECLARE @pairs TABLE (canonical SYSNAME, duplicate SYSNAME);
INSERT INTO @pairs (canonical, duplicate)
VALUES
    (N'approverStatus', N'ApproverStatus'),
    (N'approverStatus', N'ApprovStatus'),
    (N'approverStatus', N'status'),
    (N'approverStatus', N'Status'),
    (N'CustomerName', N'customerName'),
    (N'CustomerName', N'vendorName'),
    (N'CustomerName', N'VendorName'),
    (N'CustomerName', N'Vendorname'),
    (N'CustomerGroup', N'customerGroup'),
    (N'CustomerGroup', N'vendorCategory'),
    (N'CustomerGroup', N'VendorCategory'),
    (N'CustomerGroup', N'Vendor_Category'),
    (N'CustomerType', N'vendorType'),
    (N'CustomerType', N'VendorType'),
    (N'customerEmail', N'CustomerEmail'),
    (N'customerEmail', N'vendorEmail'),
    (N'customerEmail', N'VendorEmail');

DECLARE @drop_only TABLE (column_name SYSNAME);
INSERT INTO @drop_only (column_name)
VALUES
    (N'validateBy'),
    (N'Validate By'),
    (N'msmeValidate'),
    (N'MSME Validate'),
    (N'pfValidate'),
    (N'PF Validate'),
    (N'esiValidate'),
    (N'ESI Validate'),
    (N'bankValidate'),
    (N'Bank Validate'),
    (N'companyName'),
    (N'registrationPath'),
    (N'registrationLink'),
    (N'sqlRecordId'),
    (N'reviewDecision'),
    (N'ReviewDecision'),
    (N'rejectTo'),
    (N'Reject To'),
    (N'invitedById'),
    (N'invitedByEmail'),
    (N'InvitedByEmail'),
    (N'Invited By Email'),
    (N'contactPersonEmail'),
    (N'ContactPersonEmail'),
    (N'ContactEmailID'),
    (N'contactEmailId'),
    (N'rubaminApproverHod'),
    (N'RubaminApproverHod'),
    (N'RubaminApprovalHod'),
    (N'rubaminContactPerson'),
    (N'rubaminContactPersonEmail'),
    (N'rubaminContactPersonDepartment'),
    (N'Rubamin Contact Person'),
    (N'Rubamin Contact Person Email'),
    (N'Rubamin Contact Person Department');

DECLARE @full_name NVARCHAR(300);
DECLARE target_cursor CURSOR FAST_FORWARD FOR
    SELECT full_name
    FROM @targets;

OPEN target_cursor;
FETCH NEXT FROM target_cursor INTO @full_name;

WHILE @@FETCH_STATUS = 0
BEGIN
    IF OBJECT_ID(@full_name, N'U') IS NOT NULL
    BEGIN
        DECLARE @schema SYSNAME = ISNULL(PARSENAME(@full_name, 2), N'dbo');
        DECLARE @table SYSNAME = PARSENAME(@full_name, 1);
        DECLARE @lookup_name NVARCHAR(300) = @schema + N'.' + @table;
        DECLARE @qualified_name NVARCHAR(400) = QUOTENAME(@schema) + N'.' + QUOTENAME(@table);
        DECLARE @sql NVARCHAR(MAX);

        IF COL_LENGTH(@lookup_name, 'approverStatus') IS NULL
        BEGIN
            SET @sql = N'ALTER TABLE ' + @qualified_name + N' ADD [approverStatus] NVARCHAR(200) NULL;';
            EXEC sp_executesql @sql;
        END

        IF COL_LENGTH(@lookup_name, 'CustomerName') IS NULL
        BEGIN
            SET @sql = N'ALTER TABLE ' + @qualified_name + N' ADD [CustomerName] NVARCHAR(500) NULL;';
            EXEC sp_executesql @sql;
        END

        IF COL_LENGTH(@lookup_name, 'CustomerGroup') IS NULL
        BEGIN
            SET @sql = N'ALTER TABLE ' + @qualified_name + N' ADD [CustomerGroup] NVARCHAR(500) NULL;';
            EXEC sp_executesql @sql;
        END

        IF COL_LENGTH(@lookup_name, 'CustomerType') IS NULL
        BEGIN
            SET @sql = N'ALTER TABLE ' + @qualified_name + N' ADD [CustomerType] NVARCHAR(500) NULL;';
            EXEC sp_executesql @sql;
        END

        IF COL_LENGTH(@lookup_name, 'customerEmail') IS NULL
        BEGIN
            SET @sql = N'ALTER TABLE ' + @qualified_name + N' ADD [customerEmail] NVARCHAR(320) NULL;';
            EXEC sp_executesql @sql;
        END

        DECLARE @canonical SYSNAME;
        DECLARE @duplicate SYSNAME;
        DECLARE pair_cursor CURSOR FAST_FORWARD FOR
            SELECT canonical, duplicate
            FROM @pairs;

        OPEN pair_cursor;
        FETCH NEXT FROM pair_cursor INTO @canonical, @duplicate;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            IF LOWER(@canonical) <> LOWER(@duplicate)
               AND COL_LENGTH(@lookup_name, @canonical) IS NOT NULL
               AND COL_LENGTH(@lookup_name, @duplicate) IS NOT NULL
            BEGIN
                SET @sql = N'
                    UPDATE ' + @qualified_name + N'
                    SET [' + @canonical + N'] = [' + @duplicate + N']
                    WHERE [' + @duplicate + N'] IS NOT NULL
                      AND (
                        [' + @canonical + N'] IS NULL
                        OR LTRIM(RTRIM(CAST([' + @canonical + N'] AS NVARCHAR(MAX)))) = ''''
                      );
                    ALTER TABLE ' + @qualified_name + N' DROP COLUMN [' + @duplicate + N'];';
                EXEC sp_executesql @sql;
            END

            FETCH NEXT FROM pair_cursor INTO @canonical, @duplicate;
        END

        CLOSE pair_cursor;
        DEALLOCATE pair_cursor;

        DECLARE @drop_col SYSNAME;
        DECLARE drop_cursor CURSOR FAST_FORWARD FOR
            SELECT column_name
            FROM @drop_only;

        OPEN drop_cursor;
        FETCH NEXT FROM drop_cursor INTO @drop_col;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            IF COL_LENGTH(@lookup_name, @drop_col) IS NOT NULL
            BEGIN
                SET @sql = N'ALTER TABLE ' + @qualified_name + N' DROP COLUMN [' + @drop_col + N'];';
                EXEC sp_executesql @sql;
            END

            FETCH NEXT FROM drop_cursor INTO @drop_col;
        END

        CLOSE drop_cursor;
        DEALLOCATE drop_cursor;
    END

    FETCH NEXT FROM target_cursor INTO @full_name;
END

CLOSE target_cursor;
DEALLOCATE target_cursor;
