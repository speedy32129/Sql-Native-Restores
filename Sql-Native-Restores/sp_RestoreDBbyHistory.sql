USE [msdb];
GO

IF EXISTS
(
    SELECT *
    FROM sys.objects
    WHERE object_id = OBJECT_ID(N'[dbo].[sp_RestoreDBbyHistory]')
          AND type IN ( N'P', N'PC' )
)
    DROP PROCEDURE [dbo].[sp_RestoreDBbyHistory];
GO

USE [msdb];
GO

CREATE PROCEDURE [dbo].[sp_RestoreDBbyHistory]
(
    @SourceDB VARCHAR(200)
  , @DestinationDB VARCHAR(200) = NULL
  , @RestoreToTime DATETIME = NULL            --'yyyy-mm-dd hh:mi:ss', ie. '2012-04-27 22:19:20'
  , @RecoveryMode VARCHAR(10) = 'Recovery'    --'Recovery' or 'Norecovery'
  , @ListMode VARCHAR(10) = 'OnlyValid'       --'All' or 'OnlyValid'
  , @ProgressValue INT = 1
  , @MirroringRestore BIT = 0                 -- Added by Jim for mirroring restores
  , @SQLLocationOneDrive varchar(250) = NULL  -- Used for resoring to different server with a different path
  , @SQLLocationData varchar(250) = NULL      -- Used for resoring to different server with a different path
  , @SQLLocationLogs varchar(250) = NULL      -- Used for resoring to different server with a different path
  , @SQLLocationOther varchar(250) = NULL     -- Used for resoring to different server with a different path
)
AS
/*********************************************************************************************
Original Author: xu james
blog: http://jamessql.blogspot.com/

Generate database restore script by backup history, it can 
1.  give out the restore script and sequence
        @ListMode='All' :   lists all possible restore scripts before @RestoreToTime
        @ListMode='OnlyValid' :  gives out a best valid restore way with mininum script.
2.  check if the backup file existed

Example:
1.  Generate restore script for database RS
        EXEC sp_RestoreDBbyHistory @SourceDB='RS'
2.  Generate all possible restore script for database RS, restore it as name 'RS_restored' to '2012-04-30 14:13:30.000' and with norecovery
        EXEC [sp_RestoreDBbyHistory] @SourceDB = 'RS'
                           , @DestinationDB = 'RS_restored'
                           , @RestoreToTime = '2012-04-30 14:13:30.000'
                           , @RecoveryMode = 'Norecovery'
                           , @ListMode = 'All';

Please test it by yourself first.
*********************************************************************************************/
BEGIN

    DECLARE @Exists INT
          , @File_Exists VARCHAR(30)
          , @backup_start_date DATETIME
          , @backup_finish_date DATETIME
          , @first_lsn NUMERIC(25, 0)
          , @last_lsn NUMERIC(25, 0)
          , @position INT
          , @backup_type VARCHAR(20)
          , @backup_size NUMERIC(20, 0)
          , @physical_device_name VARCHAR(500)
          , @backupset_name VARCHAR(500)
          , @differential_base_lsn NUMERIC(25, 0)
          , @database_backup_lsn NUMERIC(25, 0);


    DECLARE @restore_command VARCHAR(MAX)
          , @lastfile BIT
          , @stopat VARCHAR(50)
          , @MOVETO VARCHAR(MAX)
          , @MOVETO_temp VARCHAR(MAX)
          , @first_backup_set_id INT;

    IF (@SourceDB IS NULL)
    BEGIN
        PRINT 'Please input the @SourceDB name!';
        RETURN;
    END;

    IF (@DestinationDB IS NULL)
        SET @DestinationDB = @SourceDB;

    IF (@RestoreToTime IS NULL)
    BEGIN
        SET @stopat = '';
        SET @RestoreToTime = GETDATE();
    END;
    ELSE
        SET @stopat = ', STOPAT = ''' + CONVERT(VARCHAR(50), @RestoreToTime) + '''';

    IF (@RecoveryMode NOT IN ( 'Recovery', 'Norecovery' ))
    BEGIN
        PRINT 'Please set parameter @RecoveryMode with value ''Recovery'' or ''Norecovery''';
        RETURN;
    END;

    IF (@ListMode NOT IN ( 'All', 'OnlyValid' ))
    BEGIN
        PRINT 'Please set parameter @@Mode with value ''All'' or ''OnlyValid''';
        RETURN;
    END;

    SET @lastfile = 0;
    SET @MOVETO = '';

    --Find the last valid full backup
    SELECT TOP 1
           @first_backup_set_id = backup_set_id
    FROM msdb.dbo.backupset bs WITH ( NOLOCK )
    WHERE bs.type = 'D'
          AND bs.backup_start_date <= @RestoreToTime
          AND bs.database_name = @SourceDB
          AND bs.is_copy_only = 0                          -- Add Jim to ignore copy only backups as they are non-recoverable
    ORDER BY bs.backup_start_date DESC;

    IF (@first_backup_set_id IS NULL)
    BEGIN
        SELECT @first_backup_set_id = MIN(bs.backup_set_id)
        FROM msdb.dbo.backupset bs WITH ( NOLOCK )
        WHERE bs.database_name = @SourceDB
              AND bs.backup_start_date <= @RestoreToTime
              AND bs.is_copy_only = 0;                     -- Add Jim to ignore copy only backups as they are non-recoverable

        IF (@first_backup_set_id IS NULL)
        BEGIN
            PRINT 'There is no any valid backup!!!';
            RETURN;
        END;
        ELSE
            PRINT 'There is no valid full backup!!!';

    END;
    ELSE
    BEGIN
        SELECT @MOVETO
            = @MOVETO + 'MOVE N''' + bf.logical_name + ''' TO N'''
              -- Set Path
              + CASE
                    WHEN bf.file_number = 1 THEN
                        -- Data
                        COALESCE(
                                    @SQLLocationOneDrive
                                  , @SQLLocationData
                                  , REVERSE(RIGHT(REVERSE(bf.physical_name), (LEN(bf.physical_name) - CHARINDEX('\', REVERSE(bf.physical_name), 1)) + 1))
                                )
                    WHEN (bf.file_number <> 1)
                         AND (bf.file_type = 'L') THEN
                        -- Log
                        COALESCE(
                                    @SQLLocationOneDrive
                                  , @SQLLocationLogs
                                  , REVERSE(RIGHT(REVERSE(bf.physical_name), (LEN(bf.physical_name) - CHARINDEX('\', REVERSE(bf.physical_name), 1)) + 1))
                                )
                    WHEN (bf.file_number <> 1)
                         AND (bf.file_type = 'D') THEN
                        -- Known NDF
                        COALESCE(
                                    @SQLLocationOneDrive
                                  , @SQLLocationOther
                                  , REVERSE(RIGHT(REVERSE(bf.physical_name), (LEN(bf.physical_name) - CHARINDEX('\', REVERSE(bf.physical_name), 1)) + 1))
                                )
                    ELSE
                        -- Other NDF
                        COALESCE(
                                    @SQLLocationOneDrive
                                  , @SQLLocationOther
                                  , REVERSE(RIGHT(REVERSE(bf.physical_name), (LEN(bf.physical_name) - CHARINDEX('\', REVERSE(bf.physical_name), 1)) + 1))
                                )
                END
              -- Set FileName
              + CASE
                    WHEN bf.file_number = 1 THEN
                        @DestinationDB + '_data.mdf'
                    WHEN (bf.file_number <> 1)
                         AND (bf.file_type = 'L') THEN
                        --@DestinationDB + '_log' + CONVERT(VARCHAR(3), bf.file_number) + '.ldf'
                        @DestinationDB + '_log.ldf'
                    WHEN (bf.file_number <> 1)
                         AND (bf.file_type = 'D') THEN
                        @DestinationDB + '_data' + CONVERT(VARCHAR(3), bf.file_number) + '.ndf'
                    ELSE
                        @DestinationDB + '_' + CONVERT(VARCHAR(3), bf.file_number) + '.ndf'
                END + ''','
        FROM msdb.dbo.backupfile bf WITH (NOLOCK)
        LEFT JOIN msdb.dbo.backupset bs WITH (NOLOCK)
            ON bf.backup_set_id = bs.backup_set_id
        WHERE bf.backup_set_id = @first_backup_set_id
              AND bs.is_copy_only = 0;                     -- Add Jim to ignore copy only backups as they are non-recoverable

        SET @MOVETO_temp = @MOVETO;

    END;

    CREATE TABLE #RestoreCommand
    (
        ID INT NOT NULL IDENTITY(1, 1)
      , backup_start_date DATETIME
      , backup_finish_date DATETIME
      , database_backup_lsn NUMERIC(25, 0)
      , differential_base_lsn NUMERIC(25, 0)
      , first_lsn NUMERIC(25, 0)
      , last_lsn NUMERIC(25, 0)
      , postion INT
      , backup_type VARCHAR(20)
      , backup_size NUMERIC(20, 0)
      , physical_device_name VARCHAR(500)
      , backupset_name VARCHAR(500)
      , restore_command VARCHAR(MAX)
      , fileExist VARCHAR(20)
    );
    IF (@ListMode = 'All')
        SET @first_backup_set_id = 1;
    DECLARE backup_cursor CURSOR FOR
    SELECT bs.backup_start_date
         , bs.backup_finish_date
         , bs.database_backup_lsn
         , bs.differential_base_lsn
         , bs.first_lsn
         , bs.last_lsn
         , position
         , CASE bs.type
               WHEN 'D' THEN
                   'Full'
               WHEN 'L' THEN
                   'Log'
               WHEN 'I' THEN
                   'Diff'
           END AS backup_type
         , bs.backup_size
         , bmf.physical_device_name
         , bs.name AS backupset_name
    FROM msdb.dbo.backupmediafamily bmf WITH ( NOLOCK )
    INNER JOIN msdb.dbo.backupset bs WITH ( NOLOCK )
        ON bmf.media_set_id = bs.media_set_id
    WHERE bs.database_name = @SourceDB
          AND bs.backup_set_id >= @first_backup_set_id
          AND bs.is_copy_only = 0 -- Add Jim to ignore copy only backups as they are non-recoverable
    ORDER BY bs.backup_start_date;

    OPEN backup_cursor;
    FETCH NEXT FROM backup_cursor
    INTO @backup_start_date
       , @backup_finish_date
       , @database_backup_lsn
       , @differential_base_lsn
       , @first_lsn
       , @last_lsn
       , @position
       , @backup_type
       , @backup_size
       , @physical_device_name
       , @backupset_name;

    WHILE ((@@FETCH_STATUS = 0) AND (@lastfile <> 1))
    BEGIN
        --check if file exist
        EXEC master.dbo.xp_fileexist @physical_device_name, @Exists OUT;
        IF (@Exists = 1)
            SET @File_Exists = 'File Found';
        ELSE
            SET @File_Exists = 'File Not Found';

        IF (@backup_start_date <= @RestoreToTime)
        BEGIN
            --if this diff backup, then remove all log backup before it.
            IF ((@backup_type = 'Diff') AND (@ListMode = 'OnlyValid'))
                DELETE FROM #RestoreCommand
                WHERE backup_type IN ( 'Log', 'Diff' );

            IF @backup_type = 'Full'
            BEGIN
                SET @MOVETO_temp = @MOVETO;
                IF (@ListMode = 'OnlyValid')
                    DELETE FROM #RestoreCommand;
            END;
            ELSE
                SET @MOVETO_temp = '';
            SET @restore_command
                = CASE
                      WHEN @backup_type = 'Full' THEN
                          'RESTORE DATABASE [' + @DestinationDB + '] FROM DISK = N''' + @physical_device_name
                          + ''' WITH  FILE = ' + CONVERT(VARCHAR(3), @position) + ', ' + @MOVETO_temp
                          + 'NORECOVERY, NOUNLOAD, REPLACE, STATS = ' + CAST(@ProgressValue AS VARCHAR(20))
                      WHEN @backup_type = 'Diff' THEN
                          'RESTORE DATABASE [' + @DestinationDB + '] FROM DISK = N''' + @physical_device_name
                          + ''' WITH  FILE = ' + CONVERT(VARCHAR(3), @position) + ', ' + @MOVETO_temp
                          + 'NORECOVERY, NOUNLOAD, STATS = ' + CAST(@ProgressValue AS VARCHAR(20))
                      WHEN @backup_type = 'Log' THEN
                          'RESTORE LOG [' + @DestinationDB + '] FROM DISK = N''' + @physical_device_name
                          + ''' WITH  FILE = ' + CONVERT(VARCHAR(3), @position) + ', ' + @MOVETO_temp
                          + 'NORECOVERY, NOUNLOAD, STATS = ' + CAST(@ProgressValue AS VARCHAR(20))
                  END;

            INSERT INTO #RestoreCommand
            (
                backup_start_date
              , backup_finish_date
              , database_backup_lsn
              , differential_base_lsn
              , first_lsn
              , last_lsn
              , postion
              , backup_type
              , backup_size
              , physical_device_name
              , backupset_name
              , restore_command
              , fileExist
            )
            VALUES
            (@backup_start_date, @backup_finish_date, @database_backup_lsn, @differential_base_lsn, @first_lsn, @last_lsn, @position, @backup_type, @backup_size
           , @physical_device_name, @backupset_name, @restore_command, @File_Exists);
        END;
        ELSE IF (@backup_type = 'Log')
        BEGIN
            SET @lastfile = 1;
            SET @restore_command
                = 'RESTORE LOG [' + @DestinationDB + '] FROM  DISK = N''' + @physical_device_name + ''' WITH  FILE = ' + CONVERT(VARCHAR(3), @position)
                  + ',NORECOVERY, NOUNLOAD, STATS = ' + CAST(@ProgressValue AS VARCHAR(20));
            SET @restore_command = REPLACE(@restore_command, 'NORECOVERY', 'RECOVERY') + @stopat;
            INSERT INTO #RestoreCommand
            (
                backup_start_date
              , backup_finish_date
              , database_backup_lsn
              , differential_base_lsn
              , first_lsn
              , last_lsn
              , postion
              , backup_type
              , backup_size
              , physical_device_name
              , backupset_name
              , restore_command
              , fileExist
            )
            VALUES
            (@backup_start_date, @backup_finish_date, @database_backup_lsn, @differential_base_lsn, @first_lsn, @last_lsn, @position, @backup_type, @backup_size
           , @physical_device_name, @backupset_name, @restore_command, @File_Exists);
        END;
        FETCH NEXT FROM backup_cursor
        INTO @backup_start_date
           , @backup_finish_date
           , @database_backup_lsn
           , @differential_base_lsn
           , @first_lsn
           , @last_lsn
           , @position
           , @backup_type
           , @backup_size
           , @physical_device_name
           , @backupset_name;
    END;
    CLOSE backup_cursor;
    DEALLOCATE backup_cursor;

    -- accept last file as for recovery if not doing mirroring restore -- added by Jim
    IF (@lastfile <> 1 AND @MirroringRestore = 0)
    BEGIN
        -- Get max id
        DECLARE @MaxId INT;

        SET @MaxId =
        (
            SELECT MAX(ID) FROM #RestoreCommand
        );

        IF @MaxId IS NOT NULL
            UPDATE #RestoreCommand
            SET restore_command = REPLACE(restore_command, 'NORECOVERY', 'RECOVERY')
            WHERE ID = @MaxId;

    END;

    IF (@lastfile <> 1)
        INSERT INTO #RestoreCommand
        (
            restore_command
        )
        VALUES
        ('You need to back up the Tail of the Log on database [' + @SourceDB + '] before restoring, then restore the tail-log backup with recovery as last step!');

    SELECT *
    FROM #RestoreCommand
    ORDER BY [ID];
    DROP TABLE #RestoreCommand;

END;