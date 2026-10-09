USE [ILDB]
GO

/****** Object:  StoredProcedure [pct].[ReleaseMaster]    Script Date: 07/09/2026 15:37:41 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




--===============================================================================================================
-- Filename:	ReleaseMaster
-- Author:		Paul Davis / Chris Bartlett
-- Create date: 14/11/2024
-- Description:	Procedure to start PROCAT release process calling stored procedure
--				usp_editorial_find_ready_editset
--===============================================================================================================


CREATE OR ALTER PROCEDURE [pct].[ReleaseMaster]
AS
BEGIN
    -- Temporary table to store the results from sp_edit_find_ready_editsets
    DECLARE @EditSets TABLE (
        edit_set_id INT,
        name NVARCHAR(100)
    );

    -- Insert results from sp_edit_find_ready_editsets into the temporary table
    INSERT INTO @EditSets (edit_set_id, name)
    EXEC pct.FindReadyEditsets;

    DECLARE @t_iEditSetID INT;
	DECLARE @edit_set_type INT;
    DECLARE @name NVARCHAR(100);

    -- Loop through each row in @EditSets
    WHILE EXISTS (SELECT 1 FROM @EditSets)
    BEGIN
        -- Fetch the next row
        SELECT TOP 1 @t_iEditSetID = edit_set_id, @name = name
        FROM @EditSets;

		
		
		-- TODO: this is a bit daft - we can just return 'type' from  usp_editorial_find_ready_editsets
		--       but for now it mirrors closely what the original code did.

	    SELECT @edit_set_type = Edit_Set_Type
	    FROM edit_set
	    WHERE Edit_set_ID = @t_iEditSetID;

		PRINT 'Release master: processing ' + CAST(@t_iEditSetID AS NVARCHAR(10)) + ' type ' + CAST(@edit_set_type AS NVARCHAR(10))

	    -- Check the value of Edit_Set_Type and call the appropriate stored procedure
	    IF @edit_set_type = 4
	    BEGIN
	        -- Its a delete	
			-- create the files for the verity index file converter
			PRINT 'Release master: Deleting '
	
	        EXEC pct.ReleaseDeletionEditset @t_iEditSetID;
		
	    END
	    ELSE
			BEGIN
				-- Its an Add/Update
				-- TODO-  find moved files
				-- TODO - create the files for the verity index file converter - including the moved files (needs investigation)
				PRINT 'Release master: Inserting/Updtating '
				PRINT 'Release master: Releasing editset '
				EXEC pct.ReleaseEditset @t_iEditSetID;
		
				-- update title scopes	
				PRINT 'Release master: Updating lettercode title '
				EXEC pct.UpdateLettercodeTitle @t_iEditSetID;	
				
				PRINT 'Release master: Updating division title '
				EXEC pct.UpdateDivisionTitle  @t_iEditSetID;	
				
				PRINT 'Release master: Updating class title '
				EXEC pct.UpdateClassTitle  @t_iEditSetID;	
				
				PRINT 'Release master: Updating header title '
				EXEC pct.UpdateHeaderTitle  @t_iEditSetID;	
				
				PRINT 'Release master: Updating subheader title '
				EXEC pct.UpdateSubHeaderTitle  @t_iEditSetID;
				
				PRINT 'Release master: Updating piece title '
				EXEC pct.UpdatePieceTitle  @t_iEditSetID;	
				
				PRINT 'Release master: Updating item title '
				EXEC pct.UpdateItemTitle  @t_iEditSetID;
				
			END
		
		IF @@ERROR = 0 
		BEGIN
        -- Remove the processed row from @EditSets
        DELETE FROM @EditSets
        WHERE edit_set_id = @t_iEditSetID;

		EXEC pct.DeleteEditset @t_iEditSetID;	
		END
    END
END;
GO

