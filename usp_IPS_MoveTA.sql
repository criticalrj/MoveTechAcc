/* =============================================================================
   usp_IPS_MoveTA  -  Relocates a Technical Account (TA) to another code
   STAR IPS 3.62.x  (tested on 3.62.14 and 3.62.18)
   Author: Flavio Miguel

   CHANGE LOG
   Date        Version  Change
   ----------  -------  -------------------------------------------------------
   2026-10-22  2.3      FIX: free code when moving or renumbering a TA.
                        usp_IPS_MoveTA has a new optional last parameter
                        @To_Father. Before, the father of a new target TA was
                        always derived from the source suffix length (801.01
                        -> 701.01 needs father 701), so a code with another
                        level length, e.g. 731 -> 9999.731 inside 9999, was
                        impossible (the GUI blocked it with "format 9999N").
                        The GUI now always passes the father explicitly.
                        Without @To_Father the old rule still applies.
                        All comments translated to English.
                        Version marker "MoveTA-Procs-Version" added to every
                        procedure: the GUI updates outdated procedures by itself
                        and can drop them from the customer database.
   (earlier)   1.0-2.2  See the version history in README.md.

   Moves ONE TA (source -> target) inside a plant. To move several
   (e.g. 801.01, 801.04, 801.10 + subaccounts to 701.xx) call the proc once
   per pair, parents before children, inside ONE transaction (the
   Move-TA.ps1 GUI does that).

   What it does, in order:
     1. Validates plant and source TA (ST_Function + ST_PlantFunction + object).
     2. Ensures the target TA:
          - ST_Function      : reactivates it if deleted, or creates it (father =
                               @To_Father, or same depth as the source, or the
                               longest existing prefix in the class;
                               Function_Length/Function_OnLower kept).
          - ST_PlantFunction : reactivates or creates it copying ALL properties
                               of the source (criticality, maker, serial, free
                               text...).
          - ST_Object (type 2): via st_CreateObjectId, as IPS does.
          - ST_PLANTFUNCEQ   : copies the equipment specification, if any.
     3. Moves the connections (soft delete of the old row + insert/reactivate
        of the new one, never changes a PK, never deletes physically - safe
        for replication):
          - ST_PlantFunctionObject  (jobs 320, spare parts 221, counters 210,
                                     equipment, etc.)
          - ST_WPWHERE2             (work plan where-list, by the TA OBJECT_ID)
          - ST_WHWHERE2             (work history where-list)
          - ST_WorkHistory          (WH_Function_Id / Name / PLANTFUNC_*)
          - ST_POAITEM2             (MR/PO lines)
          - ST_MATONHANDSTAT        (default TA of the material in the plant)
          - ST_MATTRANS2            (material transaction history)
          - ST_BudItem, ST_PlantCost, ST_UserGroupRestrictFunc (direct update)
     4. Optional: retires the source TA in the plant (and in the class, if
        nobody else uses it).
     5. Returns a result set with the log (Tabela, Acao, Linhas, Detalhe).
        The log text stays in Portuguese.

   Mapping from the "Move TA_3.53_r5.xls" spreadsheet:
     3.53  ST_WPWHERE  (WP_Object_Id, FuncClass_Id, Function_Id)
     3.62  ST_WPWHERE2 (WP_OBJECT_ID, OBJECT_ID = ST_PlantFunction.PLANTFUNC_OBJECT_ID)
     3.53  ST_WHWHERE  (WH_Object_Id, WHW_FuncClass_Id, WHW_Function_Id)
     3.62  ST_WHWHERE2 (WH_OBJECT_ID, OBJECT_ID = ST_PlantFunction.PLANTFUNC_OBJECT_ID)
     The old tables ST_WPWhere/ST_WHWhere still exist but are empty in 3.62;
     the views VI_WPWHERE/VI_WHWHERE already read from the *2 tables.

   Transactions: when called with no open transaction, it opens its own and
   commits/rolls back. When called inside a transaction (GUI), it joins it and,
   on error, THROWs so the caller can ROLLBACK. @DryRun=1 runs everything and
   undoes it at the end (savepoint when there is an outer transaction).
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
/* -----------------------------------------------------------------------------
   usp_IPS_RetireTA - retires (soft delete) a TA in the plant, if nothing
   active is still linked to it. Called by usp_IPS_MoveTA and also by the GUI
   in a second pass (children first, then the parents, which only become
   "empty" after all subaccounts were moved).
   Does not open its own transaction: joins the caller's (or autocommit).
   ----------------------------------------------------------------------------- */
CREATE OR ALTER PROCEDURE dbo.usp_IPS_RetireTA
    @Plant_Id        varchar(20),
    @Function_Id     varchar(100),
    @RetireInClass   bit = 0,              -- also retire in ST_Function if no plant of the class uses it
    @Rec_User        varchar(100) = NULL
AS
BEGIN
    -- MoveTA-Procs-Version: 2.3   (Move-TA.ps1 compares this marker with the bundled file and updates the procedures; bump it on every procedure change)
    SET NOCOUNT ON;
    DECLARE @log TABLE (Seq int IDENTITY(1,1), Tabela sysname, Acao varchar(30), Linhas int, Detalhe varchar(400));
    DECLARE @FC varchar(10), @obj numeric(15,0), @u varchar(100), @left int = 0, @kids int = 0, @Father varchar(100), @msg varchar(400);
    SET @u = ISNULL(NULLIF(LTRIM(RTRIM(@Rec_User)),''), UPPER(SUSER_SNAME()));
    SET @Function_Id = LTRIM(RTRIM(@Function_Id));

    SELECT @FC = FuncClass_Id FROM ST_Plant WHERE Plant_Id = @Plant_Id AND Rec_Deleted = 0;
    IF @FC IS NULL BEGIN SET @msg = 'Planta ' + @Plant_Id + ' nao encontrada.'; THROW 50003, @msg, 1; END
    SELECT @obj = PLANTFUNC_OBJECT_ID FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Function_Id AND Rec_Deleted = 0;
    IF @obj IS NULL
    BEGIN
        INSERT @log VALUES ('ST_PlantFunction', 'ja inativa', 0, @Function_Id + ' nao esta ativa na planta ' + @Plant_Id);
        SELECT Tabela, Acao, Linhas, Detalhe FROM @log ORDER BY Seq; RETURN;
    END

    SELECT @left =
          (SELECT COUNT(*) FROM ST_PlantFunctionObject WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Function_Id AND Rec_Deleted = 0)
        + (SELECT COUNT(*) FROM ST_WPWHERE2 WHERE OBJECT_ID = @obj AND REC_DELETED = 0)
        + (SELECT COUNT(*) FROM ST_WHWHERE2 WHERE OBJECT_ID = @obj AND REC_DELETED = 0)
        + (SELECT COUNT(*) FROM ST_WorkHistory WHERE Plant_Id = @Plant_Id AND WH_FuncClass_Id = @FC AND WH_Function_Id = @Function_Id AND Rec_Deleted = 0)
        + (SELECT COUNT(*) FROM ST_POAITEM2 p JOIN ST_Object o ON o.Object_Id = p.POITEM_OBJECT_ID WHERE o.ObjClass_Id = @Plant_Id AND p.FUNCCLASS_ID = @FC AND p.FUNCTION_ID = @Function_Id AND p.REC_DELETED = 0)
        + (SELECT COUNT(*) FROM ST_MATONHANDSTAT WHERE PLANT_ID = @Plant_Id AND MatOH_Def_FuncClass_Id = @FC AND MatOH_Def_Function_Id = @Function_Id);
    SELECT @kids = COUNT(*) FROM ST_Function c JOIN ST_PlantFunction pf ON pf.FuncClass_Id = c.FuncClass_Id AND pf.Function_Id = c.Function_Id AND pf.Plant_Id = @Plant_Id AND pf.Rec_Deleted = 0
     WHERE c.FuncClass_Id = @FC AND c.Function_Father = @Function_Id AND c.Rec_Deleted = 0;

    IF @left > 0 OR @kids > 0
        INSERT @log VALUES ('ST_PlantFunction', 'MANTIDA', 0, 'TA ' + @Function_Id + ' NAO desativada: ' + CAST(@left AS varchar(10)) + ' referencias ativas restantes, ' + CAST(@kids AS varchar(10)) + ' subaccounts ativos na planta');
    ELSE
    BEGIN
        UPDATE ST_PlantFunction SET Rec_Deleted = -1, Rec_Revisor = @u WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Function_Id;
        UPDATE ST_PLANTFUNCEQ SET REC_DELETED = -1, REC_REVISOR = @u WHERE PLANT_ID = @Plant_Id AND FUNCCLASS_ID = @FC AND FUNCTION_ID = @Function_Id AND REC_DELETED = 0;
        UPDATE ST_Object SET Rec_Deleted = -1, Rec_Revisor = @u WHERE Object_Id = @obj;
        INSERT @log VALUES ('ST_PlantFunction', 'desativada', 1, 'TA ' + @Function_Id + ' desativada na planta ' + @Plant_Id + ' (objeto ' + CAST(@obj AS varchar(20)) + ')');

        IF @RetireInClass = 1
        BEGIN
            IF EXISTS (SELECT 1 FROM ST_PlantFunction WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id AND Rec_Deleted = 0)
                INSERT @log VALUES ('ST_Function', 'MANTIDA', 0, 'outras plantas da classe ' + @FC + ' ainda usam ' + @Function_Id);
            ELSE IF EXISTS (SELECT 1 FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Father = @Function_Id AND Rec_Deleted = 0)
                INSERT @log VALUES ('ST_Function', 'MANTIDA', 0, @Function_Id + ' ainda tem subaccounts ativos na classe');
            ELSE
            BEGIN
                SELECT @Father = Function_Father FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id;
                UPDATE ST_Function SET Rec_Deleted = -1, Rec_Revisor = @u WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id;
                UPDATE f SET Function_OnLower = (SELECT COUNT(*) FROM ST_Function c WHERE c.FuncClass_Id = f.FuncClass_Id AND c.Function_Father = f.Function_Id AND c.Rec_Deleted = 0), Rec_Revisor = @u
                  FROM ST_Function f WHERE f.FuncClass_Id = @FC AND f.Function_Id = @Father;
                INSERT @log VALUES ('ST_Function', 'desativada', 1, @Function_Id + ' desativada na classe ' + @FC);
            END
        END
    END
    SELECT Tabela, Acao, Linhas, Detalhe FROM @log ORDER BY Seq;
END
GO
CREATE OR ALTER PROCEDURE dbo.usp_IPS_MoveTA
    @Plant_Id            varchar(20),
    @From_Function_Id    varchar(100),
    @To_Function_Id      varchar(100),
    @To_Function_Name    varchar(100) = NULL,   -- target TA name if it is created (default: source name)
    @CreateMissing       bit = 1,               -- create/reactivate ST_Function, ST_PlantFunction and object of the target
    @MoveWorkHistory     bit = 1,
    @MovePOItems         bit = 1,
    @MoveMaterialRefs    bit = 1,               -- ST_MATONHANDSTAT (default TA) + ST_MATTRANS2 (history)
    @MoveOtherRefs       bit = 1,               -- ST_BudItem, ST_PlantCost, ST_UserGroupRestrictFunc
    @RetireSource        bit = 1,               -- retire the source TA in the plant if nothing is left linked to it
    @RetireSourceInClass bit = 0,               -- also retire in ST_Function if no plant of the class uses it
    @ListItems           bit = 1,               -- list every object moved via ST_PlantFunctionObject in the log
    @Rec_User            varchar(100) = NULL,   -- Rec_Creator/Rec_Revisor (default: current login)
    @DryRun              bit = 1,               -- 1 = simulate (undo at the end) ; 0 = save
    @To_Father           varchar(100) = NULL    -- 2.3: father of the target if it is created (default: same depth as the source)
AS
BEGIN
    -- MoveTA-Procs-Version: 2.3   (Move-TA.ps1 compares this marker with the bundled file and updates the procedures; bump it on every procedure change)
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @log TABLE (Seq int IDENTITY(1,1), Tabela sysname, Acao varchar(30), Linhas int, Detalhe varchar(400));
    DECLARE @FC varchar(10), @S_obj numeric(15,0), @D_obj numeric(15,0),
            @FromName varchar(100), @u varchar(100), @n int, @n2 int, @msg varchar(400),
            @Father varchar(100), @DestFuncState varchar(20), @DestPFState varchar(20),
            @outer bit, @sql nvarchar(max), @cols nvarchar(max);

    SET @u = ISNULL(NULLIF(LTRIM(RTRIM(@Rec_User)),''), UPPER(SUSER_SNAME()));
    SET @From_Function_Id = LTRIM(RTRIM(@From_Function_Id));
    SET @To_Function_Id   = LTRIM(RTRIM(@To_Function_Id));
    SET @To_Function_Name = NULLIF(LTRIM(RTRIM(@To_Function_Name)),'');
    SET @To_Father        = NULLIF(LTRIM(RTRIM(@To_Father)),'');

    /* ---------------- validations that need no transaction ---------------- */
    IF @From_Function_Id IS NULL OR @From_Function_Id = '' OR @To_Function_Id IS NULL OR @To_Function_Id = ''
        THROW 50001, 'Informe @From_Function_Id e @To_Function_Id.', 1;
    IF @From_Function_Id = @To_Function_Id
        THROW 50002, 'Origem e destino sao iguais.', 1;
    IF @To_Father IS NOT NULL AND (LEN(@To_Function_Id) <= LEN(@To_Father) OR LEFT(@To_Function_Id, LEN(@To_Father)) <> @To_Father)
    BEGIN
        SET @msg = 'Destino ' + @To_Function_Id + ' precisa comecar com o pai ' + @To_Father + ' e ser mais longo.';
        THROW 50008, @msg, 1;
    END

    SELECT @FC = FuncClass_Id FROM ST_Plant WHERE Plant_Id = @Plant_Id AND Rec_Deleted = 0;
    IF @FC IS NULL
    BEGIN
        SET @msg = 'Planta ' + @Plant_Id + ' nao encontrada (ST_Plant).';
        THROW 50003, @msg, 1;
    END

    SELECT @FromName = Function_Name FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @From_Function_Id AND Rec_Deleted = 0;
    IF @FromName IS NULL
    BEGIN
        SET @msg = 'TA origem ' + @From_Function_Id + ' nao existe (ativa) na classe ' + @FC + '.';
        THROW 50004, @msg, 1;
    END

    SELECT @S_obj = PLANTFUNC_OBJECT_ID FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @From_Function_Id AND Rec_Deleted = 0;
    IF @S_obj IS NULL
    BEGIN
        SET @msg = 'TA origem ' + @From_Function_Id + ' nao esta ativa na planta ' + @Plant_Id + ' ou nao tem PLANTFUNC_OBJECT_ID (base anterior a 3.60?).';
        THROW 50005, @msg, 1;
    END

    /* ---------------- transaction ---------------- */
    SET @outer = CASE WHEN @@TRANCOUNT > 0 THEN 1 ELSE 0 END;
    IF @outer = 1 SAVE TRANSACTION MoveTA; ELSE BEGIN TRANSACTION;

    BEGIN TRY

        INSERT @log VALUES ('-', 'inicio', NULL, 'Planta ' + @Plant_Id + ' / classe ' + @FC + ' : ' + @From_Function_Id + ' (' + ISNULL(@FromName,'') + ') -> ' + @To_Function_Id + CASE WHEN @DryRun = 1 THEN '   [SIMULACAO]' ELSE '' END);

        /* ======================= 2. Target TA in ST_Function ======================= */
        IF EXISTS (SELECT 1 FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @To_Function_Id AND Rec_Deleted = 0)
        BEGIN
            SET @DestFuncState = 'existe';
            IF @To_Function_Name IS NULL SELECT @To_Function_Name = Function_Name FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @To_Function_Id AND Rec_Deleted = 0;
        END
        ELSE
        BEGIN
            IF @CreateMissing = 0
            BEGIN
                SET @msg = 'TA destino ' + @To_Function_Id + ' nao existe na classe ' + @FC + ' e @CreateMissing = 0.';
                THROW 50006, @msg, 1;
            END
            IF @To_Function_Name IS NULL SET @To_Function_Name = @FromName;

            -- father: @To_Father when given (2.3 fix: the GUI always passes it, which allows
            -- a free code, e.g. 731 -> 9999.731 inside 9999). Otherwise same depth as the
            -- source: 801.01 has father 801 (suffix '.01' = 3 chars) => 701.01 needs father 701.
            -- If the source is a root, use the longest existing prefix.
            DECLARE @srcFather varchar(100), @sufLen int;
            SELECT @srcFather = Function_Father FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @From_Function_Id;
            SET @sufLen = CASE WHEN @srcFather IS NULL THEN NULL ELSE LEN(@From_Function_Id) - LEN(@srcFather) END;
            IF @To_Father IS NOT NULL
                SET @Father = @To_Father;
            ELSE IF @sufLen IS NOT NULL AND @sufLen > 0 AND LEN(@To_Function_Id) > @sufLen
                SET @Father = LEFT(@To_Function_Id, LEN(@To_Function_Id) - @sufLen);
            ELSE
                SELECT TOP 1 @Father = Function_Id FROM ST_Function
                 WHERE FuncClass_Id = @FC AND Rec_Deleted = 0 AND LEN(Function_Id) < LEN(@To_Function_Id)
                   AND LEFT(@To_Function_Id, LEN(Function_Id)) = Function_Id
                 ORDER BY LEN(Function_Id) DESC;
            IF @Father IS NULL OR NOT EXISTS (SELECT 1 FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @Father AND Rec_Deleted = 0)
            BEGIN
                SET @msg = 'TA pai ' + ISNULL(@Father,'?') + ' nao existe (ativa) na classe ' + @FC + ' para criar ' + @To_Function_Id + '. Crie o grupo pai antes.';
                THROW 50007, @msg, 1;
            END

            IF EXISTS (SELECT 1 FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @To_Function_Id)   -- exists, deleted
            BEGIN
                UPDATE ST_Function
                   SET Rec_Deleted = 0, Function_Name = @To_Function_Name, Function_Father = @Father,
                       Function_Length = LEN(@To_Function_Id), Rec_Revisor = @u
                 WHERE FuncClass_Id = @FC AND Function_Id = @To_Function_Id;
                SET @DestFuncState = 'reativada';
            END
            ELSE
            BEGIN
                INSERT ST_Function (FuncClass_Id, Function_Id, Function_Father, Function_ThisLevelId, Function_Name, Function_Length, Function_OnLower, Rec_Deleted, Rec_Creator, Rec_Revisor, Function_FreeText)
                SELECT @FC, @To_Function_Id, @Father, NULL, @To_Function_Name, LEN(@To_Function_Id), NULL, 0, @u, @u, Function_FreeText
                  FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @From_Function_Id;
                SET @DestFuncState = 'criada';
            END
            -- recount the father's children
            UPDATE f SET Function_OnLower = (SELECT COUNT(*) FROM ST_Function c WHERE c.FuncClass_Id = f.FuncClass_Id AND c.Function_Father = f.Function_Id AND c.Rec_Deleted = 0), Rec_Revisor = @u
              FROM ST_Function f WHERE f.FuncClass_Id = @FC AND f.Function_Id = @Father;
        END
        INSERT @log VALUES ('ST_Function', @DestFuncState, 1, @To_Function_Id + ' (' + ISNULL(@To_Function_Name,'') + ')' + CASE WHEN @Father IS NOT NULL THEN '  pai=' + @Father ELSE '' END);

        /* ======================= 2b. Target TA in ST_PlantFunction + ST_Object ======================= */
        SELECT @D_obj = PLANTFUNC_OBJECT_ID FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @To_Function_Id AND Rec_Deleted = 0;
        IF @D_obj IS NOT NULL
        BEGIN
            SET @DestPFState = 'existe';
        END
        ELSE
        BEGIN
            IF @CreateMissing = 0
            BEGIN
                SET @msg = 'TA destino ' + @To_Function_Id + ' nao esta ativa na planta ' + @Plant_Id + ' e @CreateMissing = 0.';
                THROW 50008, @msg, 1;
            END
            IF EXISTS (SELECT 1 FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @To_Function_Id)  -- exists, deleted
            BEGIN
                SELECT @D_obj = PLANTFUNC_OBJECT_ID FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @To_Function_Id;
                IF @D_obj IS NULL
                BEGIN
                    EXEC dbo.st_CreateObjectId @D_obj OUTPUT, @Plant_Id, 2, @To_Function_Id, @To_Function_Name, 0, 0, @u;
                    UPDATE ST_PlantFunction SET PLANTFUNC_OBJECT_ID = @D_obj WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @To_Function_Id;
                END
                UPDATE ST_PlantFunction SET Rec_Deleted = 0, Rec_Revisor = @u WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @To_Function_Id;
                UPDATE ST_Object SET Rec_Deleted = 0, Obj_Description = @To_Function_Name, Rec_Revisor = @u WHERE Object_Id = @D_obj;
                UPDATE ST_PLANTFUNCEQ SET REC_DELETED = 0, REC_REVISOR = @u WHERE PLANT_ID = @Plant_Id AND FUNCCLASS_ID = @FC AND FUNCTION_ID = @To_Function_Id;
                SET @DestPFState = 'reativada';
            END
            ELSE
            BEGIN
                -- TA object (ObjType 2), as IPS and trigger TI_ST_PLANTFUNCTION do
                EXEC dbo.st_CreateObjectId @D_obj OUTPUT, @Plant_Id, 2, @To_Function_Id, @To_Function_Name, 0, 0, @u;
                IF ISNULL(@D_obj, 0) = 0 THROW 50009, 'st_CreateObjectId nao devolveu Object_Id para a TA destino.', 1;

                INSERT ST_PlantFunction
                       (Plant_Id, FuncClass_Id, Function_Id, Criticality_Id, Image_Object_Id, Template_Id, Condition_Id, PlantFunc_CondDate,
                        PlantFunc_Id, PlantFunc_RefId, PlantFunc_Location, PlantFunc_InstDate, PlantFunc_QtyUnit, PlantFunc_Quantity,
                        PlantFunc_CountValue, PlantFunc_CountDate, PlantFunc_CountEvPrDay, PlantFunc_Remarks, PlantFunc_Maker, PlantFunc_MakersNo,
                        PlantFunc_Model, PlantFunc_SerialNo, PlantFunc_IsEquipment, PlantFunc_IsStore, PlantFunc_CondRemark, PlantFunc_ImageZoom,
                        PlantFunc_Owner, Rec_Deleted, Rec_Creator, Rec_Revisor, PlantFunc_PlainText, PlantFunc_FreeText, PLANTFUNC_NOX,
                        TECHCLASS_ID, TECHTYPE_ID, PLANTFUNC_MANUFDATE, PLANTFUNC_COMMISSION, PLANTFUNC_WARRANTYDATE, Depart_Id, Account_Id,
                        Maker_AddrObj_Id, PlantFunc_ClassRel, PlantFunc_ISMChapter_ID, PlantFunc_ISMQuest_ID, PLANTFUNC_MASTERDRAW_OBJ_ID,
                        PLANTFUNC_MASTERDRAW_POSITION, PLANTFUNC_MASTERDRAW_PAGE, PLANTFUNC_OBJECT_ID, OPERATIONMODE_ID, PLANTFUNC_PMOD)
                SELECT  Plant_Id, FuncClass_Id, @To_Function_Id, Criticality_Id, Image_Object_Id, Template_Id, Condition_Id, PlantFunc_CondDate,
                        PlantFunc_Id, PlantFunc_RefId, PlantFunc_Location, PlantFunc_InstDate, PlantFunc_QtyUnit, PlantFunc_Quantity,
                        PlantFunc_CountValue, PlantFunc_CountDate, PlantFunc_CountEvPrDay, PlantFunc_Remarks, PlantFunc_Maker, PlantFunc_MakersNo,
                        PlantFunc_Model, PlantFunc_SerialNo, PlantFunc_IsEquipment, PlantFunc_IsStore, PlantFunc_CondRemark, PlantFunc_ImageZoom,
                        PlantFunc_Owner, 0, @u, @u, PlantFunc_PlainText, PlantFunc_FreeText, PLANTFUNC_NOX,
                        TECHCLASS_ID, TECHTYPE_ID, PLANTFUNC_MANUFDATE, PLANTFUNC_COMMISSION, PLANTFUNC_WARRANTYDATE, Depart_Id, Account_Id,
                        Maker_AddrObj_Id, PlantFunc_ClassRel, PlantFunc_ISMChapter_ID, PlantFunc_ISMQuest_ID, PLANTFUNC_MASTERDRAW_OBJ_ID,
                        PLANTFUNC_MASTERDRAW_POSITION, PLANTFUNC_MASTERDRAW_PAGE, @D_obj, OPERATIONMODE_ID, PLANTFUNC_PMOD
                  FROM ST_PlantFunction
                 WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @From_Function_Id;

                -- equipment specification (ST_PLANTFUNCEQ): dynamic copy of all PFEQ_* columns
                IF EXISTS (SELECT 1 FROM ST_PLANTFUNCEQ WHERE PLANT_ID = @Plant_Id AND FUNCCLASS_ID = @FC AND FUNCTION_ID = @From_Function_Id AND REC_DELETED = 0)
                   AND NOT EXISTS (SELECT 1 FROM ST_PLANTFUNCEQ WHERE PLANT_ID = @Plant_Id AND FUNCCLASS_ID = @FC AND FUNCTION_ID = @To_Function_Id)
                BEGIN
                    SELECT @cols = STUFF((SELECT ', ' + QUOTENAME(c.name) FROM sys.columns c WHERE c.object_id = OBJECT_ID('dbo.ST_PLANTFUNCEQ') AND c.name LIKE 'PFEQ[_]%' ORDER BY c.column_id FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 2, '');
                    SET @sql = N'INSERT ST_PLANTFUNCEQ (PLANT_ID, FUNCCLASS_ID, FUNCTION_ID, REC_DELETED, REC_CREATOR, REC_REVISOR, ' + @cols + N')
                                 SELECT PLANT_ID, FUNCCLASS_ID, @To, 0, @u, @u, ' + @cols + N'
                                   FROM ST_PLANTFUNCEQ WHERE PLANT_ID = @P AND FUNCCLASS_ID = @FC AND FUNCTION_ID = @From AND REC_DELETED = 0';
                    EXEC sp_executesql @sql, N'@To varchar(100), @u varchar(100), @P varchar(20), @FC varchar(10), @From varchar(100)', @To_Function_Id, @u, @Plant_Id, @FC, @From_Function_Id;
                    INSERT @log VALUES ('ST_PLANTFUNCEQ', 'copiada', @@ROWCOUNT, 'especificacao de equipamento copiada da origem');
                END
                SET @DestPFState = 'criada';
            END
        END
        INSERT @log VALUES ('ST_PlantFunction', @DestPFState, 1, @To_Function_Id + '  objeto TA origem=' + CAST(@S_obj AS varchar(20)) + '  destino=' + CAST(@D_obj AS varchar(20)));

        /* ======================= 3a. ST_PlantFunctionObject ======================= */
        IF @ListItems = 1
            INSERT @log (Tabela, Acao, Linhas, Detalhe)
            SELECT 'ST_PlantFunctionObject', 'item', 1,
                   RTRIM(CAST(pfo.ObjType_Id AS varchar(10))) + ' ' + ISNULL(ot.ObjType_Descr,'') + ': ' + ISNULL(o.Obj_Code,'') + ' - ' + ISNULL(o.Obj_Description,'')
              FROM ST_PlantFunctionObject pfo
              JOIN ST_Object o ON o.Object_Id = pfo.Object_Id
              LEFT JOIN ST_ObjectType ot ON ot.ObjType_Id = pfo.ObjType_Id
             WHERE pfo.Plant_Id = @Plant_Id AND pfo.FuncClass_Id = @FC AND pfo.Function_Id = @From_Function_Id AND pfo.Rec_Deleted = 0
             ORDER BY pfo.ObjType_Id, o.Obj_Code;

        -- reactivate/update rows that already exist on the target (same object)
        UPDATE d SET Rec_Deleted = 0, ObjType_Id = s.ObjType_Id, PFO_String = s.PFO_String, PFO_Double = s.PFO_Double, PFO_Date = s.PFO_Date, PFO_QUALIFYER = s.PFO_QUALIFYER, Rec_Revisor = @u
          FROM ST_PlantFunctionObject d
          JOIN ST_PlantFunctionObject s ON s.Plant_Id = d.Plant_Id AND s.FuncClass_Id = d.FuncClass_Id AND s.Object_Id = d.Object_Id
         WHERE d.Plant_Id = @Plant_Id AND d.FuncClass_Id = @FC AND d.Function_Id = @To_Function_Id
           AND s.Function_Id = @From_Function_Id AND s.Rec_Deleted = 0;
        SET @n = @@ROWCOUNT;

        INSERT ST_PlantFunctionObject (Plant_Id, FuncClass_Id, Function_Id, Object_Id, ObjType_Id, PFO_String, PFO_Double, PFO_Date, Rec_Deleted, Rec_Creator, Rec_Revisor, PFO_QUALIFYER)
        SELECT s.Plant_Id, s.FuncClass_Id, @To_Function_Id, s.Object_Id, s.ObjType_Id, s.PFO_String, s.PFO_Double, s.PFO_Date, 0, @u, @u, s.PFO_QUALIFYER
          FROM ST_PlantFunctionObject s
         WHERE s.Plant_Id = @Plant_Id AND s.FuncClass_Id = @FC AND s.Function_Id = @From_Function_Id AND s.Rec_Deleted = 0
           AND NOT EXISTS (SELECT 1 FROM ST_PlantFunctionObject d WHERE d.Plant_Id = s.Plant_Id AND d.FuncClass_Id = s.FuncClass_Id AND d.Function_Id = @To_Function_Id AND d.Object_Id = s.Object_Id);
        SET @n2 = @@ROWCOUNT;

        UPDATE ST_PlantFunctionObject SET Rec_Deleted = -1, Rec_Revisor = @u
         WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @From_Function_Id AND Rec_Deleted = 0;
        INSERT @log VALUES ('ST_PlantFunctionObject', 'movidas', @@ROWCOUNT, CAST(@n2 AS varchar(10)) + ' inseridas no destino, ' + CAST(@n AS varchar(10)) + ' ja existiam (merge)');

        /* ======================= 3b. ST_WPWHERE2 (work plan where-list) ======================= */
        UPDATE d SET REC_DELETED = 0,
                     WPWHERE_MAIN = CASE WHEN ISNULL(s.WPWHERE_MAIN,0) = -1 OR ISNULL(d.WPWHERE_MAIN,0) = -1 THEN -1 ELSE d.WPWHERE_MAIN END,
                     WPWHERE_OK = ISNULL(s.WPWHERE_OK, d.WPWHERE_OK), WPWHERE_COMMENTS = ISNULL(s.WPWHERE_COMMENTS, d.WPWHERE_COMMENTS),
                     WPWHERE_EV_OBJECT_ID = ISNULL(s.WPWHERE_EV_OBJECT_ID, d.WPWHERE_EV_OBJECT_ID),
                     WPWHERE_COND_ID_BEFORE = ISNULL(s.WPWHERE_COND_ID_BEFORE, d.WPWHERE_COND_ID_BEFORE), WPWHERE_COND_ID_AFTER = ISNULL(s.WPWHERE_COND_ID_AFTER, d.WPWHERE_COND_ID_AFTER),
                     WPWHERESTATUS_ID = ISNULL(s.WPWHERESTATUS_ID, d.WPWHERESTATUS_ID), WPWHERE_WO_OBJECT_ID = ISNULL(s.WPWHERE_WO_OBJECT_ID, d.WPWHERE_WO_OBJECT_ID),
                     WPWHERE_SORTORDER = ISNULL(s.WPWHERE_SORTORDER, d.WPWHERE_SORTORDER), REC_REVISOR = @u
          FROM ST_WPWHERE2 d JOIN ST_WPWHERE2 s ON s.WP_OBJECT_ID = d.WP_OBJECT_ID
         WHERE d.OBJECT_ID = @D_obj AND s.OBJECT_ID = @S_obj AND s.REC_DELETED = 0;
        SET @n = @@ROWCOUNT;

        INSERT ST_WPWHERE2 (WP_OBJECT_ID, OBJECT_ID, WPWHERE_MAIN, WPWHERE_OK, WPWHERE_COMMENTS, WPWHERE_EV_OBJECT_ID, WPWHERE_COND_ID_BEFORE, WPWHERE_COND_ID_AFTER, WPWHERESTATUS_ID, WPWHERE_WO_OBJECT_ID, WPWHERE_SORTORDER, REC_DELETED, REC_CREATOR, REC_REVISOR)
        SELECT s.WP_OBJECT_ID, @D_obj, s.WPWHERE_MAIN, s.WPWHERE_OK, s.WPWHERE_COMMENTS, s.WPWHERE_EV_OBJECT_ID, s.WPWHERE_COND_ID_BEFORE, s.WPWHERE_COND_ID_AFTER, s.WPWHERESTATUS_ID, s.WPWHERE_WO_OBJECT_ID, s.WPWHERE_SORTORDER, 0, @u, @u
          FROM ST_WPWHERE2 s
         WHERE s.OBJECT_ID = @S_obj AND s.REC_DELETED = 0
           AND NOT EXISTS (SELECT 1 FROM ST_WPWHERE2 d WHERE d.WP_OBJECT_ID = s.WP_OBJECT_ID AND d.OBJECT_ID = @D_obj);
        SET @n2 = @@ROWCOUNT;

        UPDATE ST_WPWHERE2 SET REC_DELETED = -1, REC_REVISOR = @u WHERE OBJECT_ID = @S_obj AND REC_DELETED = 0;
        INSERT @log VALUES ('ST_WPWHERE2', 'movidas', @@ROWCOUNT, CAST(@n2 AS varchar(10)) + ' inseridas, ' + CAST(@n AS varchar(10)) + ' merge (jobs/WOs/atividades que tinham a TA na where-list)');

        -- legacy table (only if it still has rows; in 3.62 it is usually empty)
        IF EXISTS (SELECT 1 FROM ST_WPWhere WHERE FuncClass_Id = @FC AND Function_Id = @From_Function_Id AND Rec_Deleted = 0)
        BEGIN
            UPDATE d SET Rec_Deleted = 0, WPWhere_Main = s.WPWhere_Main, WPWhere_Ok = s.WPWhere_Ok, WPWhere_Comments = s.WPWhere_Comments, WPWhere_Ev_Object_Id = s.WPWhere_Ev_Object_Id, Rec_Revisor = @u
              FROM ST_WPWhere d JOIN ST_WPWhere s ON s.WP_Object_Id = d.WP_Object_Id AND s.FuncClass_Id = d.FuncClass_Id
             WHERE d.Function_Id = @To_Function_Id AND s.Function_Id = @From_Function_Id AND s.Rec_Deleted = 0
               AND EXISTS (SELECT 1 FROM ST_Object o WHERE o.Object_Id = s.WP_Object_Id AND o.ObjClass_Id = @Plant_Id);
            INSERT ST_WPWhere (WP_Object_Id, FuncClass_Id, Function_Id, WPWhere_Main, WPWhere_Ok, WPWhere_Comments, WPWhere_Ev_Object_Id, Rec_Deleted, Rec_Creator, Rec_Revisor, WPWHERE_COND_ID_BEFORE, WPWHERE_COND_ID_AFTER, WPWHERESTATUS_ID, WPWHERE_WO_OBJECT_ID, WPWHERE_SORTORDER)
            SELECT s.WP_Object_Id, s.FuncClass_Id, @To_Function_Id, s.WPWhere_Main, s.WPWhere_Ok, s.WPWhere_Comments, s.WPWhere_Ev_Object_Id, 0, @u, @u, s.WPWHERE_COND_ID_BEFORE, s.WPWHERE_COND_ID_AFTER, s.WPWHERESTATUS_ID, s.WPWHERE_WO_OBJECT_ID, s.WPWHERE_SORTORDER
              FROM ST_WPWhere s
             WHERE s.FuncClass_Id = @FC AND s.Function_Id = @From_Function_Id AND s.Rec_Deleted = 0
               AND EXISTS (SELECT 1 FROM ST_Object o WHERE o.Object_Id = s.WP_Object_Id AND o.ObjClass_Id = @Plant_Id)
               AND NOT EXISTS (SELECT 1 FROM ST_WPWhere d WHERE d.WP_Object_Id = s.WP_Object_Id AND d.FuncClass_Id = s.FuncClass_Id AND d.Function_Id = @To_Function_Id);
            UPDATE s SET Rec_Deleted = -1, Rec_Revisor = @u FROM ST_WPWhere s
             WHERE s.FuncClass_Id = @FC AND s.Function_Id = @From_Function_Id AND s.Rec_Deleted = 0
               AND EXISTS (SELECT 1 FROM ST_Object o WHERE o.Object_Id = s.WP_Object_Id AND o.ObjClass_Id = @Plant_Id);
            INSERT @log VALUES ('ST_WPWhere (legada)', 'movidas', @@ROWCOUNT, 'tabela antiga ainda tinha linhas; movidas tambem');
        END

        /* ======================= 3c. ST_WHWHERE2 + ST_WorkHistory ======================= */
        IF @MoveWorkHistory = 1
        BEGIN
            UPDATE d SET REC_DELETED = 0,
                         WHWHERE_MAIN = CASE WHEN ISNULL(s.WHWHERE_MAIN,0) = -1 OR ISNULL(d.WHWHERE_MAIN,0) = -1 THEN -1 ELSE d.WHWHERE_MAIN END,
                         WHWHERE_OK = ISNULL(s.WHWHERE_OK, d.WHWHERE_OK), WHWHERE_COMMENTS = ISNULL(s.WHWHERE_COMMENTS, d.WHWHERE_COMMENTS),
                         WHWHERE_EV_OBJECT_ID = ISNULL(s.WHWHERE_EV_OBJECT_ID, d.WHWHERE_EV_OBJECT_ID), REC_REVISOR = @u
              FROM ST_WHWHERE2 d JOIN ST_WHWHERE2 s ON s.WH_OBJECT_ID = d.WH_OBJECT_ID
             WHERE d.OBJECT_ID = @D_obj AND s.OBJECT_ID = @S_obj AND s.REC_DELETED = 0;
            SET @n = @@ROWCOUNT;

            INSERT ST_WHWHERE2 (WH_OBJECT_ID, OBJECT_ID, WHWHERE_OBJTYPE_ID, WHWHERE_OBJ_CODE, WHWHERE_OBJ_DESCR, WHWHERE_FUNCCLASS_ID, WHWHERE_FUNCTION_ID, WHWHERE_PLANTFUNC_ID, WHWHERE_PLANTFUNC_PMOD,
                                WHWHERE_EQUIP_OBJECT_ID, WHWHERE_EQUIP_CODE, WHWHERE_EQUIP_DESCR, WHWHERE_EQUIP_SERIALNO, WHWHERE_ASSEMBLY_OBJECT_ID, WHWHERE_ASSEMBLY_OBJ_CODE, WHWHERE_ASSEMBLY_OBJ_DESCR, WHWHERE_ASSEMBLY_SERIALNO,
                                WHWHERE_ASM_SUBSYS_OBJ_ID, WHWHERE_ASM_SUBSYS_OBJ_CODE, WHWHERE_ASM_SUBSYS_OBJ_DESCR, WHWHERE_STORE_OBJECT_ID, WHWHERE_STORE_OBJ_CODE, WHWHERE_STORE_OBJ_DESCR, WHWHERE_BIN_OBJECT_ID, WHWHERE_BIN_OBJ_CODE, WHWHERE_BIN_OBJ_DESCR,
                                WHWHERE_STOREAREA_ID, WHWHERE_STOREAREA_DESCR, WHWHERE_BIN_LOCATION, WHWHERE_MAIN, WHWHERE_OK, WHWHERE_COMMENTS, WHWHERE_EV_OBJECT_ID, WHWHERE_COND_ID_BEFORE, WHWHERE_COND_DESCR_BEFORE, WHWHERE_COND_ID_AFTER, WHWHERE_COND_DESCR_AFTER,
                                WHWHERESTATUS_ID, WHWHERESTATUS_DESCR, WHWHERE_WO_OBJECT_ID, WHWHERE_WO_OBJ_CODE, WHWHERE_WO_OBJ_DESCRIPTION, WHWHERE_SORTORDER, REC_DELETED, REC_CREATOR, REC_REVISOR)
            SELECT s.WH_OBJECT_ID, @D_obj, s.WHWHERE_OBJTYPE_ID,
                   CASE WHEN s.WHWHERE_OBJ_CODE IS NULL THEN NULL ELSE @To_Function_Id END,
                   CASE WHEN s.WHWHERE_OBJ_DESCR IS NULL THEN NULL ELSE @To_Function_Name END,
                   s.WHWHERE_FUNCCLASS_ID,
                   CASE WHEN s.WHWHERE_FUNCTION_ID IS NULL THEN NULL ELSE @To_Function_Id END,
                   CASE WHEN s.WHWHERE_PLANTFUNC_ID IS NULL THEN NULL ELSE pf.PlantFunc_Id END,
                   CASE WHEN s.WHWHERE_PLANTFUNC_PMOD IS NULL THEN NULL ELSE pf.PLANTFUNC_PMOD END,
                   s.WHWHERE_EQUIP_OBJECT_ID, s.WHWHERE_EQUIP_CODE, s.WHWHERE_EQUIP_DESCR, s.WHWHERE_EQUIP_SERIALNO, s.WHWHERE_ASSEMBLY_OBJECT_ID, s.WHWHERE_ASSEMBLY_OBJ_CODE, s.WHWHERE_ASSEMBLY_OBJ_DESCR, s.WHWHERE_ASSEMBLY_SERIALNO,
                   s.WHWHERE_ASM_SUBSYS_OBJ_ID, s.WHWHERE_ASM_SUBSYS_OBJ_CODE, s.WHWHERE_ASM_SUBSYS_OBJ_DESCR, s.WHWHERE_STORE_OBJECT_ID, s.WHWHERE_STORE_OBJ_CODE, s.WHWHERE_STORE_OBJ_DESCR, s.WHWHERE_BIN_OBJECT_ID, s.WHWHERE_BIN_OBJ_CODE, s.WHWHERE_BIN_OBJ_DESCR,
                   s.WHWHERE_STOREAREA_ID, s.WHWHERE_STOREAREA_DESCR, s.WHWHERE_BIN_LOCATION, s.WHWHERE_MAIN, s.WHWHERE_OK, s.WHWHERE_COMMENTS, s.WHWHERE_EV_OBJECT_ID, s.WHWHERE_COND_ID_BEFORE, s.WHWHERE_COND_DESCR_BEFORE, s.WHWHERE_COND_ID_AFTER, s.WHWHERE_COND_DESCR_AFTER,
                   s.WHWHERESTATUS_ID, s.WHWHERESTATUS_DESCR, s.WHWHERE_WO_OBJECT_ID, s.WHWHERE_WO_OBJ_CODE, s.WHWHERE_WO_OBJ_DESCRIPTION, s.WHWHERE_SORTORDER, 0, @u, @u
              FROM ST_WHWHERE2 s
              CROSS JOIN (SELECT PlantFunc_Id, PLANTFUNC_PMOD FROM ST_PlantFunction WHERE PLANTFUNC_OBJECT_ID = @D_obj) pf
             WHERE s.OBJECT_ID = @S_obj AND s.REC_DELETED = 0
               AND NOT EXISTS (SELECT 1 FROM ST_WHWHERE2 d WHERE d.WH_OBJECT_ID = s.WH_OBJECT_ID AND d.OBJECT_ID = @D_obj);
            SET @n2 = @@ROWCOUNT;

            UPDATE ST_WHWHERE2 SET REC_DELETED = -1, REC_REVISOR = @u WHERE OBJECT_ID = @S_obj AND REC_DELETED = 0;
            INSERT @log VALUES ('ST_WHWHERE2', 'movidas', @@ROWCOUNT, CAST(@n2 AS varchar(10)) + ' inseridas, ' + CAST(@n AS varchar(10)) + ' merge');

            UPDATE wh SET WH_Function_Id = @To_Function_Id, WH_Function_Name = @To_Function_Name,
                          PLANTFUNC_ID = CASE WHEN wh.PLANTFUNC_ID IS NULL THEN NULL ELSE pf.PlantFunc_Id END,
                          PLANTFUNC_LOCATION = CASE WHEN wh.PLANTFUNC_LOCATION IS NULL THEN NULL ELSE pf.PlantFunc_Location END,
                          Rec_Revisor = @u
              FROM ST_WorkHistory wh
              CROSS JOIN (SELECT PlantFunc_Id, PlantFunc_Location FROM ST_PlantFunction WHERE PLANTFUNC_OBJECT_ID = @D_obj) pf
             WHERE wh.Plant_Id = @Plant_Id AND wh.WH_FuncClass_Id = @FC AND wh.WH_Function_Id = @From_Function_Id AND wh.Rec_Deleted = 0;
            INSERT @log VALUES ('ST_WorkHistory', 'atualizadas', @@ROWCOUNT, 'WH_Function_Id/Name (TA principal do historico)');

            IF EXISTS (SELECT 1 FROM ST_WHWhere w JOIN ST_WorkHistory wh ON wh.WH_Object_Id = w.WH_Object_Id WHERE wh.Plant_Id = @Plant_Id AND w.WHW_FuncClass_Id = @FC AND w.WHW_Function_Id = @From_Function_Id AND w.Rec_Deleted = 0)
            BEGIN
                UPDATE d SET Rec_Deleted = 0, WHWhere_Ok = s.WHWhere_Ok, WHWhere_Comments = s.WHWhere_Comments, WHWhere_Ev_Object_Id = s.WHWhere_Ev_Object_Id, Rec_Revisor = @u
                  FROM ST_WHWhere d JOIN ST_WHWhere s ON s.WH_Object_Id = d.WH_Object_Id AND s.WHW_FuncClass_Id = d.WHW_FuncClass_Id
                  JOIN ST_WorkHistory wh ON wh.WH_Object_Id = s.WH_Object_Id AND wh.Plant_Id = @Plant_Id
                 WHERE d.WHW_Function_Id = @To_Function_Id AND s.WHW_Function_Id = @From_Function_Id AND s.Rec_Deleted = 0;
                INSERT ST_WHWhere (WH_Object_Id, WHW_FuncClass_Id, WHW_Function_Id, WHWhere_Ok, WHWhere_Comments, WHWhere_Ev_Object_Id, Rec_Deleted, Rec_Creator, Rec_Revisor)
                SELECT s.WH_Object_Id, s.WHW_FuncClass_Id, @To_Function_Id, s.WHWhere_Ok, s.WHWhere_Comments, s.WHWhere_Ev_Object_Id, 0, @u, @u
                  FROM ST_WHWhere s JOIN ST_WorkHistory wh ON wh.WH_Object_Id = s.WH_Object_Id AND wh.Plant_Id = @Plant_Id
                 WHERE s.WHW_FuncClass_Id = @FC AND s.WHW_Function_Id = @From_Function_Id AND s.Rec_Deleted = 0
                   AND NOT EXISTS (SELECT 1 FROM ST_WHWhere d WHERE d.WH_Object_Id = s.WH_Object_Id AND d.WHW_FuncClass_Id = s.WHW_FuncClass_Id AND d.WHW_Function_Id = @To_Function_Id);
                UPDATE s SET Rec_Deleted = -1, Rec_Revisor = @u FROM ST_WHWhere s JOIN ST_WorkHistory wh ON wh.WH_Object_Id = s.WH_Object_Id AND wh.Plant_Id = @Plant_Id
                 WHERE s.WHW_FuncClass_Id = @FC AND s.WHW_Function_Id = @From_Function_Id AND s.Rec_Deleted = 0;
                INSERT @log VALUES ('ST_WHWhere (legada)', 'movidas', @@ROWCOUNT, 'tabela antiga ainda tinha linhas; movidas tambem');
            END
        END
        ELSE
            INSERT @log VALUES ('ST_WHWHERE2 / ST_WorkHistory', 'pulado', 0, '@MoveWorkHistory = 0');

        /* ======================= 3d. ST_POAITEM2 (MR/PO lines) ======================= */
        IF @MovePOItems = 1
        BEGIN
            UPDATE p SET FUNCTION_ID = @To_Function_Id, REC_REVISOR = @u
              FROM ST_POAITEM2 p JOIN ST_Object o ON o.Object_Id = p.POITEM_OBJECT_ID
             WHERE o.ObjClass_Id = @Plant_Id AND p.FUNCCLASS_ID = @FC AND p.FUNCTION_ID = @From_Function_Id AND p.REC_DELETED = 0;
            INSERT @log VALUES ('ST_POAITEM2', 'atualizadas', @@ROWCOUNT, 'linhas de MR/PO da planta');
        END
        ELSE INSERT @log VALUES ('ST_POAITEM2', 'pulado', 0, '@MovePOItems = 0');

        /* ======================= 3e. Materials ======================= */
        IF @MoveMaterialRefs = 1
        BEGIN
            UPDATE ST_MATONHANDSTAT SET MatOH_Def_Function_Id = @To_Function_Id, REC_REVISOR = @u
             WHERE PLANT_ID = @Plant_Id AND MatOH_Def_FuncClass_Id = @FC AND MatOH_Def_Function_Id = @From_Function_Id;
            INSERT @log VALUES ('ST_MATONHANDSTAT', 'atualizadas', @@ROWCOUNT, 'TA default do material na planta');

            UPDATE ST_MATTRANS2 SET MATTR_FUNCTION_ID = @To_Function_Id, MATTR_FUNCTION_NAME = @To_Function_Name, REC_REVISOR = @u
             WHERE MATTR_FUNCCLASS_ID = @FC AND MATTR_FUNCTION_ID = @From_Function_Id
               AND ISNULL(MATTR_OWNER_PLANT_ID, MATTR_STORE_PLANT_ID) = @Plant_Id;
            INSERT @log VALUES ('ST_MATTRANS2', 'atualizadas', @@ROWCOUNT, 'historico de transacoes de material');
        END
        ELSE INSERT @log VALUES ('ST_MATONHANDSTAT / ST_MATTRANS2', 'pulado', 0, '@MoveMaterialRefs = 0');

        /* ======================= 3f. Other references ======================= */
        IF @MoveOtherRefs = 1
        BEGIN
            UPDATE b SET Function_Id = @To_Function_Id, Rec_Revisor = @u
              FROM ST_BudItem b JOIN ST_Object o ON o.Object_Id = b.Bud_Object_Id
             WHERE o.ObjClass_Id = @Plant_Id AND b.FuncClass_Id = @FC AND b.Function_Id = @From_Function_Id
               AND NOT EXISTS (SELECT 1 FROM ST_BudItem d WHERE d.Bud_Object_Id = b.Bud_Object_Id AND d.FuncClass_Id = b.FuncClass_Id AND d.Function_Id = @To_Function_Id AND d.Account_Id = b.Account_Id);
            SET @n = @@ROWCOUNT;
            SELECT @n2 = COUNT(*) FROM ST_BudItem b JOIN ST_Object o ON o.Object_Id = b.Bud_Object_Id WHERE o.ObjClass_Id = @Plant_Id AND b.FuncClass_Id = @FC AND b.Function_Id = @From_Function_Id;
            INSERT @log VALUES ('ST_BudItem', CASE WHEN @n2 > 0 THEN 'ATENCAO' ELSE 'atualizadas' END, @n, CASE WHEN @n2 > 0 THEN CAST(@n2 AS varchar(10)) + ' itens de orcamento NAO movidos (ja existe item igual no destino) - ajustar manualmente' ELSE 'itens de orcamento' END);

            UPDATE p SET Function_Id = @To_Function_Id, Rec_Revisor = @u
              FROM ST_PlantCost p
             WHERE p.Plant_Id = @Plant_Id AND p.FuncClass_Id = @FC AND p.Function_Id = @From_Function_Id;
            INSERT @log VALUES ('ST_PlantCost', 'atualizadas', @@ROWCOUNT, 'custos por TA');

            UPDATE r SET Function_Id = @To_Function_Id, Rec_Revisor = @u
              FROM ST_UserGroupRestrictFunc r
             WHERE r.Plant_Id = @Plant_Id AND r.FuncClass_Id = @FC AND r.Function_Id = @From_Function_Id
               AND NOT EXISTS (SELECT 1 FROM ST_UserGroupRestrictFunc d WHERE d.UserGroup_Id = r.UserGroup_Id AND d.Plant_Id = r.Plant_Id AND d.FuncClass_Id = r.FuncClass_Id AND d.Function_Id = @To_Function_Id);
            INSERT @log VALUES ('ST_UserGroupRestrictFunc', 'atualizadas', @@ROWCOUNT, 'restricoes de grupo de usuario por TA');
        END

        /* ======================= 3g. References that are only reported ======================= */
        INSERT @log (Tabela, Acao, Linhas, Detalhe)
        SELECT x.Tabela, 'VERIFICAR', x.n, 'referencias a ' + @From_Function_Id + ' nesta tabela nao sao movidas pela tool (sem contexto de planta seguro) - revisar manualmente'
          FROM (
                SELECT 'ST_NonConformity' Tabela, COUNT(*) n FROM ST_NonConformity WHERE NC_Function_Id = @From_Function_Id
                UNION ALL SELECT 'ST_GuaranteeClaim', COUNT(*) FROM ST_GuaranteeClaim WHERE GC_Function_Id = @From_Function_Id
                UNION ALL SELECT 'ST_INVAHead', COUNT(*) FROM ST_INVAHead WHERE Inv_Function_Id = @From_Function_Id
                UNION ALL SELECT 'ST_EventConsequence', COUNT(*) FROM ST_EventConsequence WHERE EvCons_Function_Id = @From_Function_Id
                UNION ALL SELECT 'ST_EventLessonLearned', COUNT(*) FROM ST_EventLessonLearned WHERE Function_Id = @From_Function_Id AND Plant_Id = @Plant_Id
                UNION ALL SELECT 'ST_WorkPlanProject', COUNT(*) FROM ST_WorkPlanProject WHERE Function_Id = @From_Function_Id
                UNION ALL SELECT 'ST_WorkHistoryProject', COUNT(*) FROM ST_WorkHistoryProject WHERE WHProject_Function_Id = @From_Function_Id
                UNION ALL SELECT 'ST_ASSEMBLYNODE', COUNT(*) FROM ST_ASSEMBLYNODE WHERE INSTALLED_FUNCTION_ID = @From_Function_Id
                UNION ALL SELECT 'ST_SELECTGROUP', COUNT(*) FROM ST_SELECTGROUP WHERE PLANT_ID = @Plant_Id AND FUNCTION_ID = @From_Function_Id
               ) x
         WHERE x.n > 0;

        /* ======================= 4. Retire the source TA (usp_IPS_RetireTA) ======================= */
        IF @RetireSource = 1
        BEGIN
            DECLARE @ret TABLE (Tabela sysname, Acao varchar(30), Linhas int, Detalhe varchar(400));
            INSERT @ret EXEC dbo.usp_IPS_RetireTA @Plant_Id = @Plant_Id, @Function_Id = @From_Function_Id, @RetireInClass = @RetireSourceInClass, @Rec_User = @u;
            INSERT @log (Tabela, Acao, Linhas, Detalhe) SELECT Tabela, Acao, Linhas, Detalhe FROM @ret;
        END

        INSERT @log VALUES ('-', 'fim', NULL, CASE WHEN @DryRun = 1 THEN 'SIMULACAO: nada foi gravado' ELSE 'concluido' END);

        IF @DryRun = 1
        BEGIN
            IF @outer = 1 ROLLBACK TRANSACTION MoveTA; ELSE ROLLBACK TRANSACTION;
        END
        ELSE IF @outer = 0
            COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @outer = 0 AND @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH

    SELECT Seq, Tabela, Acao, Linhas, Detalhe FROM @log ORDER BY Seq;
END
GO

/* -----------------------------------------------------------------------------
   usp_IPS_CreateTA - creates (or reactivates) a TA in the plant, inside @Father.
   ST_Function (class) + ST_Object type 2 (st_CreateObjectId) + ST_PlantFunction.
   If the code already exists active in the class, it is only activated in the
   plant (the class name is kept).
   Joins the caller's transaction; with no open transaction, uses its own.
   ----------------------------------------------------------------------------- */
CREATE OR ALTER PROCEDURE dbo.usp_IPS_CreateTA
    @Plant_Id       varchar(20),
    @Function_Id    varchar(100),
    @Function_Name  varchar(100),
    @Father         varchar(100),
    @Rec_User       varchar(100) = NULL
AS
BEGIN
    -- MoveTA-Procs-Version: 2.3   (Move-TA.ps1 compares this marker with the bundled file and updates the procedures; bump it on every procedure change)
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @log TABLE (Seq int IDENTITY(1,1), Tabela sysname, Acao varchar(30), Linhas int, Detalhe varchar(400));
    DECLARE @FC varchar(10), @u varchar(100), @obj numeric(15,0), @msg varchar(400), @outer bit, @code varchar(100), @nm varchar(100);
    SET @u = ISNULL(NULLIF(LTRIM(RTRIM(@Rec_User)),''), UPPER(SUSER_SNAME()));
    SET @Function_Id = LTRIM(RTRIM(@Function_Id)); SET @Function_Name = LTRIM(RTRIM(@Function_Name)); SET @Father = LTRIM(RTRIM(@Father));
    IF ISNULL(@Function_Id,'') = '' OR ISNULL(@Function_Name,'') = '' THROW 50101, 'Informe codigo e nome da TA.', 1;
    SELECT @FC = FuncClass_Id FROM ST_Plant WHERE Plant_Id = @Plant_Id AND Rec_Deleted = 0;
    IF @FC IS NULL BEGIN SET @msg = 'Planta ' + @Plant_Id + ' nao encontrada.'; THROW 50102, @msg, 1; END
    IF NOT EXISTS (SELECT 1 FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Father AND Rec_Deleted = 0)
    BEGIN SET @msg = 'TA pai ' + @Father + ' nao esta ativa na planta ' + @Plant_Id + '.'; THROW 50103, @msg, 1; END
    IF EXISTS (SELECT 1 FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Function_Id AND Rec_Deleted = 0)
    BEGIN SET @msg = 'TA ' + @Function_Id + ' ja existe na planta ' + @Plant_Id + '.'; THROW 50104, @msg, 1; END

    SET @outer = CASE WHEN @@TRANCOUNT > 0 THEN 1 ELSE 0 END;
    IF @outer = 0 BEGIN TRANSACTION;
    BEGIN TRY
        IF EXISTS (SELECT 1 FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id AND Rec_Deleted = 0)
            INSERT @log VALUES ('ST_Function', 'existe', 1, @Function_Id + ' ja existe na classe ' + @FC + '; nome da classe mantido');
        ELSE IF EXISTS (SELECT 1 FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id)
        BEGIN
            UPDATE ST_Function SET Rec_Deleted = 0, Function_Name = @Function_Name, Function_Father = @Father, Function_Length = LEN(@Function_Id), Rec_Revisor = @u
             WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id;
            INSERT @log VALUES ('ST_Function', 'reativada', 1, @Function_Id + ' (' + @Function_Name + ')  pai=' + @Father);
        END
        ELSE
        BEGIN
            INSERT ST_Function (FuncClass_Id, Function_Id, Function_Father, Function_Name, Function_Length, Rec_Deleted, Rec_Creator, Rec_Revisor)
            VALUES (@FC, @Function_Id, @Father, @Function_Name, LEN(@Function_Id), 0, @u, @u);
            INSERT @log VALUES ('ST_Function', 'criada', 1, @Function_Id + ' (' + @Function_Name + ')  pai=' + @Father);
        END
        UPDATE f SET Function_OnLower = (SELECT COUNT(*) FROM ST_Function c WHERE c.FuncClass_Id = f.FuncClass_Id AND c.Function_Father = f.Function_Id AND c.Rec_Deleted = 0), Rec_Revisor = @u
          FROM ST_Function f WHERE f.FuncClass_Id = @FC AND f.Function_Id = @Father;
        SELECT @nm = Function_Name FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id;

        IF EXISTS (SELECT 1 FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Function_Id)
        BEGIN
            SELECT @obj = PLANTFUNC_OBJECT_ID FROM ST_PlantFunction WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Function_Id;
            IF @obj IS NULL
            BEGIN
                SET @code = @Function_Id;
                EXEC dbo.st_CreateObjectId @obj OUTPUT, @Plant_Id, 2, @code OUTPUT, @nm, 0, 0, @u;
                UPDATE ST_PlantFunction SET PLANTFUNC_OBJECT_ID = @obj WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Function_Id;
            END
            UPDATE ST_PlantFunction SET Rec_Deleted = 0, Rec_Revisor = @u WHERE Plant_Id = @Plant_Id AND FuncClass_Id = @FC AND Function_Id = @Function_Id;
            UPDATE ST_Object SET Rec_Deleted = 0, Obj_Description = @nm, Rec_Revisor = @u WHERE Object_Id = @obj;
            INSERT @log VALUES ('ST_PlantFunction', 'reativada', 1, @Function_Id + ' na planta ' + @Plant_Id + '  objeto=' + CAST(@obj AS varchar(20)));
        END
        ELSE
        BEGIN
            SET @code = @Function_Id;
            EXEC dbo.st_CreateObjectId @obj OUTPUT, @Plant_Id, 2, @code OUTPUT, @nm, 0, 0, @u;
            IF ISNULL(@obj, 0) = 0 THROW 50105, 'st_CreateObjectId nao devolveu Object_Id.', 1;
            INSERT ST_PlantFunction (Plant_Id, FuncClass_Id, Function_Id, PlantFunc_IsEquipment, PlantFunc_IsStore, PLANTFUNC_OBJECT_ID, Rec_Deleted, Rec_Creator, Rec_Revisor)
            VALUES (@Plant_Id, @FC, @Function_Id, 0, 0, @obj, 0, @u, @u);
            INSERT @log VALUES ('ST_PlantFunction', 'criada', 1, @Function_Id + ' na planta ' + @Plant_Id + '  objeto=' + CAST(@obj AS varchar(20)));
        END
        IF @outer = 0 COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @outer = 0 AND @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
    SELECT Tabela, Acao, Linhas, Detalhe FROM @log ORDER BY Seq;
END
GO

/* -----------------------------------------------------------------------------
   usp_IPS_RenameTA - renames a TA. The name lives in ST_Function (valid for the
   whole class) and in the TA object description (ST_Object) of every plant of
   the class. Archived history (ST_WorkHistory.WH_Function_Name) keeps the old name.
   ----------------------------------------------------------------------------- */
CREATE OR ALTER PROCEDURE dbo.usp_IPS_RenameTA
    @Plant_Id     varchar(20),
    @Function_Id  varchar(100),
    @NewName      varchar(100),
    @Rec_User     varchar(100) = NULL
AS
BEGIN
    -- MoveTA-Procs-Version: 2.3   (Move-TA.ps1 compares this marker with the bundled file and updates the procedures; bump it on every procedure change)
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @log TABLE (Seq int IDENTITY(1,1), Tabela sysname, Acao varchar(30), Linhas int, Detalhe varchar(400));
    DECLARE @FC varchar(10), @u varchar(100), @old varchar(100), @msg varchar(400), @outer bit;
    SET @u = ISNULL(NULLIF(LTRIM(RTRIM(@Rec_User)),''), UPPER(SUSER_SNAME()));
    SET @Function_Id = LTRIM(RTRIM(@Function_Id)); SET @NewName = LTRIM(RTRIM(@NewName));
    IF ISNULL(@NewName,'') = '' THROW 50111, 'Informe o novo nome.', 1;
    SELECT @FC = FuncClass_Id FROM ST_Plant WHERE Plant_Id = @Plant_Id AND Rec_Deleted = 0;
    IF @FC IS NULL BEGIN SET @msg = 'Planta ' + @Plant_Id + ' nao encontrada.'; THROW 50112, @msg, 1; END
    SELECT @old = Function_Name FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id AND Rec_Deleted = 0;
    IF @old IS NULL BEGIN SET @msg = 'TA ' + @Function_Id + ' nao existe (ativa) na classe ' + @FC + '.'; THROW 50113, @msg, 1; END

    SET @outer = CASE WHEN @@TRANCOUNT > 0 THEN 1 ELSE 0 END;
    IF @outer = 0 BEGIN TRANSACTION;
    BEGIN TRY
        UPDATE ST_Function SET Function_Name = @NewName, Rec_Revisor = @u WHERE FuncClass_Id = @FC AND Function_Id = @Function_Id;
        INSERT @log VALUES ('ST_Function', 'renomeada', @@ROWCOUNT, @Function_Id + ': "' + ISNULL(@old,'') + '" -> "' + @NewName + '" (classe ' + @FC + ')');
        UPDATE o SET Obj_Description = @NewName, Rec_Revisor = @u
          FROM ST_Object o JOIN ST_PlantFunction pf ON pf.PLANTFUNC_OBJECT_ID = o.Object_Id
         WHERE pf.FuncClass_Id = @FC AND pf.Function_Id = @Function_Id;
        INSERT @log VALUES ('ST_Object', 'renomeada', @@ROWCOUNT, 'descricao do objeto da TA em cada planta da classe');
        IF @outer = 0 COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @outer = 0 AND @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
    SELECT Tabela, Acao, Linhas, Detalhe FROM @log ORDER BY Seq;
END
GO