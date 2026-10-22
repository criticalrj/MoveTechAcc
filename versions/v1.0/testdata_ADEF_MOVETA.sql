/* Dados de teste para a tool Move TA (banco ADEF_MOVETA, planta FMHUB / classe FMHUB).
   Cria a arvore 7/701/701.01(.01/.02) e 8/801/801.01/801.04/801.10 (+ subaccounts),
   jobs (320) ligados por ST_PlantFunctionObject + ST_WPWHERE2 e
   historicos (WO 340 + ST_WorkHistory) ligados por ST_WHWHERE2.
   Tudo marcado com Rec_Creator = 'MOVETA.TEST' e descricao 'TESTE MoveTA'. */
SET NOCOUNT ON;
SET XACT_ABORT ON;
DECLARE @P varchar(20) = 'FMHUB', @FC varchar(10) = 'FMHUB', @u varchar(100) = 'MOVETA.TEST';
IF EXISTS (SELECT 1 FROM ST_PlantFunction WHERE Plant_Id = @P) THROW 50000, 'FMHUB ja tem TAs; script de teste e so para banco vazio.', 1;

DECLARE @ta TABLE (seq int IDENTITY, id varchar(100), father varchar(100), name varchar(100));
INSERT @ta (id, father, name) VALUES
 ('7','*','SYSTEMS FOR MACHINERY MAIN COMPONENTS'),
 ('701','7','Fuel Oil Transfer And Drain System'),
 ('701.01','701','Fuel oil transfer pumps'),
 ('701.01.01','701.01','M.D.O. Transfer Pump No.1 & El. Motor'),
 ('701.01.02','701.01','M.D.O. Transfer Pump No.2 & El. Motor'),
 ('8','*','SHIP SYSTEMS'),
 ('801','8','Ballast Systems'),
 ('801.01','801','Ballast Pumps'),
 ('801.01.01','801.01','Ballast Pump No.1 & El. Motor'),
 ('801.01.02','801.01','Ballast Pump No.2 & El. Motor'),
 ('801.04','801','Ballast Pipes,Valves And Filter'),
 ('801.04.01','801.04','Ballast Piping''s, Valves & Fittings - All'),
 ('801.04.02','801.04','Actuators for Ballast Tanks'),
 ('801.10','801','Ballast Water Treatment System Assembly'),
 ('801.10.01','801.10','BWTS Control Touch Panel (+CP)'),
 ('801.10.02','801.10','BWTS Filter Control Terminals Panel (+CP)'),
 ('801.10.03','801.10','BWTS Sensor Box 1 (+SB1)'),
 ('801.10.04','801.10','BWTS UV Power Control Cabinet (+UVP01)'),
 ('801.10.05','801.10','BWTS Ex Interface Panel (+EXIP)'),
 ('801.10.06','801.10','BWTS GPS Junction Box (+GPS)'),
 ('801.10.07','801.10','BWTS Interlock Panel (+EXIP)'),
 ('801.10.08','801.10','BWTS Interlock Panel (+IP)'),
 ('801.10.09','801.10','BWTS Differential Pressure (DP) Transmitter'),
 ('801.10.10','801.10','BWTS Temperature Transmitter'),
 ('801.10.11','801.10','BWTS UV Sensor'),
 ('801.10.12','801.10','BWTS Filter Back Flush Pump & El. Motor'),
 ('801.10.13','801.10','BWTS Filter'),
 ('801.10.14','801.10','BWTS Pipings, Valves & Fittings - All'),
 ('801.10.15','801.10','BWTS Flow Pressure Valve (FPV01)'),
 ('801.10.16','801.10','BWTS UV Chamber'),
 ('801.10.17','801.10','BWTS UV Bulb'),
 ('801.10.18','801.10','BWTS Filter Control Panel (+FC)'),
 ('801.10.19','801.10','BWTS UV System Assembly'),
 ('801.10.20','801.10','BWTS Flowmeter (FM01)'),
 ('801.10.21','801.10','BWTS Air Release Valve'),
 ('801.10.22','801.10','BWTS UV Assembly Nitrogen Ex p System'),
 ('801.10.23','801.10','BWTS UV Assembly Ex p System Piping, Valves & Fittings - All'),
 ('801.10.24','801.10','BWTS Butterfly Valves - All'),
 ('801.10.25','801.10','BWTS UV Assembly Nitrogen Manual Pressure Regulator'),
 ('801.10.26','801.10','BWTS By-pass Valve'),
 ('801.10.27','801.10','BWTS Exi Junction Box (+EXJB1)'),
 ('801.10.28','801.10','BWTS Terminal Box (+TB01)'),
 ('801.10.29','801.10','BWTS Filter Rotating Gear Reducer Motor Unit & El. Motor');

BEGIN TRANSACTION;

/* ---- TAs: ST_Function + objeto tipo 2 + ST_PlantFunction ---- */
INSERT ST_Function (FuncClass_Id, Function_Id, Function_Father, Function_Name, Function_Length, Rec_Deleted, Rec_Creator, Rec_Revisor)
SELECT @FC, id, father, name, LEN(id), 0, @u, @u FROM @ta ORDER BY seq;

IF NOT EXISTS (SELECT 1 FROM ST_PlantFunction WHERE Plant_Id = @P AND Function_Id = '*')
BEGIN
    DECLARE @rootObj numeric(15,0), @rootName varchar(100) = (SELECT Function_Name FROM ST_Function WHERE FuncClass_Id=@FC AND Function_Id='*'), @rootCode varchar(100) = '*';
    EXEC dbo.st_CreateObjectId @rootObj OUTPUT, @P, 2, @rootCode OUTPUT, @rootName, 0, 0, @u;
    INSERT ST_PlantFunction (Plant_Id, FuncClass_Id, Function_Id, PlantFunc_IsEquipment, PlantFunc_IsStore, PLANTFUNC_OBJECT_ID, Rec_Deleted, Rec_Creator, Rec_Revisor)
    VALUES (@P, @FC, '*', 0, 0, @rootObj, 0, @u, @u);
END

DECLARE @i int = 1, @id varchar(100), @nm varchar(100), @code varchar(100), @obj numeric(15,0);
WHILE @i <= (SELECT MAX(seq) FROM @ta)
BEGIN
    SELECT @id = id, @nm = name FROM @ta WHERE seq = @i;
    SET @code = @id; SET @obj = NULL;
    EXEC dbo.st_CreateObjectId @obj OUTPUT, @P, 2, @code OUTPUT, @nm, 0, 0, @u;
    INSERT ST_PlantFunction (Plant_Id, FuncClass_Id, Function_Id, PlantFunc_IsEquipment, PlantFunc_IsStore, PLANTFUNC_OBJECT_ID, Criticality_Id, PlantFunc_Remarks, Rec_Deleted, Rec_Creator, Rec_Revisor)
    VALUES (@P, @FC, @id, 0, 0, @obj, NULL, 'TESTE MoveTA', 0, @u, @u);
    SET @i += 1;
END
UPDATE f SET Function_OnLower = (SELECT COUNT(*) FROM ST_Function c WHERE c.FuncClass_Id = f.FuncClass_Id AND c.Function_Father = f.Function_Id AND c.Rec_Deleted = 0)
  FROM ST_Function f WHERE f.FuncClass_Id = @FC;

/* ---- Jobs (320): um por TA folha + jobs com where-list extra ---- */
DECLARE @jobs TABLE (seq int IDENTITY, main varchar(100), extra1 varchar(100), extra2 varchar(100), descr varchar(100));
INSERT @jobs (main, extra1, extra2, descr)
SELECT t.id, NULL, NULL, 'TESTE MoveTA - Inspecao ' + t.name
  FROM @ta t WHERE NOT EXISTS (SELECT 1 FROM @ta c WHERE c.father = t.id) ORDER BY t.seq;
INSERT @jobs (main, extra1, extra2, descr) VALUES
 ('801.10', '801.10.13', '801.10.16', 'TESTE MoveTA - BWTS overhaul (where-list 801.10.13 + 801.10.16)'),
 ('701.01', '801.01.01', NULL,        'TESTE MoveTA - Transfer pumps check (where-list 801.01.01)'),
 ('801.04', '701.01.02', NULL,        'TESTE MoveTA - Ballast piping (where-list 701.01.02)');

DECLARE @wp numeric(15,0), @m varchar(100), @e1 varchar(100), @e2 varchar(100), @d varchar(100), @n int = 1;
WHILE @n <= (SELECT MAX(seq) FROM @jobs)
BEGIN
    SELECT @m = main, @e1 = extra1, @e2 = extra2, @d = descr FROM @jobs WHERE seq = @n;
    SET @code = 'TST-' + RIGHT('000' + CAST(@n AS varchar(3)), 3); SET @wp = NULL;
    EXEC dbo.st_CreateObjectId @wp OUTPUT, @P, 320, @code OUTPUT, @d, -1, 0, @u;   -- -1: cria ST_WorkPlan
    UPDATE ST_WorkPlan SET WP_Corrective = 0, WP_FreeText = @d WHERE WP_Object_Id = @wp;
    INSERT ST_WPDUE (WP_OBJECT_ID, WP_NEXTDUEDATE, WP_LASTDONEDATE, REC_DELETED, REC_CREATOR, REC_REVISOR)
    VALUES (@wp, DATEADD(day, 30 * @n, '2026-11-01'), '2026-05-01', 0, @u, @u);
    INSERT ST_PlantFunctionObject (Plant_Id, FuncClass_Id, Function_Id, Object_Id, ObjType_Id, Rec_Deleted, Rec_Creator, Rec_Revisor)
    VALUES (@P, @FC, @m, @wp, 320, 0, @u, @u);
    INSERT ST_WPWHERE2 (WP_OBJECT_ID, OBJECT_ID, WPWHERE_MAIN, WPWHERE_SORTORDER, REC_DELETED, REC_CREATOR, REC_REVISOR)
    SELECT @wp, PLANTFUNC_OBJECT_ID, -1, 1, 0, @u, @u FROM ST_PlantFunction WHERE Plant_Id = @P AND Function_Id = @m;
    IF @e1 IS NOT NULL
        INSERT ST_WPWHERE2 (WP_OBJECT_ID, OBJECT_ID, WPWHERE_MAIN, WPWHERE_SORTORDER, WPWHERE_COMMENTS, REC_DELETED, REC_CREATOR, REC_REVISOR)
        SELECT @wp, PLANTFUNC_OBJECT_ID, 0, 2, 'where-list extra', 0, @u, @u FROM ST_PlantFunction WHERE Plant_Id = @P AND Function_Id = @e1;
    IF @e2 IS NOT NULL
        INSERT ST_WPWHERE2 (WP_OBJECT_ID, OBJECT_ID, WPWHERE_MAIN, WPWHERE_SORTORDER, WPWHERE_COMMENTS, REC_DELETED, REC_CREATOR, REC_REVISOR)
        SELECT @wp, PLANTFUNC_OBJECT_ID, 0, 3, 'where-list extra', 0, @u, @u FROM ST_PlantFunction WHERE Plant_Id = @P AND Function_Id = @e2;
    SET @n += 1;
END

/* ---- Work history: WO 340 arquivada + ST_WorkHistory + ST_WHWHERE2 ---- */
DECLARE @whs TABLE (seq int IDENTITY, ta varchar(100), extra varchar(100));
INSERT @whs (ta, extra) VALUES
 ('801.01.01', NULL), ('801.01.02', NULL), ('801.04.01', NULL), ('801.04.02', NULL),
 ('801.10.12', NULL), ('801.10.13', '801.10.14'), ('801.10.16', NULL), ('801.10', NULL), ('801.01', NULL),
 ('701.01.01', '801.01.01');
DECLARE @wh numeric(15,0), @ta2 varchar(100), @ex varchar(100), @tn varchar(100), @job numeric(15,0), @k int = 1;
WHILE @k <= (SELECT MAX(seq) FROM @whs)
BEGIN
    SELECT @ta2 = ta, @ex = extra FROM @whs WHERE seq = @k;
    SELECT @tn = Function_Name FROM ST_Function WHERE FuncClass_Id = @FC AND Function_Id = @ta2;
    SELECT TOP 1 @job = pfo.Object_Id FROM ST_PlantFunctionObject pfo WHERE pfo.Plant_Id = @P AND pfo.Function_Id = @ta2 AND pfo.ObjType_Id = 320;
    SET @code = 'TWO-' + RIGHT('000' + CAST(@k AS varchar(3)), 3); SET @d = 'TESTE MoveTA - WO arquivada ' + @tn; SET @wh = NULL;
    EXEC dbo.st_CreateObjectId @wh OUTPUT, @P, 340, @code OUTPUT, @d, 0, 0, @u;
    INSERT ST_WorkHistory (WH_Object_Id, WH_ShortDescr, Plant_Id, WH_FuncClass_Id, WH_Function_Id, WH_Function_Name, WH_Obj_Code, WP_SH_Object_Id, WP_Corrective,
                           WP_ActStartDate, WP_ActFinishDate, WH_FILEDBY, WH_FILEDDATE, Rec_Deleted, Rec_Creator, Rec_Revisor)
    VALUES (@wh, @d, @P, @FC, @ta2, @tn, @code, @job, 0, DATEADD(day, -10 * @k, '2026-10-01'), DATEADD(day, -10 * @k + 1, '2026-10-01'), @u, '2026-10-01', 0, @u, @u);
    INSERT ST_WHWHERE2 (WH_OBJECT_ID, OBJECT_ID, WHWHERE_MAIN, WHWHERE_OK, WHWHERE_SORTORDER, REC_DELETED, REC_CREATOR, REC_REVISOR)
    SELECT @wh, PLANTFUNC_OBJECT_ID, -1, -1, 1, 0, @u, @u FROM ST_PlantFunction WHERE Plant_Id = @P AND Function_Id = @ta2;
    IF @ex IS NOT NULL
        INSERT ST_WHWHERE2 (WH_OBJECT_ID, OBJECT_ID, WHWHERE_MAIN, WHWHERE_OK, WHWHERE_SORTORDER, WHWHERE_COMMENTS, REC_DELETED, REC_CREATOR, REC_REVISOR)
        SELECT @wh, PLANTFUNC_OBJECT_ID, 0, -1, 2, 'where-list extra', 0, @u, @u FROM ST_PlantFunction WHERE Plant_Id = @P AND Function_Id = @ex;
    SET @k += 1;
END

COMMIT TRANSACTION;

SELECT (SELECT COUNT(*) FROM ST_PlantFunction WHERE Plant_Id=@P AND Rec_Deleted=0) AS tas,
       (SELECT COUNT(*) FROM ST_PlantFunctionObject WHERE Plant_Id=@P AND Rec_Deleted=0) AS pfo,
       (SELECT COUNT(*) FROM ST_WPWHERE2 WHERE REC_DELETED=0) AS wpwhere2,
       (SELECT COUNT(*) FROM ST_WorkHistory WHERE Plant_Id=@P AND Rec_Deleted=0) AS wh,
       (SELECT COUNT(*) FROM ST_WHWHERE2 WHERE REC_DELETED=0) AS whwhere2;
