SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.ErrorAnalysisLookups', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ErrorAnalysisLookups
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        GroupCode NVARCHAR(100) NOT NULL,
        [Value] NVARCHAR(200) NOT NULL,
        DisplayText NVARCHAR(200) NOT NULL,
        SortOrder INT NOT NULL,
        IsActive BIT NOT NULL CONSTRAINT DF_ErrorAnalysisLookups_IsActive DEFAULT (1)
    );
END;

IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartSupplier1') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartSupplier1 NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartCav1') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartCav1 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartLot1') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartLot1 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartSupplier2') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartSupplier2 NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartCav2') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartCav2 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartLot2') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartLot2 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartSupplier3') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartSupplier3 NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartCav3') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartCav3 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartLot3') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartLot3 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartCode4') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartCode4 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartName4') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartName4 NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartSupplier4') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartSupplier4 NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartCav4') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartCav4 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartLot4') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartLot4 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartCode5') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartCode5 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartName5') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartName5 NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartSupplier5') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartSupplier5 NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartCav5') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartCav5 NVARCHAR(100) NULL;
IF COL_LENGTH(N'dbo.ErrorAnalyses', N'PartLot5') IS NULL ALTER TABLE dbo.ErrorAnalyses ADD PartLot5 NVARCHAR(100) NULL;

IF EXISTS (
    SELECT 1
    FROM sys.columns c
    INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
    WHERE c.object_id = OBJECT_ID(N'dbo.ErrorAnalyses')
      AND c.name = N'TemporaryMeasureDate'
      AND t.name = N'date'
)
BEGIN
    ALTER TABLE dbo.ErrorAnalyses ALTER COLUMN TemporaryMeasureDate DATETIME NULL;
END;

IF EXISTS (
    SELECT 1
    FROM sys.columns c
    INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
    WHERE c.object_id = OBJECT_ID(N'dbo.ErrorAnalyses')
      AND c.name = N'PermanentMeasureDate'
      AND t.name = N'date'
)
BEGIN
    ALTER TABLE dbo.ErrorAnalyses ALTER COLUMN PermanentMeasureDate DATETIME NULL;
END;

IF OBJECT_ID(N'dbo.ErrorAnalyses', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ErrorAnalyses
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Phenomenon NVARCHAR(200) NULL,
        ProblemContent NVARCHAR(MAX) NULL,
        Category NVARCHAR(200) NULL,
        Model NVARCHAR(200) NULL,
        StageClassification NVARCHAR(200) NULL,
        OccurrenceProcess NVARCHAR(200) NULL,
        Line NVARCHAR(100) NULL,
        OccurrenceCount INT NULL,
        OccurrenceDate DATETIME NULL,
        ShipmentStop NVARCHAR(200) NULL,
        Investigator NVARCHAR(200) NULL,
        BAction NVARCHAR(100) NULL,
        ChronicDefect NVARCHAR(100) NULL,
        LineStopTime NVARCHAR(200) NULL,
        InvestigationContent NVARCHAR(MAX) NULL,
        PartCode1 NVARCHAR(100) NULL,
        PartName1 NVARCHAR(200) NULL,
        PartSupplier1 NVARCHAR(200) NULL,
        PartCav1 NVARCHAR(100) NULL,
        PartLot1 NVARCHAR(100) NULL,
        PartCode2 NVARCHAR(100) NULL,
        PartName2 NVARCHAR(200) NULL,
        PartSupplier2 NVARCHAR(200) NULL,
        PartCav2 NVARCHAR(100) NULL,
        PartLot2 NVARCHAR(100) NULL,
        PartCode3 NVARCHAR(100) NULL,
        PartName3 NVARCHAR(200) NULL,
        PartSupplier3 NVARCHAR(200) NULL,
        PartCav3 NVARCHAR(100) NULL,
        PartLot3 NVARCHAR(100) NULL,
        PartCode4 NVARCHAR(100) NULL,
        PartName4 NVARCHAR(200) NULL,
        PartSupplier4 NVARCHAR(200) NULL,
        PartCav4 NVARCHAR(100) NULL,
        PartLot4 NVARCHAR(100) NULL,
        PartCode5 NVARCHAR(100) NULL,
        PartName5 NVARCHAR(200) NULL,
        PartSupplier5 NVARCHAR(200) NULL,
        PartCav5 NVARCHAR(100) NULL,
        PartLot5 NVARCHAR(100) NULL,
        InvestigationProgress NVARCHAR(200) NULL,
        CauseClassification NVARCHAR(200) NULL,
        DetailedCause NVARCHAR(MAX) NULL,
        InvestigationStartTime DATETIME NULL,
        InvestigationEndTime DATETIME NULL,
        DaysOff INT NULL,
        NightShiftDays INT NULL,
        InvestigationHours FLOAT NULL,
        AnalysisContent NVARCHAR(MAX) NULL,
        NextAction NVARCHAR(MAX) NULL,
        TemporaryMeasureClassification NVARCHAR(200) NULL,
        TemporaryMeasureDate DATETIME NULL,
        TemporaryMeasureDetail NVARCHAR(MAX) NULL,
        PermanentMeasureClassification NVARCHAR(200) NULL,
        PermanentMeasureDate DATETIME NULL,
        PermanentMeasureDetail NVARCHAR(MAX) NULL,
        AttachmentPath NVARCHAR(500) NULL,
        Factory NVARCHAR(100) NULL,
        Status NVARCHAR(100) NULL,
        OccurrencePeriod NVARCHAR(100) NULL,
        Team NVARCHAR(200) NULL,
        PersonInCharge NVARCHAR(200) NULL,
        ImageStatus NVARCHAR(500) NULL,
        CauseDescription NVARCHAR(MAX) NULL,
        NextActionSummary NVARCHAR(MAX) NULL,
        DefectClassification NVARCHAR(500) NULL,
        TemporaryMeasureSummary NVARCHAR(MAX) NULL,
        PermanentMeasureSummary NVARCHAR(MAX) NULL,
        InvestigationCompletionTime NVARCHAR(100) NULL,
        MeasureCompletionTime NVARCHAR(100) NULL,
        BActionStatus NVARCHAR(100) NULL,
        Sharing NVARCHAR(100) NULL,
        Report NVARCHAR(200) NULL,
        UpdatedDate DATETIME NULL,
        UpdatedBy NVARCHAR(100) NULL
    );
END;

DELETE FROM dbo.ErrorAnalysisLookups;

INSERT INTO dbo.ErrorAnalysisLookups (GroupCode, [Value], DisplayText, SortOrder, IsActive) VALUES
(N'Investigator', N'Lu?n', N'Lu?n', 1, 1),
(N'Investigator', N'M?nh', N'M?nh', 2, 1),
(N'Investigator', N'Giang', N'Giang', 3, 1),
(N'Investigator', N'?à', N'?à', 4, 1),
(N'Investigator', N'T?', N'T?', 5, 1),
(N'Investigator', N'Ngh?a', N'Ngh?a', 6, 1),
(N'Investigator', N'T?', N'T?', 7, 1),
(N'Investigator', N'Chinh', N'Chinh', 8, 1),
(N'Investigator', N'?.Quang', N'?.Quang', 9, 1),
(N'Investigator', N'Th?ng', N'Th?ng', 10, 1),
(N'Investigator', N'Tính', N'Tính', 11, 1),
(N'Investigator', N'Quy?n', N'Quy?n', 12, 1),
(N'Investigator', N'Hi?u', N'Hi?u', 13, 1),
(N'Investigator', N'H.Quân', N'H.Quân', 14, 1),
(N'Investigator', N'Thành', N'Thành', 15, 1),
(N'Investigator', N'Lâm', N'Lâm', 16, 1),
(N'Investigator', N'B.C??ng', N'B.C??ng', 17, 1),
(N'Investigator', N'T.Quang', N'T.Quang', 18, 1),
(N'Investigator', N'Minh', N'Minh', 19, 1),
(N'Investigator', N'D.C??ng', N'D.C??ng', 20, 1),
(N'Investigator', N'Phong', N'Phong', 21, 1),
(N'Investigator', N'Ph??ng', N'Ph??ng', 22, 1),
(N'Investigator', N'Nh?t', N'Nh?t', 23, 1),
(N'Investigator', N'Kiên', N'Kiên', 24, 1),
(N'Investigator', N'Hào Quân', N'Hào Quân', 25, 1),
(N'Investigator', N'H. Minh', N'H. Minh', 26, 1),
(N'Investigator', N'N.Anh', N'N.Anh', 27, 1),
(N'Investigator', N'??c', N'??c', 28, 1),
(N'Investigator', N'Toàn', N'Toàn', 29, 1),
(N'Investigator', N'Nam', N'Nam', 30, 1),
(N'Investigator', N'Trung', N'Trung', 31, 1),

(N'Category', N'Mono (??)', N'Mono (??)', 1, 1),
(N'Category', N'Color (??)', N'Color (??)', 2, 1),
(N'Category', N'Domino', N'Domino', 3, 1),
(N'Category', N'LM (TD2a, PD7, PJ8)', N'LM (TD2a, PD7, PJ8)', 4, 1),
(N'Category', N'Process (DRTN)', N'Process (DRTN)', 5, 1),

(N'Model', N'DSL', N'DSL', 1, 1),
(N'Model', N'ESL', N'ESL', 2, 1),
(N'Model', N'ECL', N'ECL', 3, 1),
(N'Model', N'FCL', N'FCL', 4, 1),
(N'Model', N'ELLe', N'ELLe', 5, 1),
(N'Model', N'ELL', N'ELL', 6, 1),
(N'Model', N'ELL PHE', N'ELL PHE', 7, 1),
(N'Model', N'FC', N'FC', 8, 1),
(N'Model', N'DLL DRTN', N'DLL DRTN', 9, 1),
(N'Model', N'DL DRTN', N'DL DRTN', 10, 1),
(N'Model', N'BL DRTN', N'BL DRTN', 11, 1),
(N'Model', N'FL DRTN', N'FL DRTN', 12, 1),
(N'Model', N'ELL DRTN', N'ELL DRTN', 13, 1),
(N'Model', N'ELLe DRTN', N'ELLe DRTN', 14, 1),
(N'Model', N'ESL DRTN', N'ESL DRTN', 15, 1),
(N'Model', N'DSL DRTN', N'DSL DRTN', 16, 1),
(N'Model', N'DCL DRTN', N'DCL DRTN', 17, 1),
(N'Model', N'ECL DRTN', N'ECL DRTN', 18, 1),
(N'Model', N'FCL DRTN', N'FCL DRTN', 19, 1),
(N'Model', N'BLL DRTN', N'BLL DRTN', 20, 1),
(N'Model', N'TD2a', N'TD2a', 21, 1),
(N'Model', N'PD7', N'PD7', 22, 1),
(N'Model', N'BH17/19', N'BH17/19', 23, 1),
(N'Model', N'PIJ', N'PIJ', 24, 1),
(N'Model', N'TIJ', N'TIJ', 25, 1),
(N'Model', N'PJ8', N'PJ8', 26, 1),
(N'Model', N'PD7 DASH', N'PD7 DASH', 27, 1),
(N'Model', N'BH21HT/SC', N'BH21HT/SC', 28, 1),

(N'Phenomenon', N'Ti?ng kêu ??', N'Ti?ng kêu ??', 1, 1),
(N'Phenomenon', N'Ngo?i quan ??', N'Ngo?i quan ??', 2, 1),
(N'Phenomenon', N'V?n chuy?n gi?y ??', N'V?n chuy?n gi?y ??', 3, 1),
(N'Phenomenon', N'Ch?t l??ng hình ?nh ??', N'Ch?t l??ng hình ?nh ??', 4, 1),
(N'Phenomenon', N'L?p ráp ??', N'L?p ráp ??', 5, 1),
(N'Phenomenon', N'Rò m?c ?????', N'Rò m?c ?????', 6, 1),
(N'Phenomenon', N'Chuy?n ??ng ??', N'Chuy?n ??ng ??', 7, 1),
(N'Phenomenon', N'L?ch qui cách ????', N'L?ch qui cách ????', 8, 1),
(N'Phenomenon', N'An toàn ??', N'An toàn ??', 9, 1),
(N'Phenomenon', N'Khác ???', N'Khác ???', 10, 1),

(N'StageClassification', N'Th? khuôn ??', N'Th? khuôn ??', 1, 1),
(N'StageClassification', N'Th? l??ng ??', N'Th? l??ng ??', 2, 1),
(N'StageClassification', N'S?n xu?t ??', N'S?n xu?t ??', 3, 1),

(N'OccurrenceProcess', N'QA m? h?p ????', N'QA m? h?p ????', 1, 1),
(N'OccurrenceProcess', N'QA l?y m?u Line ??????', N'QA l?y m?u Line ??????', 2, 1),
(N'OccurrenceProcess', N'QA ?? b?n tin c?y ?????', N'QA ?? b?n tin c?y ?????', 3, 1),
(N'OccurrenceProcess', N'QA môi tr??ng QA ??', N'QA môi tr??ng QA ??', 4, 1),
(N'OccurrenceProcess', N'QA ??c bi?t QA??', N'QA ??c bi?t QA??', 5, 1),
(N'OccurrenceProcess', N'Shipping', N'Shipping', 6, 1),
(N'OccurrenceProcess', N'Packing', N'Packing', 7, 1),
(N'OccurrenceProcess', N'Ki?m tra thành ph?m ?????', N'Ki?m tra thành ph?m ?????', 8, 1),
(N'OccurrenceProcess', N'Hontai', N'Hontai', 9, 1),
(N'OccurrenceProcess', N'Unit', N'Unit', 10, 1),
(N'OccurrenceProcess', N'Sparepart', N'Sparepart', 11, 1),
(N'OccurrenceProcess', N'Sub line', N'Sub line', 12, 1),
(N'OccurrenceProcess', N'Th? tr??ng ??', N'Th? tr??ng ??', 13, 1),
(N'OccurrenceProcess', N'Máy s? 0 0??', N'Máy s? 0 0??', 14, 1),
(N'OccurrenceProcess', N'?ánh giá ??c bi?t ????', N'?ánh giá ??c bi?t ????', 15, 1),

(N'InvestigationProgress', N'Close', N'Close', 1, 1),
(N'InvestigationProgress', N'?ang ?i?u tra ???', N'?ang ?i?u tra ???', 2, 1),
(N'InvestigationProgress', N'?ang th?o lu?n ??i sách c? h?u ???????', N'?ang th?o lu?n ??i sách c? h?u ???????', 3, 1),

(N'CauseClassification', N'Không rõ NN ????', N'Không rõ NN ????', 1, 1),
(N'CauseClassification', N'Linh ki?n ??', N'Linh ki?n ??', 2, 1),
(N'CauseClassification', N'Thao tác ??', N'Thao tác ??', 3, 1),
(N'CauseClassification', N'Thi?t k? ??', N'Thi?t k? ??', 4, 1),
(N'CauseClassification', N'Thi?t b? ??', N'Thi?t b? ??', 5, 1),
(N'CauseClassification', N'?? gá ??', N'?? gá ??', 6, 1),
(N'CauseClassification', N'Th?c l?c c?a máy ??', N'Th?c l?c c?a máy ??', 7, 1),
(N'CauseClassification', N'?ánh giá OK OK??', N'?ánh giá OK OK??', 8, 1),
(N'CauseClassification', N'Khác ???', N'Khác ???', 9, 1),

(N'TemporaryMeasureClassification', N'Linh ki?n ??', N'Linh ki?n ??', 1, 1),
(N'TemporaryMeasureClassification', N'Thao tác ??', N'Thao tác ??', 2, 1),
(N'TemporaryMeasureClassification', N'?? gá ??', N'?? gá ??', 3, 1),
(N'TemporaryMeasureClassification', N'?ánh giá OK OK??', N'?ánh giá OK OK??', 4, 1),
(N'TemporaryMeasureClassification', N'Thi?t b? ??', N'Thi?t b? ??', 5, 1),
(N'TemporaryMeasureClassification', N'Khác ???', N'Khác ???', 6, 1),

(N'PermanentMeasureClassification', N'Linh ki?n ??', N'Linh ki?n ??', 1, 1),
(N'PermanentMeasureClassification', N'Thao tác ??', N'Thao tác ??', 2, 1),
(N'PermanentMeasureClassification', N'?? gá ??', N'?? gá ??', 3, 1),
(N'PermanentMeasureClassification', N'?ánh giá OK OK??', N'?ánh giá OK OK??', 4, 1),
(N'PermanentMeasureClassification', N'Thi?t b? ??', N'Thi?t b? ??', 5, 1),
(N'PermanentMeasureClassification', N'Thi?t k? ??', N'Thi?t k? ??', 6, 1),
(N'PermanentMeasureClassification', N'Khác ???', N'Khác ???', 7, 1),

(N'Factory', N'Factory 1', N'Factory 1', 1, 1),
(N'Factory', N'Factory 2', N'Factory 2', 2, 1),
(N'Factory', N'Factory 3', N'Factory 3', 3, 1),
(N'Factory', N'Factory 4', N'Factory 4', 4, 1),
(N'Factory', N'Factory 5', N'Factory 5', 5, 1),
(N'Factory', N'Factory TTF', N'Factory TTF', 6, 1),

(N'ShipmentStop', N'B?o l?u xu?t hàng ????', N'B?o l?u xu?t hàng ????', 1, 1),
(N'ShipmentStop', N'D?ng xu?t hàng ????', N'D?ng xu?t hàng ????', 2, 1),

(N'ChronicDefect', N'Có ?', N'Có ?', 1, 1),
(N'ChronicDefect', N'Không ?', N'Không ?', 2, 1),

(N'BAction', N'QA SYS LSPM', N'QA SYS LSPM', 1, 1),
(N'BAction', N'QA', N'QA', 2, 1),
(N'BAction', N'SYS', N'SYS', 3, 1),
(N'BAction', N'LSPM', N'LSPM', 4, 1);
