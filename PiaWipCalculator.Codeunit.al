codeunit 50100 "Pia WIP Calculator"
{

    // =====================================================
    // =================== CORE LOGIK =======================
    // =====================================================

    local procedure SumCostForUnitsForJob(JobRec: Record "PVS Job"; UnitCodes: List of [Code[20]]): Decimal
    var
        CalcUnit: Record "PVS Job Calculation Unit";
        CalcDetail: Record "PVS Job Calculation Detail";
        TotalCost: Decimal;
        UnitCode: Code[20];
    begin
        TotalCost := 0;

        foreach UnitCode in UnitCodes do begin
            CalcUnit.Reset();
            CalcUnit.SetRange("ID", JobRec."ID");
            CalcUnit.SetRange("Job", JobRec."Job");
            CalcUnit.SetRange("Version", JobRec."Version");
            CalcUnit.SetRange("Unit", UnitCode);

            if CalcUnit.FindSet() then
                repeat
                    CalcDetail.Reset();
                    CalcDetail.SetRange("ID", CalcUnit."ID");
                    CalcDetail.SetRange("Job", CalcUnit."Job");
                    CalcDetail.SetRange("Version", CalcUnit."Version");
                    CalcDetail.SetRange("Plan ID", CalcUnit."Plan ID");
                    CalcDetail.SetRange("Unit Entry No.", CalcUnit."Entry No.");
                    CalcDetail.SetRange("Calc. Unit", UnitCode);

                    CalcDetail.CalcSums("Cost Amount");
                    TotalCost += CalcDetail."Cost Amount";
                until CalcUnit.Next() = 0;
        end;

        exit(TotalCost);
    end;


    // =====================================================
    // ======= SPLIT PRINT INTO PAPER / NON-PAPER ==========
    // =====================================================

    local procedure SumCostForUnitsForJob_SplitPaper(
        JobRec: Record "PVS Job";
        UnitCodes: List of [Code[20]];
        var PaperCost: Decimal;
        var NonPaperCost: Decimal)
    var
        CalcUnit: Record "PVS Job Calculation Unit";
        CalcDetail: Record "PVS Job Calculation Detail";
        UnitCode: Code[20];
    begin
        PaperCost := 0;
        NonPaperCost := 0;

        foreach UnitCode in UnitCodes do begin
            CalcUnit.Reset();
            CalcUnit.SetRange("ID", JobRec."ID");
            CalcUnit.SetRange("Job", JobRec."Job");
            CalcUnit.SetRange("Version", JobRec."Version");
            CalcUnit.SetRange("Unit", UnitCode);

            if CalcUnit.FindSet() then
                repeat
                    CalcDetail.Reset();
                    CalcDetail.SetRange("ID", CalcUnit."ID");
                    CalcDetail.SetRange("Job", CalcUnit."Job");
                    CalcDetail.SetRange("Version", CalcUnit."Version");
                    CalcDetail.SetRange("Plan ID", CalcUnit."Plan ID");
                    CalcDetail.SetRange("Unit Entry No.", CalcUnit."Entry No.");
                    CalcDetail.SetRange("Calc. Unit", UnitCode);

                    if CalcDetail.FindSet() then
                        repeat
                            if CalcDetail."Item Type" = CalcDetail."Item Type"::Paper then
                                PaperCost += CalcDetail."Cost Amount"
                            else
                                NonPaperCost += CalcDetail."Cost Amount";
                        until CalcDetail.Next() = 0;

                until CalcUnit.Next() = 0;
        end;
    end;


    // =====================================================
    // ======= COLLECT PAPER ITEM NOS PER CASE =============
    // =====================================================

    local procedure CollectPaperItemNosForCase(
    CaseRec: Record "PVS Case";
    PrintUnits: List of [Code[20]];
    var PaperItemNos: Dictionary of [Code[20], Boolean])
    var
        JobRec: Record "PVS Job";
        CalcUnit: Record "PVS Job Calculation Unit";
        CalcDetail: Record "PVS Job Calculation Detail";
        UnitCode: Code[20];
    begin
        Clear(PaperItemNos);

        JobRec.Reset();
        JobRec.SetRange("ID", CaseRec."ID");
        JobRec.SetFilter("Production Status Code", '%1|%2', 'EFTERBEHANDLING', 'LEVERANS');

        if JobRec.FindSet() then
            repeat
                foreach UnitCode in PrintUnits do begin
                    CalcUnit.Reset();
                    CalcUnit.SetRange("ID", JobRec."ID");
                    CalcUnit.SetRange("Job", JobRec."Job");
                    CalcUnit.SetRange("Version", JobRec."Version");
                    CalcUnit.SetRange("Unit", UnitCode);

                    if CalcUnit.FindSet() then
                        repeat
                            CalcDetail.Reset();
                            CalcDetail.SetRange("ID", CalcUnit."ID");
                            CalcDetail.SetRange("Job", CalcUnit."Job");
                            CalcDetail.SetRange("Version", CalcUnit."Version");
                            CalcDetail.SetRange("Plan ID", CalcUnit."Plan ID");
                            CalcDetail.SetRange("Unit Entry No.", CalcUnit."Entry No.");
                            CalcDetail.SetRange("Calc. Unit", UnitCode);

                            if CalcDetail.FindSet() then
                                repeat
                                    if (CalcDetail."Item Type" = CalcDetail."Item Type"::Paper) and
                                       (CalcDetail."Item No." <> '') then begin

                                        if PaperItemNos.ContainsKey(CalcDetail."Item No.") then
                                            PaperItemNos.Set(CalcDetail."Item No.", true)
                                        else
                                            PaperItemNos.Add(CalcDetail."Item No.", true);

                                    end;

                                until CalcDetail.Next() = 0;

                        until CalcUnit.Next() = 0;

                end;

            until JobRec.Next() = 0;
    end;


    // =====================================================
    // ============ WRITE PURCHASE ROWS =====================
    // =====================================================

    local procedure WritePurchaseRowsForOrderNo(
        var ExcelBuf: Record "Excel Buffer" temporary;
        OrderNo: Code[20];
        CaseStatusCode: Code[20];
        PaperItemNos: Dictionary of [Code[20], Boolean])
    var
        PurchInvLine: Record "Purch. Inv. Line";
        Amt: Decimal;
    begin
        if OrderNo = '' then
            exit;

        PurchInvLine.Reset();
        PurchInvLine.SetRange("PVS Order No.", OrderNo);
        PurchInvLine.SetFilter(Amount, '<>%1', 0);

        if PurchInvLine.FindSet() then
            repeat
                if not PaperItemNos.ContainsKey(PurchInvLine."No.") then begin
                    Amt := PurchInvLine.Amount;

                    ExcelBuf.NewRow();
                    ExcelBuf.AddColumn(OrderNo, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(PurchInvLine.Description, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(Format(CaseStatusCode), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);

                    ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(Amt, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(Amt, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                end;
            until PurchInvLine.Next() = 0;
    end;


    // =====================================================
    // ================== EXPORT ============================
    // =====================================================

    procedure ExportPiaToExcel_DetailPerJob()
    var
        JobRec: Record "PVS Job";
        CaseRec: Record "PVS Case";
        ExcelBuf: Record "Excel Buffer" temporary;

        PrepressUnits: List of [Code[20]];
        PrintUnits: List of [Code[20]];
        FinishUnits: List of [Code[20]];

        Prepress: Decimal;
        PrintNonPaper: Decimal;
        Paper: Decimal;
        Finish: Decimal;
        Total: Decimal;

        OrderNo: Code[20];
        JobName: Text[250];

        PaperItemNos: Dictionary of [Code[20], Boolean];

        FriendlyFileName: Text;
        SafeCompany: Text;
        DateStamp: Text;
    begin
        // === Units ===
        PrepressUnits.Add('210-PE20');
        PrepressUnits.Add('210-PE10');

        PrintUnits.Add('340-PE10');
        PrintUnits.Add('350-PE10');
        PrintUnits.Add('360-PE10');
        PrintUnits.Add('380-PE10');
        PrintUnits.Add('510-PE10');
        PrintUnits.Add('510-PE20');
        PrintUnits.Add('550-PE10');
        PrintUnits.Add('560-PE10');
        PrintUnits.Add('570-PE10');
        PrintUnits.Add('580-PE10');
        PrintUnits.Add('770-PE20');

        FinishUnits.Add('610-PE10');
        FinishUnits.Add('610-PE20');
        FinishUnits.Add('610-PE30');
        FinishUnits.Add('620-PE10');
        FinishUnits.Add('630-PE10');
        FinishUnits.Add('640-PE10');
        FinishUnits.Add('645-PE10');
        FinishUnits.Add('650-PE10');
        FinishUnits.Add('655-PE10');
        FinishUnits.Add('660-PE10');
        FinishUnits.Add('670-PE10');
        FinishUnits.Add('680-PE10');
        FinishUnits.Add('690-PE10');
        FinishUnits.Add('750-PE10');
        FinishUnits.Add('750-PE15');
        FinishUnits.Add('750-PE20');
        FinishUnits.Add('750-PE25');
        FinishUnits.Add('760-PE10');
        FinishUnits.Add('780-PE10');
        FinishUnits.Add('790-PE10');

        // === Header ===
        ExcelBuf.NewRow();
        ExcelBuf.AddColumn('Order No.', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Job', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Version', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Job Name', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Purchase Name', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Production Status', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Prepress PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Tryck PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Papper PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Efterbeh. PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Inköp PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Total PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);

        // === Loop per Case ===
        CaseRec.Reset();
        CaseRec.SetFilter("Status Code", '%1|%2|%3|%4|%5|%6|%7|%8',
            'ORDER', 'PROD.FÖRB', 'PREPRESS', 'KORREKTUR', 'PRODUKTION', 'EFTERBEHANDLING', 'LEVERANS', 'EFTERKALKYL');

        CaseRec.SetCurrentKey("Order No.");
        CaseRec.Ascending(true);

        if CaseRec.FindSet() then
            repeat
                OrderNo := CaseRec."Order No.";

                // Samla Paper-itemnos för detta case
                CollectPaperItemNosForCase(CaseRec, PrintUnits, PaperItemNos);

                // Jobbrader
                JobRec.Reset();
                JobRec.SetRange("ID", CaseRec."ID");
                JobRec.SetFilter("Production Status Code", '%1|%2|%3',
                    'PRODUKTION', 'EFTERBEHANDLING', 'LEVERANS');

                if JobRec.FindSet() then
                    repeat
                        Prepress := 0;
                        PrintNonPaper := 0;
                        Paper := 0;
                        Finish := 0;

                        if JobRec."Production Status Code" in ['PRODUKTION', 'EFTERBEHANDLING', 'LEVERANS'] then
                            Prepress := SumCostForUnitsForJob(JobRec, PrepressUnits);

                        if JobRec."Production Status Code" in ['EFTERBEHANDLING', 'LEVERANS'] then
                            SumCostForUnitsForJob_SplitPaper(JobRec, PrintUnits, Paper, PrintNonPaper);

                        if JobRec."Production Status Code" = 'LEVERANS' then
                            Finish := SumCostForUnitsForJob(JobRec, FinishUnits);

                        Total := Prepress + PrintNonPaper + Paper + Finish;

                        JobName := JobRec."Job Name";

                        ExcelBuf.NewRow();
                        ExcelBuf.AddColumn(OrderNo, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(Format(JobRec."Job"), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(Format(JobRec."Version"), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(JobName, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn('', false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(Format(JobRec."Production Status Code"), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(Prepress, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(PrintNonPaper, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(Paper, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(Finish, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(Total, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                    until JobRec.Next() = 0;

                // Purchase-rader (exkl papper)
                WritePurchaseRowsForOrderNo(
                    ExcelBuf,
                    OrderNo,
                    CaseRec."Status Code",
                    PaperItemNos);

            until CaseRec.Next() = 0;

        ExcelBuf.CreateNewBook('PIA');
        ExcelBuf.WriteSheet('PIA', CompanyName(), UserId());
        ExcelBuf.CloseBook();

        SafeCompany := MakeSafeFileNamePart(CompanyName());
        DateStamp := Format(Today(), 0, '<Year4><Month,2><Day,2>');
        FriendlyFileName := StrSubstNo('PIA-rapport_%1_%2', SafeCompany, DateStamp);

        ExcelBuf.SetFriendlyFilename(FriendlyFileName);
        ExcelBuf.OpenExcel();
    end;


    local procedure MakeSafeFileNamePart(Value: Text): Text
    begin
        Value := DelChr(Value, '=', '\/:*?"<>|');
        Value := ConvertStr(Value, ' ', '_');
        exit(Value);
    end;

}