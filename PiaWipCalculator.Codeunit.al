codeunit 50100 "Pia WIP Calculator"
{
    procedure CalculateDummy(): Decimal
    var
        Result: Decimal;
    begin
        Result := 12345.67;
        exit(Result);
    end;

    procedure CountCalcUnits(): Integer
    var
        CalcUnit: Record "PVS Job Calculation Unit";
    begin
        exit(CalcUnit.Count());
    end;

    procedure CountPrepressUnits(): Integer
    var
        CalcUnit: Record "PVS Job Calculation Unit";
    begin
        CalcUnit.SetRange("Unit", '210-PE20');
        exit(CalcUnit.Count());
    end;

    procedure SumCostForUnit(UnitCode: Code[20]): Decimal
    var
        CalcUnit: Record "PVS Job Calculation Unit";
        CalcDetail: Record "PVS Job Calculation Detail";
        TotalCost: Decimal;
    begin
        TotalCost := 0;

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

        exit(TotalCost);
    end;

    procedure SumPrepressCost(): Decimal
    begin
        exit(SumCostForUnit('210-PE20'));
    end;

    procedure SumPrepressCostForProductionJobs(): Decimal
    var
        JobRec: Record "PVS Job";
        CalcUnit: Record "PVS Job Calculation Unit";
        CalcDetail: Record "PVS Job Calculation Detail";
        TotalCost: Decimal;
    begin
        TotalCost := 0;

        JobRec.SetRange("Production Status Code", 'PRODUKTION');

        if JobRec.FindSet() then
            repeat
                CalcUnit.Reset();
                CalcUnit.SetRange("ID", JobRec."ID");
                CalcUnit.SetRange("Job", JobRec."Job");
                CalcUnit.SetRange("Version", JobRec."Version");
                CalcUnit.SetRange("Unit", '210-PE20');

                if CalcUnit.FindSet() then
                    repeat
                        CalcDetail.Reset();
                        CalcDetail.SetRange("ID", CalcUnit."ID");
                        CalcDetail.SetRange("Job", CalcUnit."Job");
                        CalcDetail.SetRange("Version", CalcUnit."Version");
                        CalcDetail.SetRange("Plan ID", CalcUnit."Plan ID");
                        CalcDetail.SetRange("Unit Entry No.", CalcUnit."Entry No.");
                        CalcDetail.SetRange("Calc. Unit", '210-PE20');

                        CalcDetail.CalcSums("Cost Amount");
                        TotalCost += CalcDetail."Cost Amount";
                    until CalcUnit.Next() = 0;
            until JobRec.Next() = 0;

        exit(TotalCost);
    end;

    procedure CountJobsByProductionStatus(StatusCode: Code[20]): Integer
    var
        JobRec: Record "PVS Job";
    begin
        JobRec.SetRange("Production Status Code", StatusCode);
        exit(JobRec.Count());
    end;

    procedure SumCostForUnitsByJobStatus(StatusCode: Code[20]; UnitCodes: List of [Code[20]]): Decimal
    var
        JobRec: Record "PVS Job";
        CalcUnit: Record "PVS Job Calculation Unit";
        CalcDetail: Record "PVS Job Calculation Detail";
        TotalCost: Decimal;
        UnitCode: Code[20];
    begin
        TotalCost := 0;

        JobRec.SetRange("Production Status Code", StatusCode);

        if JobRec.FindSet() then
            repeat
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
            until JobRec.Next() = 0;

        exit(TotalCost);
    end;

    procedure SumPrintCostForAfterbehandlingJobs(): Decimal
    var
        Units: List of [Code[20]];
    begin
        Units.Add('340-PE10');
        Units.Add('350-PE10');
        Units.Add('360-PE10');
        Units.Add('380-PE10');
        Units.Add('510-PE10');
        Units.Add('510-PE20');
        Units.Add('550-PE10');
        Units.Add('560-PE10');
        Units.Add('570-PE10');
        Units.Add('580-PE10');
        Units.Add('770-PE20');

        exit(SumCostForUnitsByJobStatus('EFTERBEHANDLING', Units));
    end;

    procedure SumFinishingCostForLeveransJobs(): Decimal
    var
        Units: List of [Code[20]];
    begin
        Units.Add('610-PE10');
        Units.Add('610-PE20');
        Units.Add('610-PE30');
        Units.Add('620-PE10');
        Units.Add('630-PE10');
        Units.Add('640-PE10');
        Units.Add('645-PE10');
        Units.Add('650-PE10');
        Units.Add('655-PE10');
        Units.Add('660-PE10');
        Units.Add('670-PE10');
        Units.Add('680-PE10');
        Units.Add('690-PE10');
        Units.Add('750-PE10');
        Units.Add('750-PE15');
        Units.Add('750-PE20');
        Units.Add('750-PE25');
        Units.Add('760-PE10');
        Units.Add('780-PE10');
        Units.Add('790-PE10');

        exit(SumCostForUnitsByJobStatus('LEVERANS', Units));
    end;

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

    local procedure WritePurchaseRowsForOrderNo(
        var ExcelBuf: Record "Excel Buffer" temporary;
        OrderNo: Code[20];
        CaseStatusCode: Code[20])
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
                Amt := PurchInvLine.Amount;

                // Purchase-rad (samma kolumnordning som headern)
                ExcelBuf.NewRow();
                ExcelBuf.AddColumn(OrderNo, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);                      // Order No.
                ExcelBuf.AddColumn('', false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);                           // Job
                ExcelBuf.AddColumn('', false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);                           // Version
                ExcelBuf.AddColumn('', false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);                           // Job Name
                ExcelBuf.AddColumn(PurchInvLine.Description, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);     // Purchase Name
                ExcelBuf.AddColumn(Format(CaseStatusCode), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);       // Production Status (från Case)

                ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);                          // Prepress PIA
                ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);                          // Tryck PIA
                ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);                          // Efterbeh. PIA
                ExcelBuf.AddColumn(Amt, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);                        // Inköp PIA
                ExcelBuf.AddColumn(Amt, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);                        // Total PIA (inköp)
            until PurchInvLine.Next() = 0;
    end;

    procedure ExportPiaToExcel_DetailPerJob()
    var
        JobRec: Record "PVS Job";
        CaseRec: Record "PVS Case";
        ExcelBuf: Record "Excel Buffer" temporary;

        PrepressUnits: List of [Code[20]];
        PrintUnits: List of [Code[20]];
        FinishUnits: List of [Code[20]];

        Prepress: Decimal;
        Print: Decimal;
        Finish: Decimal;
        Total: Decimal;

        JobName: Text[250];
        OrderNo: Code[20];

        FriendlyFileName: Text;
        SafeCompany: Text;
        DateStamp: Text;
    begin
        // Unit-listor
        PrepressUnits.Add('210-PE20');

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

        // Excel: header (ordningen du önskar)
        ExcelBuf.NewRow();
        ExcelBuf.AddColumn('Order No.', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Job', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Version', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Job Name', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Purchase Name', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Production Status', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Prepress PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Tryck PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Efterbeh. PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Inköp PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Total PIA', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);

        // Loop per Case (för att kunna lägga inköp direkt efter rätt order)
        CaseRec.Reset();

        // Case-status filter
        CaseRec.SetFilter("Status Code", '%1|%2|%3|%4|%5|%6|%7|%8',
            'ORDER', 'PROD.FÖRB', 'PREPRESS', 'KORREKTUR', 'PRODUKTION', 'EFTERBEHANDLING', 'LEVERANS', 'EFTERKALKYL');

        if CaseRec.FindSet() then
            repeat
                OrderNo := CaseRec."Order No.";

                // Jobbrader: endast de tre statusar som används för PIA-beräkning
                JobRec.Reset();
                JobRec.SetRange("ID", CaseRec."ID");
                JobRec.SetFilter("Production Status Code", '%1|%2|%3', 'PRODUKTION', 'EFTERBEHANDLING', 'LEVERANS');

                if JobRec.FindSet() then
                    repeat
                        Prepress := 0;
                        Print := 0;
                        Finish := 0;

                        // Prepress gäller när jobbet nått PRODUKTION eller längre
                        if JobRec."Production Status Code" in ['PRODUKTION', 'EFTERBEHANDLING', 'LEVERANS'] then
                            Prepress := SumCostForUnitsForJob(JobRec, PrepressUnits);

                        // Tryck gäller när jobbet nått EFTERBEHANDLING eller längre
                        if JobRec."Production Status Code" in ['EFTERBEHANDLING', 'LEVERANS'] then
                            Print := SumCostForUnitsForJob(JobRec, PrintUnits);

                        // Efterbeh gäller endast när jobbet är i LEVERANS
                        if JobRec."Production Status Code" = 'LEVERANS' then
                            Finish := SumCostForUnitsForJob(JobRec, FinishUnits);

                        Total := Prepress + Print + Finish;


                        JobName := JobRec."Job Name";

                        ExcelBuf.NewRow();
                        ExcelBuf.AddColumn(OrderNo, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(Format(JobRec."Job"), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(Format(JobRec."Version"), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(JobName, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn('', false, '', false, false, false, '', ExcelBuf."Cell Type"::Text); // Purchase Name blank
                        ExcelBuf.AddColumn(Format(JobRec."Production Status Code"), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);

                        ExcelBuf.AddColumn(Prepress, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(Print, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(Finish, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(0, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number); // Inköp PIA
                        ExcelBuf.AddColumn(Total, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                    until JobRec.Next() = 0;

                // Inköpsrader (bara om Amount <> 0) + skriv ut Case status på dessa rader
                WritePurchaseRowsForOrderNo(ExcelBuf, OrderNo, CaseRec."Status Code");

            until CaseRec.Next() = 0;

        // Skriv workbook + filnamn (SaaS)
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
    var
        Result: Text;
    begin
        Result := Value;

        Result := DelChr(Result, '=', '\/:*?"<>|');
        Result := ConvertStr(Result, ' ', '_');

        exit(Result);
    end;
}
