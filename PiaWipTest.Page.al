page 50101 "PIA WIP Test"
{
    PageType = Card;
    ApplicationArea = All;
    UsageCategory = Administration;
    Caption = 'PIA-rapport BETA';

    layout
    {
        area(Content)
        {
            group(General)
            {
                field(Result; Result)
                {
                    ApplicationArea = All;
                    Caption = 'Result';
                    Editable = false;
                }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(ExportPiaExcel)
            {
                ApplicationArea = All;
                Caption = 'Exportera PIA-rapport till Excel (TMG Sthlm)';
                Image = Export;

                trigger OnAction()
                var
                    Calc: Codeunit "Pia WIP Calculator";
                begin
                    Calc.ExportPiaToExcel_DetailPerJob();
                end;
            }


        }
    }

    var
        Result: Decimal;
}
