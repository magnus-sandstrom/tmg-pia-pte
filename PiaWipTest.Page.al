page 50101 "PIA WIP Test"
{
    PageType = Card;
    ApplicationArea = All;
    UsageCategory = Administration;
    Caption = 'PIA WIP Test';

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
            action(CalculateDummy)
            {
                ApplicationArea = All;
                Caption = 'Calculate (Dummy)';
                Image = Calculate;

                trigger OnAction()
                var
                    Calc: Codeunit "Pia WIP Calculator";
                begin
                    Result := Calc.CalculateDummy();
                    Message('Dummy result: %1', Result);
                end;
            }
        }
    }

    var
        Result: Decimal;
}
