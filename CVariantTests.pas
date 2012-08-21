unit CVariantTests;

interface
uses
   SysUtils,
   TestFramework,
   CVariants;

{$WARN UNSAFE_TYPE OFF}
   
type
  TPrimitiveTests = class(TTestCase)
  published
    procedure TestInteger;
    procedure TestString;
  end;

  TListTests = class(TTestCase)
  protected
    FList: CVariant;
    procedure SetUp; override;
  published
  end;


implementation


{ TPrimitiveTests }

procedure TPrimitiveTests.TestInteger;
var
  I1, I2: CVariant;
begin
  I1.Create(3);
  CheckEquals(3, I1.ToInt, 'Create, ToInt');
  I2 := I1;
  CheckEquals(3, I2.ToInt, 'Create, ToInt, Assign');
  // CheckEquals(3, CVariant.Make(3).ToInt, 'Make');
  I1.Destroy;
  I2.Destroy;
end;

procedure TPrimitiveTests.TestString;
var
  I1, I2: CVariant;
begin
  I1.Create('cbd');
  CheckEquals('cbd', I1.ToString, 'Create, ToString');
  I2 := I1;
  CheckEquals('cbd', I2.ToString, 'Create, ToString, Assign');
  // CheckEquals('cbd', CVariant.Make('cbd').ToString, 'Make');
  I1.Destroy;
  I2.Destroy;
end;

{ TListTests }

procedure TListTests.SetUp;
begin
  inherited;
  
end;

initialization
  RegisterTests('CVariants',
    [TPrimitiveTests.Suite
    ]);
end.
