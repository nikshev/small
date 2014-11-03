program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  HospitalUnit in 'HospitalUnit.pas' {HospitalsForm},
  NeTrud1Unit in 'NeTrud1Unit.pas' {NeTrud1Form},
  NeTrud2Unit in 'NeTrud2Unit.pas' {NeTrud2Form},
  EnterprisesUnit in 'EnterprisesUnit.pas' {EnterprisesForm},
  RelatedCommUnit in 'RelatedCommUnit.pas' {RelatedCommForm},
  FailureUnit in 'FailureUnit.pas' {FailureForm},
  DoctorTypeUnit in 'DoctorTypeUnit.pas' {DoctorTypeForm},
  DoctorFIOUnit in 'DoctorFIOUnit.pas' {DoctorFIOForm},
  OtherUnit in 'OtherUnit.pas' {OtherForm};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
