program Create_Db;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {MainForm},
  HelpUnit in 'HelpUnit.pas' {HelpForm};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'ͨ�����ݿⴴ������';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
