program Create_Db;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {MainForm},
  HelpUnit in 'HelpUnit.pas' {HelpForm};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '通用数据库创建程序';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
