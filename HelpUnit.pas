unit HelpUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons;

type
  THelpForm = class(TForm)
    GroupBox1: TGroupBox;
    Memo1: TMemo;
    btn_Exit: TBitBtn;
    Panel1: TPanel;
    Label1: TLabel;
    procedure btn_ExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  HelpForm: THelpForm;
  function Get_Version:String;
  
implementation

{$R *.dfm}

procedure THelpForm.btn_ExitClick(Sender: TObject);
begin
  Close;
end;

function Get_Version: String;
var
  VerInfoSize,VerValueSize,Dummy:Dword;
  VerInfo:Pointer;
  VerValue:PVSFixedFileInfo;
  sVer:String;V1,V2,V3,V4:word;
begin
  VerInfoSize:=GetFileVersionInfoSize(Pchar(ParamStr(0)),Dummy);
  GetMem(VerInfo,VerInfoSize);
  GetFileVersionInfo(PChar(ParamStr(0)),0,VerInfoSize,VerInfo);
  VerQueryValue(VerInfo,'\',Pointer(VerValue),VerValueSize);
  With   VerValue^   do
  begin
      V1:=dwFileVersionMS   shr   16;
      V2:=dwFileVersionMS   and   $FFFF;
      V3:=dwFileVersionLS   shr   16;
      V4:=dwFileVersionLS   and   $FFFF;
  end;
  FreeMem(VerInfo,VerInfoSize);
  sVer:=IntToStr(V1) + '.' + IntToStr(V2)+ '.' + IntToStr(V3) + '.'+IntToStr(V4);
  Result := sVer;
end;

procedure THelpForm.FormCreate(Sender: TObject);
var
  myinifn:String;
begin
  myinifn := ExtractFilePath(Application.ExeName)+'Create_Db_Set.ini';
  if FileExists(myinifn) then
  begin
    Memo1.Lines.Clear;
    Memo1.Lines.LoadFromFile(myinifn);
  end;
end;

end.
