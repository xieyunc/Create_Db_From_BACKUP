unit MainUnit;

interface

uses
   Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
   StdCtrls ,FileCtrl,ExtCtrls, DB, ADODB, ShellApi,Dialogs, Buttons, Gauges,IniFiles,
  ComCtrls, StatusBarEx;

type
  TMainForm = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    IPEdit: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    saEdit: TEdit;
    sapwdEdit: TEdit;
    newdbfnEdit: TEdit;
    newdbsaEdit: TEdit;
    newdbpwdEdit: TEdit;
    Bevel1: TBevel;
    btn_Open: TSpeedButton;
    btn_Start: TButton;
    btn_Close: TButton;
    OpenDialog1: TOpenDialog;
    ADOConnection1: TADOConnection;
    Memo1: TMemo;
    newdbEdit: TComboBox;
    Access_Connection: TADOConnection;
    tmpquery: TADOQuery;
    access_query: TADOQuery;
    CheckBox2: TCheckBox;
    StatusBarEx1: TStatusBarEx;
    btn_Help: TButton;
    sp1: TADOStoredProc;
    chk_proc: TCheckBox;
    procedure btn_OpenClick(Sender: TObject);
    procedure btn_CloseClick(Sender: TObject);
    procedure btn_StartClick(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure btn_HelpClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    //initaccessfn:string;              //��ʼ��ACCESS���ݿ��ļ�
    //dbscriptfn:string;                //���ݿⴴ���ű��ļ���
    createprocfn:string;//�����洢���̵Ľű��ļ�
    mdffn,ldffn,backupdir,backupfn:string;
    Initialized_dbscriptfn:string;    //�Ѿ���ʼ���õ������ݿⴴ���ű��ļ���
    procedure SysInitialize;
    function DbisExists(dbName:String;var dbfn:String):Boolean;
    function Create_Db(dbName:String;dbfn,logfn:String): Boolean;
    function Drop_Db(dbName:String):Boolean;
    //function Init_Create_db_fn(sql_fn:String):Boolean;
    function CreateDataBase(const dbname,FileName:string):boolean;
    function DBSrv_Connect_Is_OK:Boolean;
    function CreateDbUser(const DbName,uSaName,uSaPwd:string):Boolean;//�������ݿ��û�
    function CreateProcedure:Boolean;//�����洢����
  public
    { Public declarations }
  end;

  function AccessDB_Is_OK(dbfn:String):Boolean;
  function SplitString(const source,ch:string):tstringlist;
  function ReplaceStr(Str, SearchStr, ReplaceStr: string): string;
  function GetLocalHostName():string;

var
  MainForm: TMainForm;
  Log_Strings,Table_Name_Strings:TStrings;

implementation

uses HelpUnit;

{$R *.dfm}
var
  vmsg:String;

function GetLocalHostName():string;
var
  s:array[1..127] of Char;
  i:DWORD;
begin
  i := 127;
  GetComputerName(@s,i);
  Result := s;
end;

//�����ַ���
function SplitString(const source,ch:string):tstringlist;
var
  temp:string;
  i:integer;
begin
  Result:=tstringlist.Create;
  temp:=source;
  i:=pos(ch,source);
  while i<>0 do
  begin
    Result.Add(copy(temp,0,i-1));
    delete(temp,1,i);
    i:=pos(ch,temp);
  end;
  Result.Add(Trim(temp));
end;

//�ַ����滻
function ReplaceStr(Str, SearchStr, ReplaceStr: string): string;
begin
  while Pos(SearchStr, Str) <> 0 do
  begin
    Insert(ReplaceStr, Str, Pos(SearchStr, Str));
    Delete(Str, Pos(SearchStr, Str), Length(SearchStr));
  end;
  Result := Str;
end;

function  TMainForm.CreateDataBase(const dbname,FileName:string):boolean;
var
  //strlist:TStringList ;
  //i:integer;
  dbfn:string;
begin
  StatusBarEx1.Panels.Items[1].Text := '���ڴ������ݿ�....';
  if DbisExists(dbname,dbfn) then
  begin
    Result := True;
    Exit;
  end;

  if not Create_Db(dbname,newdbfnEdit.Text+newdbEdit.Text+'_Data.mdf',newdbfnEdit.Text+newdbEdit.Text+'_Log.ldf') then
  begin
    Result := False;
    Exit;
  end;
  if chk_proc.Checked then
  begin
    StatusBarEx1.Panels.Items[1].Text := '���ڴ���master���ݿ�洢����....';
    CreateProcedure;
  end;
  Result:=true;
end;

function TMainForm.CreateDbUser(const DbName, uSaName, uSaPwd: string): Boolean;
var
  tmpquery:TADOQuery;
begin
  tmpquery := TADOQuery.Create(nil);
  tmpquery.Connection := ADOConnection1;
  try
    try
      with tmpquery do
      begin
        close;
        SQL.Text := 'use ['+DbName+']';
        ExecSQL;

        SQL.clear;
        sql.Add('if not exists (select * from master.dbo.syslogins where loginname = N'+quotedstr(uSaName)+')');
        sql.Add('BEGIN');
        SQL.Add('  declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'+quotedstr('master')+', @loginlang = N'+quotedstr('��������'));
        SQL.Add('  if @logindb is null or not exists (select * from master.dbo.sysdatabases where name = @logindb)');
        sql.Add('    select @logindb = N'+quotedstr('master'));
        SQL.Add('  if @loginlang is null or (not exists (select * from master.dbo.syslanguages where name = @loginlang) and @loginlang <> N'+quotedstr('us_english')+')');        sql.Add('    select @loginlang = @@language');
        sql.add('  exec sp_addlogin N'+quotedstr(uSaName)+','+quotedstr(uSaPwd)+', @logindb, @loginlang');
        sql.add('END');
        //showmessage(sql.text);
        ExecSQL;

        sql.clear;
        sql.add('if not exists (select * from dbo.sysusers where name = N'+quotedstr(uSaName)+' and uid < 16382)');
        sql.add('EXEC sp_grantdbaccess N'+quotedstr(uSaName));
        try
          Execsql;
        except
        end;

        sql.clear;
        sql.add('exec sp_addrolemember N'+quotedstr('db_owner')+', N'+quotedstr(uSaName));
        try
          ExecSQL;
        except
        end;
      end;
      Result := True;
    except
      Result := False;
    end;
  finally
    tmpquery.Free;
  end;
end;

function TMainForm.CreateProcedure: Boolean;
var
  str:string;
  sList:TStrings;
  tmpquery:TAdoquery;
  i:Integer;
begin
  Result := False;
  if not FileExists(createprocfn) then
     Exit;
  sList := TStringList.Create;
  tmpquery := TAdoQuery.Create(nil);
  tmpquery.ParamCheck := False;
  tmpquery.connection := AdoConnection1;
  try
    sList.LoadFromFile(createprocfn);
    tmpquery.close;
    tmpquery.sql.clear;
    i:=0;
    while i<sList.Count do
    begin
      str := sList[i];
      if uppercase(trim(str))='GO' then
      begin
      try
        tmpquery.ExecSql;
      except
        //on e:exception do
        begin
          Result := False;
          //ShowMessage(e.message);//+#13+tmpquery.sql.text);
          Exit;
        end;
      end;
        tmpquery.close;
        tmpquery.sql.clear;
      end
      else if (trim(str)<>'') then //and (copy(trimLeft(str),1,2)<>'--') then 
      begin
        tmpquery.sql.add(str);
      end;

      Inc(i);
    end;
    Result := True;
  finally
    tmpquery.Free;
    sList.Free;
  end;
end;

procedure TMainForm.btn_OpenClick(Sender: TObject);
const
  SELDIRHELP = 1000;
var
  tmp_dir :String;
begin
  //tmp_dir := newdbfnEdit.Text;
  
  //if SelectDirectory(tmp_dir, [sdAllowCreate, sdPerformCreate, sdPrompt],SELDIRHELP) then
  if SelectDirectory('��ѡ�����ݿ��ļ����Ŀ¼��','',tmp_dir) then
    newdbfnEdit.Text := tmp_dir;
end;

procedure TMainForm.btn_CloseClick(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.btn_StartClick(Sender: TObject);
var
  dbfn,s :string;
  //old,i,ii,j:integer;
  is_OK :Boolean;
begin
  try
    screen.Cursor := crHourGlass;
    if newdbEdit.Text='' then
    begin
      MessageBox(Handle, '���ݿ�������Ϊ�գ�������Ҫ���������ݿ����ƣ�����', 
        '���ݿ�������Ϊ��', MB_OK + MB_ICONSTOP);
      Exit;
    end;

    Log_Strings.Clear;

    btn_Start.Enabled := False;

    if not DBSrv_Connect_Is_OK then
       Exit; 

    if Application.MessageBox(pchar('���Ҫ������Ϊ��'+newdbEdit.Text+'�������ݿ���'),'����ȷ��',MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON2)<>idyes then
       Exit;

    if DbisExists(newdbEdit.Text,dbfn) then
    begin
       if Application.MessageBox(pchar('ͬ�����ݿ⣺��'+newdbEdit.Text+'���Ѿ����ڣ�'+#13+'���ļ�λ�ڣ���'+dbfn+'����'+#13+#13+'����ٴδ����Ļ��������ԭ�������ݿ⼰��������ɾ��������'+#13+#13+'��ȷ����Ҫɾ�������´�����һͬ�����ݿ���'),'����ͬ�����ݿ�',MB_YESNO+MB_ICONWARNING+MB_DEFBUTTON2)<>idyes then
          exit
       else
       begin
         s := '';

         if InputQuery('��ȷ��', '�����롼OK�������ַ��Ա�ȷ�ϣ�',s) then
         begin
           if UpperCase(s)<>'OK' then
           begin
             //is_OK := False;
             vmsg := 'ȷ���ַ�����֤ʧ�ܣ�������ȡ����';
             Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
             Application.MessageBox(PChar(vmsg),'������ֹ',MB_OK+MB_ICONERROR);
             Exit;
           end else
           begin
             //is_OK := True;
             Application.ProcessMessages;
             if not Drop_Db(newdbEdit.Text) then
             begin
               //is_OK := False;
               vmsg := '���ݿ⡼'+newdbEdit.Text+'��ɾ��ʧ�ܣ������������û�����ʹ�ã�����'+#13+'���ȷʵҪɾ�������������ݿ���������ԣ�';
               Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
               Application.MessageBox(pchar(vmsg),'���ݿ�ɾ��ʧ��',MB_OK+MB_ICONERROR);
               Exit;
             end;
           end;
         end
         else
         begin
           //is_OK := False;
           vmsg := '�û�ȡ������������ֹ��';
           Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
           Application.MessageBox(PChar(vmsg),'������ֹ',MB_OK+MB_ICONINFORMATION);
           Exit;
         end;
       end;

       s := newdbfnedit.Text ;

       if s[length(s)]<>'\' then
         s := s+'\';

       newdbfnedit.Text := s;

       if not DirectoryExists(s) then
          ForceDirectories(s);

       sleep(1000);

       Memo1.Lines.Clear;

    end;

    //if is_OK then
    begin
      s := newdbfnedit.Text ;

      if s[length(s)]<>'\' then
        s := s+'\';

      newdbfnedit.Text := s;

      if not DirectoryExists(s) then
      begin
         ForceDirectories(s);
         Sleep(1000);
      end;
      
      if not CreateDataBase(newdbEdit.Text,Initialized_dbscriptfn) then
      begin
        vmsg := '���ݿⴴ��ʧ�ܣ������ԣ�����';
        Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
        Application.MessageBox(PChar(vmsg),'����ʧ��',MB_OK+MB_ICONERROR);
        Exit;
      end;

      try
        adoconnection1.DefaultDatabase := 'master';
        sp1.Parameters.ParamByname('@restore_db_name').value := newdbEdit.Text;
        sp1.Parameters.ParamByname('@filename').value := backupfn;
        sp1.ExecProc;
        is_OK := sp1.parameters.ParamByName('@flag').value='ok';
      except
        on e:Exception do
        begin
          vmsg := e.Message;
          Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
        end;
      end;
      //if is_ok then
      is_ok := CreateDbUser(newdbedit.Text,newdbSaEdit.Text,newdbPwdEdit.Text);

    end;

    if is_OK then
    begin
      vmsg := '���ݿ⣺��'+newdbEdit.Text+'�������ɹ�������';
      Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
      Application.MessageBox(pchar(vmsg+#13+'���ס���ݿ��û���ɫ�����룡����'),'�������',MB_OK+MB_ICONINFORMATION)
    end else
    begin
      vmsg := '���ݿ⣺��'+newdbEdit.Text+'������ʧ�ܣ�����';
      Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
      Application.MessageBox(pchar(vmsg+#13+#13+'��ر����������ݿ����ӵ�Ӧ�ó�������´�����'),'����ʧ��',MB_OK+MB_ICONERROR);
    end;
    //Button2.Click;

  finally
    screen.Cursor := crDefault;
    btn_Start.Enabled := True;
    ADOConnection1.CLOSE;
  end;
end;

function TMainForm.DbisExists(dbName:String;var dbfn:String): Boolean;
var
  adoquery1:Tadoquery;
begin
  adoquery1 := TAdoquery.Create(nil);
  try
    Result := False;
    with adoquery1  do
    begin
      CommandTimeout := 300;
      Connection := AdoConnection1;
      close;
      sql.Clear;
      sql.Add('SELECT name,filename FROM master.dbo.sysdatabases WHERE name = '+quotedstr(dbName));
      Prepared := true;
      Open;
      dbfn := fieldbyname('filename').AsString;
      Result := Recordcount>0 ;
      Close;
    end;
  finally
    adoquery1.Free;
  end;
end;

function TMainForm.Drop_Db(dbName:String): Boolean;
var
  adoquery1:Tadoquery;
  sql_str :String;
begin
  adoquery1 := TAdoquery.Create(nil);
  try
    Result := False;
    vmsg := '����ɾ�������ݿ�....';
    StatusBarEx1.Panels.Items[1].Text := vmsg;
    Application.ProcessMessages;
    sql_str := 'IF EXISTS (SELECT * FROM sysdatabases WHERE name = '+quotedstr(dbname)+') BEGIN DROP database '+dbname+' END;';

    //sql_str := 'DROP database '+dbName;
    //sql_str := 'sp_detach_db '+quotedstr(dbname)+','+quotedstr('true');

    with adoquery1  do
    begin
      CommandTimeout := 300;
      Connection := AdoConnection1;
      close;
      sql.Clear;
      sql.Add(sql_str);
      try
        ExecSql;
        Result := True;
        Close;
        vmsg := vmsg+'�ɹ���';
        Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
      except
        on e:Exception do begin
          vmsg := vmsg+'ʧ�ܣ�'+e.Message;
          Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
          Result := False;
        end;
      end;
    end;
  finally
    adoquery1.Free;
  end;
end;

function TMainForm.Create_Db(dbName:String;dbfn,logfn:String): Boolean;
var
  adoquery1:Tadoquery;
  sql_str :String;
begin
  adoquery1 := TAdoquery.Create(nil);
  try
    Result := False;

    deletefile(dbfn);
    deletefile(logfn);

    sql_str := 'CREATE DATABASE '+dbName+' ON (NAME='+dbName+'_dat,FILENAME='+quotedstr(dbfn)+',SIZE=4096KB,FILEGROWTH = 10%) '+
               'LOG ON (NAME='+dbName+'_log,FILENAME='+quotedstr(logfn)+',SIZE=2048KB,FILEGROWTH = 10%)';

    with adoquery1  do
    begin
      Connection := AdoConnection1;
      CommandTimeout := 300;
      close;
      sql.Clear;
      sql.Add(sql_str);
      try
        ExecSql;
        Result := True;
        Close;
        vmsg := '���ݿⴴ���ɹ���';
        Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
      except
        on e:Exception do begin
          vmsg := e.Message;
          Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
          Result := False;
        end;
      end;
    end;
  finally
    adoquery1.Free;
  end;
end;

{
function TMainForm.Init_Create_db_fn(sql_fn:String): Boolean;
var
  //i,ii:integer;
  s:string;
begin
  try
    Result := False;

    if not fileexists(sql_fn) then
    begin
       vmsg := '���ݿⴴ���ű��ļ���'+sql_fn+'δ�ҵ������飡';
       Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
       Application.MessageBox(pchar(vmsg),'�ļ�δ�ҵ�',MB_OK+MB_ICONINFORMATION);
       Exit;
    end;

    vmsg := '���ڳ�ʼ�����ݿⴴ���ű�';
    Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
    StatusBarEx1.Panels.Items[1].Text := vmsg;
    Application.ProcessMessages;

    Memo1.Lines.LoadFromFile(sql_fn);

    Initialized_dbscriptfn := ExtractFilePath(Application.ExeName)+'create_db.sql';
    DeleteFile(Initialized_dbscriptfn);

    Application.ProcessMessages;

    s := '$DB_NAME$';
    Memo1.Text := ReplaceStr(Memo1.Text,s,newdbEdit.Text);

    s := '$DB_PATH$';
    Memo1.Text := ReplaceStr(Memo1.Text,s,newdbfnEdit.Text);

    s := '$SA_NAME$';
    Memo1.Text := ReplaceStr(Memo1.Text,s,newdbsaEdit.Text);

    s := '$SA_PWD$';
    Memo1.Text := ReplaceStr(Memo1.Text,s,newdbpwdEdit.Text);

    Application.ProcessMessages;

    Memo1.Lines.SaveToFile(Initialized_dbscriptfn);

    //
    s := ExtractFilePath(Application.ExeName);
    if s[length(s)]<>'\' then
      s := s+'\';

    Memo1.Lines.Clear;
    if not CheckBox2.Checked then
       Memo1.Lines.Add('isqlw -S "'+IPEdit.Text+'" -U '+saEdit.Text+' -P '+sapwdEdit.Text+' -i "'+sql_fn+'" -o '+s+'result.txt')
    else
       Memo1.Lines.Add('isqlw -S "'+IPEdit.Text+'" -E '+' -i "'+sql_fn+'" -o '+s+'result.txt');

    Memo1.Lines.SaveToFile(s+'create_db.bat');
    //
    
    Result := True;
    Log_Strings.Add(DateTimeToStr(now)+'  '+'........��ɣ�')
  except
    on e:Exception do begin
      vmsg := '.........ʧ�ܣ�'+e.Message;
      Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
      Result := False;
    end;
  end;
end;
}

function AccessDB_Is_OK(dbfn:String):Boolean;
begin
  Result := True;
  with MainForm.Access_Connection do
  begin
    Close;
    ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+dbfn+';Persist Security Info=False';
    try
      Open;
      Close;
      vmsg := dbfn+'���ݿ����ӳɹ���';
      Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
    except
      on e:Exception do begin
        vmsg := dbfn+'���ݿ�����ʧ�ܣ�'+e.message;
        Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
        Result := False;
      end;
    end;
  end;
end;

procedure TMainForm.CheckBox2Click(Sender: TObject);
begin
  Label2.Enabled := not CheckBox2.Checked;
  Label3.Enabled := Label2.Enabled;
  saEdit.Enabled := Label2.Enabled;
  sapwdEdit.Enabled := Label2.Enabled;
end;

procedure TMainForm.SysInitialize;
var
  ss,myinifn:string;
begin
  myinifn := ExtractFilePath(Application.ExeName)+'Create_Db_Set.ini';
  btn_Start.Enabled := FileExists(myinifn);
  if not btn_Start.Enabled then
  begin
    vmsg := 'ϵͳ��ʼ���ļ���Create_Db_Set.ini �����ڣ�����';
    Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
    MessageBox(Handle, PChar(vmsg),'��ʼ���ļ�������', MB_OK + MB_ICONSTOP);
    Application.Terminate;
  end;

  with TIniFile.Create(myinifn) do
  begin
    try
      IPEdit.Text := GetLocalHostName();

      newdbEdit.Items.Clear;
      ss := ReadString('CREATE_DB_SET','SYSNAME','');
      Application.Title := ss;
      Caption := ss+'--���ݿⴴ������';
      Caption := Caption+' Ver '+Get_Version;

      ss := ReadString('CREATE_DB_SET','DBNAMELIST','');
      newdbEdit.Items.AddStrings(SplitString(ss,'|'));
      if Self.newdbEdit.Items.Count>0 then
        newdbEdit.ItemIndex := 0;

      newdbfnEdit.Text := ReadString('CREATE_DB_SET','DBSAVEDIR','');

      newdbsaEdit.Text := ReadString('CREATE_DB_SET','SANAME','');

      createprocfn := ExtractFilePath(ParamStr(0))+ReadString('CREATE_DB_SET','CREATEPROCEDURESCRIPT','');

      mdffn := ExtractFilePath(ParamStr(0))+ReadString('CREATE_DB_SET','MDFFILE','');
      ldffn := ExtractFilePath(ParamStr(0))+ReadString('CREATE_DB_SET','LDFFILE','');
      backupdir := ReadString('CREATE_DB_SET','BACKUPDIR','');
      if not DirectoryExists(backupdir) then
         ForceDirectories(backupdir);

      backupfn := ExtractFilePath(ParamStr(0))+ReadString('CREATE_DB_SET','BACKUPFILE','');

    finally
      Free;
    end;
  end;

  btn_Start.Enabled := FileExists(backupfn);
  if not btn_Start.Enabled then
  begin
    vmsg := '���ݿⴴ���ű��ļ���'+backupfn+' �����ڣ�����';
    Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
    MessageBox(Handle, PChar(vmsg),'���ݿ�ű��ļ�������', MB_OK + MB_ICONSTOP);
    Application.Terminate;
  end;

end;

procedure TMainForm.btn_HelpClick(Sender: TObject);
begin
  with THelpForm.Create(Application) do
  begin
    ShowModal;
    Free;
  end;
end;

function TMainForm.DBSrv_Connect_Is_OK: Boolean;
var
  connect_str:String;
begin
  Result := False;

  if not CheckBox2.Checked then
  begin
    connect_str := 'Provider=SQLOLEDB.1;Password='+SaPwdEdit.Text+';Persist Security Info=True;User ID='+SaEdit.Text+';';
    connect_str := connect_str + 'Initial Catalog=master;Data Source='+IPEdit.Text+';';
    connect_str := connect_str + 'Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False';
  end else
  begin
    connect_str := 'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source='+IPEdit.Text;
  end;

  try
    vmsg := '�����������ݿ������....';
    Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
    StatusBarEx1.Panels.Items[1].Text := vmsg;
    Application.ProcessMessages;
    AdoConnection1.Close;
    ADOConnection1.ConnectionTimeout := 5;
    adoconnection1.ConnectionString := connect_str;
    Adoconnection1.Open;
    vmsg := '���ݿ���������ӳɹ���';
    Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
    Result := True;
  except
    on e:Exception do begin
      vmsg := '���ݿ����������ʧ�ܣ��������ݿ�ϵͳ����'+#13+'����Ա��ɫ�������Ƿ���ȷ��';
      Log_Strings.Add(DateTimeToStr(now)+'  '+vmsg);
      Application.MessageBox(PChar(vmsg),'�������ݿ������ʧ��',MB_OK+MB_ICONERROR);
      Result := False;
    end;
  end;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if FileExists(Initialized_dbscriptfn) then
     DeleteFile(Initialized_dbscriptfn);
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  SysInitialize;//׼��������ʼ��
end;

initialization
  Table_Name_Strings := TStringList.Create;
  Log_Strings := TStringList.Create;

finalization
  if Log_Strings.Text<>'' then
  begin
    Log_Strings.SaveToFile('Result.Log');
    if Application.MessageBox('��Ҫ�鿴���ݿⴴ��������־�𣿡���', 
      '������ʾ', MB_YESNO + MB_ICONQUESTION) = IDYES then
    begin
      ShellExecute(Application.Handle,'OPEN','Result.Log',nil,nil,1);
    end;
  end;
  FreeAndNil(Table_Name_Strings); //.Free;
  FreeAndNil(Log_Strings) ;//.Free;
end.


