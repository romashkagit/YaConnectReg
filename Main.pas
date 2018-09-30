unit Main;

interface

uses

  Winapi.Windows, Winapi.Messages,SysUtils,Variants,Classes,Graphics,Controls,Forms,Dialogs,StdCtrls,
  IdIOHandler, System.Win.ComObj,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,xmldom,XMLIntf,Xml.Win.msxmldom,
  XMLDoc,OleCtrls, SHDocVw,ExtCtrls,IniFiles,Winapi.ShellAPI,Winapi.ActiveX,IdURI,ComCtrls,
  Buttons,SelDelBox,Thread, Generics.Collections, Generics.Defaults, System.UITypes,
  Vcl.Grids;

type
  TYaConnectReg = class(TForm)
    Panel1: TPanel;
    edDom1: TEdit;
    lblDomen1: TLabel;
    edToken1: TEdit;
    lblToken1: TLabel;
    btnStart1: TButton;
    pnl1: TPanel;
    pnlTop: TPanel;
    Edit1: TEdit;
    lblUrl: TLabel;
    pnl2: TPanel;
    lblDomen2: TLabel;
    lblToken2: TLabel;
    edDom2: TEdit;
    edToken2: TEdit;
    pnl3: TPanel;
    lblDom3: TLabel;
    lblToken3: TLabel;
    edDom3: TEdit;
    edToken3: TEdit;
    meOut: TRichEdit;
    V: TPanel;
    cbLog: TComboBox;
    lblListDom: TLabel;
    lblListLogin: TLabel;
    btnEnter2box: TButton;
    cbAll: TCheckBox;
    cbDel: TCheckBox;
    btnStop1: TButton;
    btnRefresh: TSpeedButton;
    pnl4: TPanel;
    lblDom4: TLabel;
    lblToken4: TLabel;
    edToken4: TEdit;
    edDom4: TEdit;
    pcMain: TPageControl;
    tshParam: TTabSheet;
    tshComand: TTabSheet;
    rgAct: TRadioGroup;
    pnlProg: TPanel;
    pbRefresh: TProgressBar;
    lblProg: TLabel;
    pnl5: TPanel;
    lb1: TLabel;
    edPass: TEdit;
    cbPass: TCheckBox;
    tsExcel: TTabSheet;
    pnlload: TPanel;
    btnloadfrxl: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    sgDomens: TStringGrid;
    sgMailList: TStringGrid;
    sgUsers: TStringGrid;
    btnLoadYA: TButton;
    cbDomen: TComboBox;

    procedure FormCreate(Sender: TObject);
    procedure edDomChange(Sender: TObject);
    procedure edTokenChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCanResize(Sender: TObject; var NewWidth, NewHeight: Integer;
      var Resize: Boolean);
    procedure ColorLine(RE : TRichEdit; Line : Integer; LineColor : TColor);
    procedure btnEnter2boxClick(Sender: TObject);
    procedure cbDomenDropDown(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    procedure cbDomenMeasureItem(Control: TWinControl; Index: Integer;
      var Height: Integer);
    procedure cbDomenChange(Sender: TObject);
    procedure btnStart1Click(Sender: TObject);
    procedure btnStop1Click(Sender: TObject);
    procedure rgActClick(Sender: TObject);
    procedure edPassChange(Sender: TObject);
    procedure edPassClick(Sender: TObject);
    procedure cbPassClick(Sender: TObject);
    procedure cbDelClick(Sender: TObject);
    procedure loadfrxl;
    procedure btnLoadYAClick(Sender: TObject);




  public
    FDomen:String;
    FPass:String;
    FToken:String;
    Flog:String;
    FText:String;
    FBtnCap:String;
    DomenDirectory : TDictionary<string, string>;
    function  GetAct: Integer;
    procedure SetAct(Value: Integer);

  private
    FPaused:Boolean;
    FStop:Boolean;
    Start1:Start;
    Init1:TIniFile;
    FAct:Integer;

  end;
    
var
  YaConnect: TYaConnectReg;

  vcolor:TColor;
  visAll:boolean;
  visDel:boolean;
  visPass:boolean;
implementation

{$R *.dfm}


procedure TYaConnectReg.SetAct(Value:Integer);
begin
FAct:=value;
end;




function  TYaConnectReg.GetAct:Integer;
begin
Result:=FAct;
end;

procedure  TYaConnectReg.ColorLine(RE : TRichEdit; Line : Integer; LineColor : TColor);
begin
  with RE do
  begin
    SelStart := SendMessage(Handle, EM_LINEINDEX, Line, 0);
    SelLength := Length(Lines[Line]);
    if pos('rejected',Lines[Line])<>0 then
    SelAttributes.Color :=clRed
    else
    SelAttributes.Color := LineColor;
  end;
end;


procedure TYaConnectReg.btnEnter2boxClick(Sender: TObject);
begin
if cblog.Text='' then
  ShowMessage('Не указан логин!')
else
begin
Flog:=cbLog.Text;
Fdomen:=cbDomen.Text;
DomenDirectory.TryGetValue(FDomen,FToken);
setAct(3);
  if (Start1 = nil) or (Start1.finished) then
     begin
     //Вначале нужно создать экземпляр потока:
     Start1:=Start.Create(False);
     Start1.FreeOnTerminate:=true;
     //Параметр False запускает поток сразу после создания, True - запуск впоследствии , методом Resume
     //Далее можно указать параметры потока, например приоритет:
     Start1.Priority:=tpNormal;
     end;
  if Start1.Suspended then
     Start1.Resume;
end;

end;



procedure TYaConnectReg.btnStart1Click(Sender: TObject);
begin
Flog:=cbLog.Text;
FPass:=edPass.Text;
if cbAll.Checked then  visAll:=true
else visAll:=false;
if cbDel.Checked then   visDel:=true
else visDel:=false;
if cbPass.Checked then visPass:=true
else visPass:=false;
FDomen:=cbDomen.Text;
if FDomen='' then
ShowMessage('Не указано имя почтового домена!')
else
begin
DomenDirectory.TryGetValue(FDomen,FToken);
setAct(rgAct.ItemIndex);
if (Start1=nil) or (Start1.finished) then
 begin
    //Вначале нужно создать экземпляр потока:
    Start1:=Start.Create(True);
    Start1.FreeOnTerminate:=True;
    //Параметр False запускает поток сразу после создания, True - запуск впоследствии , методом Resume
    //Далее можно указать параметры потока, например приоритет:
    Start1.Priority:=tpNormal;
    Start1.Resume;
    btnStart1.Caption:='Pause';

 end
   else
 begin
    if (Start1.Suspended) then
    begin
     Start1.Resume;
     btnStart1.Caption:='Pause';
    end
    else
    begin
     Start1.Suspend;
     btnStart1.Caption:='Start' ;
    end;
 end;
end;

end;



procedure TYaConnectReg.btnStop1Click(Sender: TObject);
begin
  if not (Start1 = nil) then
  begin
    Start1.Terminate;
    if Start1.Suspended then
    Start1.Resume;
    btnStart1.Caption:='Start';
  end;
end;



procedure TYaConnectReg.loadfrxl;
var Rows, Cols, i,j: integer;
    ExcelApp, WorkSheet: OLEVariant;
    d: TDateTime;
    openDialog : TOpenDialog;
begin
 //openDialog := TOpenDialog.Create(self);
 try
  //if OpenDialog.Execute then

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.DisplayAlerts := False;
  //открываем книгу
  ExcelApp.Workbooks.Open(ExtractFilePath(Application.ExeName)+'users.xlsm');//openDialog.FileName);
  //получаем 1 лист
  WorkSheet:=ExcelApp.ActiveWorkbook.Sheets[1];
  //определяем количество строк и столбцов таблицы
  Rows:=WorkSheet.UsedRange.Rows.Count;
  Cols:=WorkSheet.UsedRange.Columns.Count;

  sgDomens.RowCount:=Rows;
  sgDomens.ColCount:=Cols;

  //выводим данные в таблицу
  for I := 0 to Rows-1 do
    for j := 0 to Cols-1 do
        sgDomens.Cells[J,I]:=WorkSheet.UsedRange.Cells[I+1,J+1];
  //end;
  //получаем 2 лист
  WorkSheet:=ExcelApp.ActiveWorkbook.Sheets[2];
  //определяем количество строк и столбцов таблицы
  Rows:=WorkSheet.UsedRange.Rows.Count;
  Cols:=WorkSheet.UsedRange.Columns.Count;

  sgMaillist.RowCount:=Rows;
  sgMaillist.ColCount:=Cols;

  //выводим данные в таблицу
  for I := 0 to Rows-1 do
    for j := 0 to Cols-1 do
        sgMaillist.Cells[J,I]:=WorkSheet.UsedRange.Cells[I+1,J+1];
  //получаем 3 лист
  WorkSheet:=ExcelApp.ActiveWorkbook.Sheets[3];
  //определяем количество строк и столбцов таблицы
  Rows:=WorkSheet.UsedRange.Rows.Count;
  Cols:=WorkSheet.UsedRange.Columns.Count;

  sgUsers.RowCount:=Rows;
  sgUsers.ColCount:=Cols;

  //выводим данные в таблицу
  for I := 0 to Rows-1 do
    for j := 0 to Cols-1 do
        sgUsers.Cells[J,I]:=WorkSheet.UsedRange.Cells[I+1,J+1];
  btnLoadYA.Enabled:=True;
 finally
 // openDialog.Destroy;
  ExcelApp.ActiveWorkbook.Close;
 end;


end;

procedure TYaConnectReg.btnLoadYAClick(Sender: TObject);
var i:Integer;
begin
setAct(5);
//Получаем список доменов из Excel
DomenDirectory := TDictionary<string, string>.Create;
for i :=1 to sgDomens.RowCount-1 do
if not DomenDirectory.ContainsKey(sgDomens.Cells[0,i]) then
DomenDirectory.add(sgDomens.Cells[0,i],sgDomens.Cells[1,i]);
//По умолчанию берем шаблон пароля из настроек
FPass:=edPass.Text;
//Запускаем обновление пользователей и рассылок в почтовом сервере
if (Start1 = nil) or (Start1.finished) then
     begin
     //Вначале нужно создать экземпляр потока:
     Start1:=Start.Create(False);
     Start1.FreeOnTerminate:=true;
     //Параметр False запускает поток сразу после создания, True - запуск впоследствии , методом Resume
     //Далее можно указать параметры потока, например приоритет:
     Start1.Priority:=tpNormal;
     end;
  if Start1.Suspended then
     Start1.Resume;
//Переходим на вкладку "Команды" с выводом лога
pcMain.ActivePage:=tshComand;
end;

procedure TYaConnectReg.cbDelClick(Sender: TObject);
begin
if  cbDel.Checked then
if messagedlg('Вы уверены, что нужно включить в эту опцию?',mtinformation,[mbyes,mbno],0)= mrYes then
   cbDel.checked:=true
   else
   cbDel.checked:=false;
end;

procedure TYaConnectReg.cbDomenChange(Sender: TObject);
begin
cbLog.Text:='';
cbLog.Items.Clear;
end;

procedure TYaConnectReg.cbDomenDropDown(Sender: TObject);
var s:string;
begin
 cbDomen.Items.Clear;
 for s in  DomenDirectory.Keys  do
 cbDomen.Items.add(s);
end;

procedure TYaConnectReg.cbDomenMeasureItem(Control: TWinControl;
  Index: Integer; var Height: Integer);
begin
  cbLog.Text:='';
  cbLog.Items.Clear;
end;



procedure TYaConnectReg.cbPassClick(Sender: TObject);
begin
if  cbPass.Checked then
if messagedlg('Вы уверены, что нужно включить в эту опцию?',mtinformation,[mbyes,mbno],0)= mrYes then
   cbPass.checked:=true
   else
   cbPass.checked:=false;
end;

procedure TYaConnectReg.edDomChange(Sender: TObject);
var vcaption, vtext:String;
begin
 vcaption:= TEdit(Sender).name;
 vtext:=TEdit(Sender).Text;
 Init1.WriteString('Domen info','Domen'+vcaption[length(vcaption)],vText);
end;


procedure TYaConnectReg.edPassChange(Sender: TObject);
begin
//CheckPass(edPass.Text);
Init1.WriteString('Password','Pass',edPass.Text);

end;



procedure TYaConnectReg.edPassClick(Sender: TObject);
begin
edPass.Font.Color:=clWindowText;
end;

procedure TYaConnectReg.edTokenChange(Sender: TObject);
var vcaption, vtext:String;
begin
vcaption:= TEdit(Sender).name;
vtext:=TEdit(Sender).Text;
Init1.WriteString('Domen info','Token'+vcaption[length(vcaption)],vtext);
end;



procedure TYaConnectReg.FormCanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
begin
Resize:=true;
end;

procedure TYaConnectReg.FormClose(Sender: TObject; var Action: TCloseAction);
begin

if not (Start1 = nil) then
  Start1.Terminate;
FPaused:=false;
FStop:=true;
FreeAndNil(DomenDirectory);
end;

procedure TYaConnectReg.FormCreate(Sender: TObject);
var i,j,k :Integer;
vdomen, vtoken: String;
 cb:TCheckBox;
begin
//Загрузка файла
loadfrxl;
lblProg.Visible:=false;
pbRefresh.Visible:=false;
//Чтение
FileSetAttr(ExtractFilePath(Application.ExeName)+'YaConnectReg.ini', faHidden);
Init1:= TIniFile.Create(ExtractFilePath(Application.ExeName)+'YaConnectReg.ini');
// Создаем справочник
DomenDirectory := TDictionary<string, string>.Create;
for i :=1 to 4 do
begin
vdomen:= init1.ReadString('Domen info','Domen'+InttoStr(i), '');
vtoken:= init1.ReadString('Domen info','Token'+InttoStr(i), '');
TEdit(FindComponent('edDom'+InttoStr(i))).Text:=vdomen;
TEdit(FindComponent('edToken'+InttoStr(i))).Text:=vtoken;
if not DomenDirectory.ContainsKey(vdomen) and (vdomen<>'') then
begin
sgUsers.cells[i,0]:= vdomen;
DomenDirectory.add(vdomen,vtoken) ;
end;
end;
{with sgUsers do
 begin
   Cells[0,0]:='Пользователь';
   sgUsers.ColCount:=DomenDirectory.Count+1;
   for k := 1 to ColCount-1 do
   for j := FixedRows to RowCount do
     Objects[k,j]:=TCheckBox.Create(sgUsers);
 end;}
edPass.Text:=init1.ReadString('Password','Pass', '$LOGIN_Cc');
end;




procedure TYaConnectReg.rgActClick(Sender: TObject);
begin
if rgAct.ItemIndex=2 then
begin
cbAll.Enabled:=true;
cbDel.Enabled:=true;
end
else
begin
cbAll.Enabled:=false;
cbDel.Enabled:=false;
end
end;


procedure TYaConnectReg.btnRefreshClick(Sender: TObject);
begin
 cbLog.Text:='';
 cbLog.Items.BeginUpdate;
 setAct(4);
 lblProg.Visible:=true;
 pbRefresh.Visible:=true;
 pbRefresh.Position:=0;
 FDomen:=cbDomen.Text;
 if FDomen='' then
  ShowMessage('Не указано имя почтового домена!')
 else
 begin
 DomenDirectory.TryGetValue(FDomen,Ftoken);
 if (Start1 = nil) or (Start1.finished) then
    begin
    //Вначале нужно создать экземпляр потока:
    Start1:=Start.Create(False);
    Start1.FreeOnTerminate:=true;
    //Параметр False запускает поток сразу после создания, True - запуск впоследствии , методом Resume
    //Далее можно указать параметры потока, например приоритет:
    Start1.Priority:=tpNormal;
    end;
 if Start1.Suspended then
   Start1.Resume;
 end;
 lblProg.Visible:=false;
 pbRefresh.Visible:=false;
end;





end.
