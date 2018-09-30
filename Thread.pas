unit Thread;

interface

uses


    Classes, Math,Controls,Winapi.ShellAPI,Winapi.ActiveX,Winapi.Windows,Variants,
    SysUtils, Generics.Collections, Generics.Defaults,IdURI, IdHTTP,xmldom,XMLIntf,Xml.Win.msxmldom,Dialogs,XMLDoc,MSHTML,
    OleCtrls,IdIOHandlerSocket,IdSSLOpenSSL,SyncObjs,SelDelBox,System.Win.ComObj,Graphics,SHDocVw, System.JSON,
    System.JSONConsts, System.typinfo;
type
  TUserInfo= record
    IName, FName, Pass, Domens, Status, Maillists : string;
    class function Create(const IName, FName: string): TUserInfo; static;
    class function CreateFull(const IName, FName, Pass, Domens, Maillists, Status: string ): TUserInfo; static;
  end;

  TName = record
  first, last, middle : string;
  class function Create(const first, last, middle: string): TName;static;
  end;

  TMailResource = class
  private
    Fnickname: string;
    Fid: string;
    Fdismissed: string;
    Fname: TName;
    Fbirthday: TDateTime;
    Fgender: string;
    FAliases: TStringList;
    Fmaillist: string;
    Fsuccess:string;
    Ferror:string;
    Ftotal:Integer;
    procedure ParseJSON(Value: TJSONValue);
    procedure ParseJSONName(Value: TJSONValue);

public
    FJSONObject: TJSONObject;
    property Nickname: string read Fnickname;
    property Id: string read Fid;
    property is_dismissed: string read Fdismissed;
    property Name: TName read Fname;
    property birthday: TDateTime read Fbirthday;
    property gender: string read Fgender;
    property Aliases: TStringList read FAliases;
    property maillist: string read Fmaillist;
    property success:string read Fsuccess;
    property error:string read Ferror;
    property total:Integer read Ftotal;
  end;

type
    Start = class(TThread)
    XMLDocument1: IXMLDocument;
    IdHTTP1: TIdHTTP;
    Webbrowser1:TWebBrowser;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;

  private
    FJSONObject: TJSONObject;
   // function CreateJson(Params:TStringList):TJSONObject;
    function GETCommand( command:string; Method: string; Params: TStringList=nil;JSON:TStringStream = nil): string;
    function request_List(domain: string; page:string;per_page:string): string;
    function request_Get_oauth_token(domain: string; login:string): string;
    function request_del(domain: string; login:string): string;
    function request_edit(domain: string; login:string; password :string=''; iname :string='';fname : string='';
    enabled : string=''; hintq: string='';hinta:string=''): string;
    function request_add(domain: string; login:string; password :string; name:TName): string;
    function request_ml_List(domain: string): string;
    function request_ml_add(domain: string; maillist:string): string;
    function request_ml_subscribers(domain: string; maillist:string):string;
    function request_ml_subscribe(domain: string; maillist:string;subscriber:string):string;
    procedure EnumArray(main:string;list:TStringlist);
    procedure Memo;
    procedure WriteMemo(text:String;LineColor :TColor);
    procedure GetWin32_UserAccountInfo;
    procedure GetYaLogins;
    procedure OpenMailBox(logname:String);
    procedure CreateMailBox;
    procedure CreateMaillist;
    procedure Registration(list:TStringList);
    procedure DelMailBox;
    procedure SelDelBoxes;
    procedure Unlockbutton;
    procedure Sync;
    procedure Save;
    procedure Restore;
    procedure Enter;
    procedure Reload;
    procedure ListUser;
    procedure GetMaillist;
    procedure LoadFromExcel;
    procedure ChangePass(list:TStringList);
    protected
    procedure Execute; override;
  end;

implementation


uses Main;


const
 { URL для доступа к методам АПИ }
  cApiURL = 'https://api.directory.yandex.net/%s/%s'; //'https://pddimp.yandex.ru/api2/admin/email/%s';
  cVersion = 'v6';

var
 hostname, vdomen,vtoken, vpass,vlog, vText, vBtnCap:String;
 maillists,loglist,dellist:TStringList;
 CriticalSection:TCriticalSection;
 vAct:Integer;
 UserDirectory,ResUserDirectory: TDictionary<string, TUserInfo>;
 Mail:TMailResource;
{ Start }

procedure Start.SelDelBoxes;
var i:integer;
rez :TModalResult;
begin

 Form2.clbDel.Items:=dellist;
 rez:=Form2.ShowModal;
 if rez=mrOK then
 for i := 0 to Form2.clbDel.Items.Count-1 do
 if not Form2.clbDel.Checked[i] then
    dellist.Delete(dellist.Indexof(Form2.clbDel.Items[i]));
 if rez=mrCancel then dellist.Clear;

end;

procedure Start.Unlockbutton;
begin

 if YaConnect.btnStart1.Enabled then
 YaConnect.btnStart1.Caption:=vBtnCap;
 if vBtnCap<>'Continue' then
 begin
 YaConnect.btnStart1.Enabled:=true;
 YaConnect.btnStop1.Enabled:=true;
 end;
end;

procedure  Start.GetWin32_UserAccountInfo;
const
  WbemUser            ='';
  WbemPassword        ='';
  WbemComputer        ='localhost';
  wbemFlagForwardOnly = $00000020;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  vFullName,vFname,vIName     : String;
  fio : TStringList;
 begin;

  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer(WbemComputer, 'root\CIMV2', WbemUser, WbemPassword);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_UserAccount WHERE SIDTYPE=1 AND status=''OK''','WQL',wbemFlagForwardOnly);
  oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  // Создаем справочник
  UserDirectory := TDictionary<string, TUserInfo>.Create;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do
  begin
  if Terminated then break;
  vFullName:= String(FWbemObject.FullName);

    if vFullName<>'' then
    begin
    fio:=Tstringlist.Create;
    fio.text:=stringReplace(vFullName,' ',#13#10,[rfReplaceAll]);//мы заменяем все ; на символы конца строки
    vFname:= fio[0];
    vIName:= fio[1];
    end;

    if not((pos('_',String(FWbemObject.Name))=1) or
     ((ANSIUpperCase((String(FWbemObject.Name))[1])>char(127)) or (String(FWbemObject.Name)='ASWM USER')))
    then
    UserDirectory.Add(UpperCase(String(FWbemObject.Name)),TUserInfo.Create(vIname,vFname));
    FWbemObject:=Unassigned;
  end;

end;

procedure Start.WriteMemo(text:String;LineColor :TColor);
begin
 vText:=text;
 vcolor:=LineColor;
end;

class function TName.Create(const first,
  last, middle: string): TName;
begin
  Result.first := first;
  Result.last := last;
  Result.middle := middle;
end;

class function TUserInfo.Create(const IName,
  FName: string): TUserInfo;
begin
  Result.IName := IName;
  Result.FName := FName;
end;

class function TUserInfo.CreateFull(const IName,
  FName,Pass,Domens,Maillists,Status: string): TUserInfo;
begin
  Result.IName := IName;
  Result.FName := FName;
  Result.Pass := Pass;
  Result.Domens := Domens;
  Result.Status := Status;
  Result.Maillists := Maillists;
end;

(*
  procedure Start.CreateJson(Params:TStringList);
  var JSONObject, InnerObject : TJSONObject;
      Pair : TJSONPair;
      JsonArray: TJSONArray;
      i:Integer;
  begin
    {создали объект верхнего уровня}
    try
    JSONObject:=TJSONObject.Create;
    {создали пару}
    Pair:=TJSONPair.Create('stringPair','TestString');
    {записали пару в объект}
    JSONObject.AddPair(Pair);
    {можно записать и так}
    JSONObject.AddPair(TJSONPair.Create('stringPair2','TestString2'));
    {и так}
    JSONObject.AddPair(TJSONPair.Create('NumberPair2',TJSONNumber.Create(128)));

    {добавляем в объект ещё один объект}
    InnerObject:=TJSONObject.Create;
    InnerObject.AddPair('Inner_StringPair','Hello World!');
    JSONObject.AddPair(TJSONPair.Create('InnerObject',InnerObject));

    {записываем в объект массив}
    //создали пустой массив
    JsonArray:=TJSONArray.Create();
    JsonArray.Add('StringElement');//первый элемент - строка
    JsonArray.Add(True); //второй элемент - True
    JsonArray.Add(100500); //третий элемент - число
    //четвертый элемент - объект
    JsonArray.AddElement(TJSONObject.Create(TJSONPair.Create('PairName','PairValue')));

    //записали массив в объект
    JSONObject.AddPair('Array',JsonArray);

    //сохраняем объект в файл
    result:=JSONObject.ToString;
  finally
    JSONObject.Destroy;
  end;
  end;
*)

procedure TMailResource.ParseJSONName(Value: TJSONValue);
var
  JObject: TJSONObject;
  JPair: TJSONPair;
  i:integer;
  MemberName: string;
begin
  if not Assigned(Value) then  Exit;
  JObject := (Value as TJSONObject);//привели значение пары к классу TJSONObject
  try
    {проходим по каждой паре}
    for I := 0 to JObject.Count-1 do
      begin
        JPair:=JObject.Pairs[i];//получили пару по её индексу
        MemberName:=JPair.JsonString.Value;//определили имя
        {ищем в какое свойство записывать значение}
        if MemberName='first' then
        Fname.first:=JPair.JsonValue.Value
        else if MemberName='middle' then
        Fname.middle:=JPair.JsonValue.Value
        else if MemberName='last' then
        Fname.last:=JPair.JsonValue.Value
      end;

  except
    raise Exception.Create('Ошибка разбора JSON');
  end;
end;

procedure TMailResource.ParseJSON(Value: TJSONValue);
var
  JObject: TJSONObject;
  JPair: TJSONPair;
  i:integer;
  MemberName: string;
begin
  if not Assigned(Value) then  Exit;
  JObject := (Value as TJSONObject);//привели значение пары к классу TJSONObject
  try
    {проходим по каждой паре}
    for I := 0 to JObject.Count-1 do
      begin
        JPair:=JObject.Pairs[i];//получили пару по её индексу
        MemberName:=JPair.JsonString.Value;//определили имя
        {ищем в какое свойство записывать значение}
        if MemberName='nickname' then
         Fnickname:=JPair.JsonValue.Value
          else if MemberName='accounts' then
          Fid:=JPair.JsonValue.Value
           else if MemberName='success' then
            Fsuccess:=JPair.JsonValue.Value
            else if MemberName='error' then
             Ferror:=JPair.JsonValue.Value
              else if MemberName='total' then
               Ftotal:=StrToInt(JPair.JsonValue.Value)
               else if MemberName='name' then
                ParseJSON(JPair.JsonValue)
                 else if MemberName='maillist' then
                 Fmaillist:=JPair.JsonValue.Value;
      end;

  except
    raise Exception.Create('Ошибка разбора JSON');
  end;
end;


function Start.GETCommand( command:string; Method: string; Params: TStringList=nil;JSON:TStringStream = nil): string;
var
  URL,ParamsStr: string;
  resString:TStringStream;

begin
    resString:=TStringStream.Create;

    URL := Format(cApiURL, [cVersion,command]);

  try
       idHTTP1:= TIdHTTP.Create(nil);
       IdSSLIOHandlerSocketOpenSSL1:=tIdSSLIOHandlerSocketOpenSSL.Create(NIL);
       idHTTP1.IOHandler:=IdSSLIOHandlerSocketOpenSSL1;
       idHTTP1.HandleRedirects:=true;
       idHTTP1.Request.CustomHeaders.Add('Authorization: OAuth '+vtoken);
       idHTTP1.Request.CustomHeaders.Add('Accept: application/json');

       if (Params <> nil)  then
        begin
        Params.Delimiter := '&';
        ParamsStr := Params.DelimitedText;
        end;

       if Method='GET' then
       begin
       URL := URL + '?' + ParamsStr;
       result:=idHTTP1.Get(TIdUri.URLEncode(URL));
       end
       else if Method='POST' then
       result:=idHTTP1.Post(TIdUri.URLEncode(URL),JSON)
       else if Method='PATCH' then
       begin
       idHTTP1.Patch(TIdUri.URLEncode(URL),JSON,resString);
       result:= resString.DataString;
       end;



   finally
    idHTTP1.Free;
  end;
end;


function Start.request_List(domain: string; page:string;per_page:string): string;
var  Params: TStringList;
ParamStream : TStringStream;
ParamsStr: string;
begin
  Params:=TStringList.Create;
  try
    Params.Add('fields=nickname');
    Params.Add('is_robot=false');
    if Length(page)>0 then
      Params.Add('page='+page);
    if Length(per_page)>0 then
      Params.Add('per_page='+per_page);

    Result:=GETCommand('users/','GET',Params,nil);
  finally
    Params.Free;
  end;
end;

function Start.request_Get_oauth_token(domain: string; login:string): string;
var  Params: TStringList;
begin
  Params:=TStringList.Create;
  try
    if Length(domain)>0 then
      Params.Add('domain='+domain);
    if Length(login)>0 then
      Params.Add('login='+login);
    //Result:=GETCommand('get_oauth_token',Params,'POST');
  finally
    Params.Free;
  end;
end;

function Start.request_add(domain: string; login:string; password :string; name:TName): string;
var  Params: TStringStream;
     JSONObject,InnerObject : TJSONObject;
begin
  Params:=TStringStream.Create;
  JSONObject:=TJSONObject.Create;
  try
     JSONObject.AddPair(TJSONPair.Create('department_id','1'));
     if Length(password)>0 then
       JSONObject.AddPair(TJSONPair.Create('password',password));
     if Length(login)>0 then
       JSONObject.AddPair(TJSONPair.Create('nickname',login));

       {добавляем в объект ещё один объект}
     InnerObject:=TJSONObject.Create;
     InnerObject.AddPair('first',name.first);
   //InnerObject.AddPair('middle',name.middle);
     InnerObject.AddPair('last',name.last);
     JSONObject.AddPair(TJSONPair.Create('name',InnerObject));


     Params := TStringStream.Create(JSONObject.ToString, TEncoding.UTF8,False);


    Result:=GETCommand('users/','POST',nil,Params);
  finally
    Params.Free;
  end;
end;

function Start.request_edit(domain: string; login:string; password :string=''; iname :string='';fname : string='';
 enabled : string=''; hintq: string='';hinta:string=''): string;
var  Params: TStringList;
begin
  Params:=TStringList.Create;
  try
    if Length(domain)>0 then
      Params.Add('domain='+domain);
    if Length(login)>0 then
      Params.Add('login='+login);
    if Length(password)>0 then
      Params.Add('password='+password);
    if Length(iname)>0 then
      Params.Add('iname='+iname);
    if Length(fname)>0 then
      Params.Add('fname='+fname);
    if Length(enabled)>0 then
      Params.Add('enabled='+enabled);
    if Length(hintq)>0 then
      Params.Add('hintq='+hintq);
    if Length(hinta)>0 then
      Params.Add('hinta='+hinta);
    //Result:=GETCommand('edit',Params,'POST');
  finally
    Params.Free;
  end;
end;

function Start.request_del(domain: string; login:string): string;
var  Params: TStringList;
begin
  Params:=TStringList.Create;
  try
    if Length(domain)>0 then
      Params.Add('domain='+domain);
    if Length(login)>0 then
      Params.Add('login='+login);
    //Result:=GETCommand('del',Params,'POST');
  finally
    Params.Free;
  end;
end;

function Start.request_ml_List(domain: string): string;
var  Params: TStringList;
begin
  Params:=TStringList.Create;
  try
    if Length(domain)>0 then
      Params.Add('domain='+domain);
    //Result:=GETCommand('ml/list',Params,'GET');
  finally
    Params.Free;
  end;
end;

function Start.request_ml_add(domain: string; maillist:string): string;
var  Params: TStringList;
begin
  Params:=TStringList.Create;
  try
    if Length(domain)>0 then
      Params.Add('domain='+domain);
    if Length(maillist)>0 then
      Params.Add('maillist='+maillist);
    //Result:=GETCommand('ml/add',Params,'POST');
  finally
    Params.Free;
  end;
end;

function Start.request_ml_subscribers(domain: string; maillist:string): string;
var  Params: TStringList;
begin
  Params:=TStringList.Create;
  try
    if Length(domain)>0 then
      Params.Add('domain='+domain);
    if Length(maillist)>0 then
      Params.Add('maillist='+maillist);
   // Result:=GETCommand('ml/subscribers',Params,'GET');
  finally
    Params.Free;
  end;
end;


function Start.request_ml_subscribe(domain: string; maillist:string;subscriber:string): string;
var  Params: TStringList;
begin
  Params:=TStringList.Create;
  try
    if Length(domain)>0 then
      Params.Add('domain='+domain);
    if Length(maillist)>0 then
      Params.Add('maillist='+maillist);
    if Length(subscriber)>0 then
      Params.Add('subscriber='+subscriber);
    //Result:=GETCommand('ml/subscribe',Params,'POST');
  finally
    Params.Free;
  end;
end;

procedure Start.GetYaLogins;
var ttl,count:Integer;
html,vpage:String;

begin
loglist:=TStringList.Create;
vpage:='1';
//читаем файл
html:=request_List(vdomen,vpage,'5');
//пробуем распарсить
try
FJSONObject:=TJSONObject.ParseJSONValue(html) as TJSONObject;
Mail:=TMailResource.Create;
Mail.ParseJSON(FJSONObject);
ttl:=Mail.total;
count:=0;
  repeat
    if not Terminated then
    begin
        count:=count+1;
        ttl:=ttl-100;
        vpage:=InttoStr(count);
        count:=count+1;
        //читаем файл
        html:=request_List(vdomen,vpage,'100');
        //пробуем распарсить
        FJSONObject:=TJSONObject.ParseJSONValue(html) as TJSONObject;
        if Assigned(FJSONObject) then //парсинг прошел успешно - считываем login
          EnumArray('result',loglist)
        else //парсинг провалился - выводим сообщение
          raise Exception.Create('Файл не содержит JSON-данные');

    end;
  until (ttl<=0) ;
finally
  Mail.Free;
end;
end;

procedure Start.GetMaillist;
var
html:String;

begin
        maillists:=TStringList.Create;
        //читаем файл
        html:=request_ml_List(vdomen);
        //пробуем распарсить
        FJSONObject:=TJSONObject.ParseJSONValue(html) as TJSONObject;
        if Assigned(FJSONObject) then //парсинг прошел успешно - считываем login
          EnumArray('maillists',maillists)
        else //парсинг провалился - выводим сообщение
          raise Exception.Create('Файл не содержит JSON-данные');


end;



procedure Start.EnumArray(main:string;list:TStringlist);
var JsonArray: TJSONArray;
    Enum: TJSONArrayEnumerator;
    Mail : TMailResource;
begin
  {получаем масиив из значения пары с названием "items"}
  JsonArray:=FJSONObject.Get(main).JsonValue as TJSONArray;
  {создаем перечислитель}
  Enum:=TJSONArrayEnumerator.Create(JsonArray);
  {перебираем все элементы массива}
  while Enum.MoveNext do
    begin
    if  main='subscribers' then
      list.Add(UpperCase(Enum.GetCurrent.Value))
      else
      begin

      Mail:=TMailResource.Create;
      {передали значение в метод для парсинга}
      Mail.ParseJSON(Enum.GetCurrent);
      if main='maillists' then
      list.Add(UpperCase(Mail.maillist))
      else if  main='result' then
      list.Add(UpperCase(Mail.nickname)) ;


      Mail.Free;
    end;
end;
end;





procedure Start.Memo;
var countline:integer;
begin
if not Terminated then
begin
countline:=YaConnect.meOut.Lines.Add(vText);
YaConnect.ColorLine(YaConnect.meOut,countline,vColor);
YaConnect.meOut.setFocus;
YaConnect.btnEnter2box.setFocus;
end;
end;

procedure  Start.ChangePass(list:TStringList);
var
vpassword,response,vlog1:String;
i:Integer;
begin
       for i := 0 to list.Count-1 do
try
       if Terminated then break;
       if  pos('@',list[i])<>0  then
         vlog1:= copy(list[i],1,pos('@',list[i])-1)
         else vlog1:= list[i];
       vpassword:=Stringreplace(vpass,'$LOGIN',vlog1,[rfReplaceAll]);
       response:= Request_edit(vdomen,vlog1,vpassword);
       FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
       try
       Mail:=TMailResource.Create;
       Mail.ParseJSON(FJSONObject);
       if Mail.success='ok' then
       begin
       WriteMemo(vlog1+':пароль успешно отредактирован!',clYellow );
       Synchronize(Memo);
       end
       else
       begin
       WriteMemo(vlog1+':пароль не отредактирован!Error:'+Mail.error,clRed);
       Synchronize(Memo);
       end;
       finally
         Mail.Free;
       end;
except  on E : Exception do
begin
WriteMemo(E.message,clRed);
Synchronize(Memo);
end;
end;
end;

procedure Start.LoadFromExcel;
var i,j: integer;
iname, fname,pass,status, maillists,domens:String;
fio :Tstringlist;
begin

   //Синхронизируем рассылки
    ResUserDirectory := TDictionary<string, TUserInfo>.Create;
    ResUserDirectory.Clear;
    for I := 1 to YaConnect.sgMaillist.rowcount-1 do
    begin
    domens:= YaConnect.sgMailList.Cells[1,i];
    status:= YaConnect.sgMailList.Cells[2,i];
    ResUserDirectory.add(YaConnect.sgMailList.Cells[0,i],TUserInfo.CreateFull('','','',domens,'',status)) ;
    end;
    CreateMaillist;
   //Синхронизируем пользователей
   ResUserDirectory.Clear;
   for I := 1 to YaConnect.sgUsers.rowcount-1 do
    begin
    fio:=Tstringlist.Create;
    fio.text:=stringReplace(YaConnect.sgUsers.Cells[1,i],' ',#13#10,[rfReplaceAll]);//мы заменяем все ; на символы конца строки
    iname:= fio[1];
    fname:=fio[0];
    pass:= YaConnect.sgUsers.Cells[2,i];
    domens:= YaConnect.sgUsers.Cells[4,i];
    maillists:= YaConnect.sgUsers.Cells[5,i];
    status:=YaConnect.sgUsers.Cells[6,i];
    ResUserDirectory.add(YaConnect.sgUsers.Cells[0,i],TUserInfo.CreateFull(iname,fname,pass,domens,maillists,status)) ;
    end;

   CreateMailBox;

end;

procedure Start.CreateMailBox;
  var reg:Tstringlist;
   i,j:Integer;
   vpassword, verror,verrtext, vlog1, viname, vfname, vstatus, response :String;
   domens, maillists : TStringList;
   pair1: TPair<string, TUserInfo>;
begin
 verrtext:='';
 verror:='';
 //Проходим по списку пользователей
  for pair1  in ResUserDirectory do
  try
       if Terminated then break;
       //ищем значение в словаре
       vlog1:= pair1.key;
       vpassword:= pair1.Value.pass;
       //Если пароль не задан, то берем по шаблону из настроек
       if vpassword='' then
       vpassword:=Stringreplace(vpass,'$LOGIN',vlog1,[rfReplaceAll]);
       //Список рассылок
       maillists:=TStringList.Create;
       maillists.Text:= pair1.Value.Maillists;
       //Список доменов
       domens:=TStringList.create;
       domens.Text:= pair1.Value.Domens;
       //Если домен не задан для пользователя, то берем выбранный нна вкладке "Команды"
       if domens.Text='' then
         domens.Text:=vdomen;
       //Проходим по списку доменов для пользователя
       for I := 0 to domens.count-1 do
       try
       verrtext:='';
       verror:='';
       vdomen:= domens[i];
       //Добавляем пользователя
       YaConnect.DomenDirectory.TryGetValue(vdomen,vtoken);
       response:=Request_add(vdomen,vlog1,vpassword,TName.Create(pair1.Value.FName,pair1.value.IName,''));
       FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
         try
         Mail:=TMailResource.Create;
         Mail.ParseJSON(FJSONObject);
         //Если успешно добавлен
         if Mail.success='ok' then
           begin
           WriteMemo(vlog1+'@'+vdomen+':почтовый ящик успешно зарегистрирован!',clYellow);
           Synchronize(Memo);
           end
         else
         //Если ошибки
           begin
           verror:= Mail.error;
           if Mail.error='occupied' then
           begin
           verrtext:=vlog1+'@'+vdomen+':почтовый ящик с таким логином уже был зарегистрирован ранее!' ;
           WriteMemo(verrtext,clGreen);
           Synchronize(Memo);
           end
           else
           begin
           verrtext:=vlog1+'@'+vdomen+':почтовый ящик не зарегистрирован!Error:'+verror;
           WriteMemo(verrtext,clRed);
           Synchronize(Memo);
           end;
           end;
         finally
           Mail.Free;
         end;
         //Если нет ошибок или пользователь уже был зарегистрирован, то редактируем
       if (verror='') or (verror='occupied') then
        begin
           viname:=pair1.Value.iname;
           vfname:= pair1.Value.Fname;
           vstatus:= pair1.Value.Status;
           if vfname='' then
           vfname:='Не указано';
           if viname='' then
           viname:= 'Не указано';
           Randomize;
           response:= Request_edit(vdomen,vlog1,vpassword,viname,vfname,vstatus,'Какое второе название нашей компании?', 'А вот нет его!!! '+InttoStr(RandomRange(1000000,9999999)));
           FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
               try
               Mail:=TMailResource.Create;
               {передали значение в метод для парсинга}
               Mail.ParseJSON(FJSONObject);
               if Mail.success='ok' then
               begin
               WriteMemo(vlog1+'@'+vdomen+':почтовый ящик успешно отредактирован!',clYellow );
               Synchronize(Memo);
               end
               else
               begin
               WriteMemo(vlog1+'@'+vdomen+':почтовый ящик не отредактирован!',clRed);
               Synchronize(Memo);
               end;
               finally
                 Mail.Free;
               end;
           //Подписываем на  рассылки

           for j := 0 to maillists.count-1 do
           begin
           response:=request_ml_subscribe(vdomen,maillists[j],vlog1);
           FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
           try
           Mail:=TMailResource.Create;
           {передали значение в метод для парсинга}
           Mail.ParseJSON(FJSONObject);
           if Mail.success='ok' then
           begin
           WriteMemo('Пользователь '+vlog1+' подписан на рассылку '+maillists[j]+'!',clYellow);
           Synchronize(Memo);
           end
           else
           begin
           if Mail.error='already_subscribed' then
            begin
              WriteMemo('Пользователь '+vlog1+' уже был подписан на рассылку '+maillists[j]+'!',clGreen);
              Synchronize(Memo);
            end
            else
            begin
            WriteMemo('Пользователь '+vlog1+' не подписан на рассылку '+maillists[j]+'!Ошибка '+Mail.error,clRed);
            Synchronize(Memo);
            end;

           end;
           finally
             Mail.Free;
           end;
           end;

           //Если учетная запись не заблокирована, то проходим дорегистрацию
           if vstatus<>'no' then
            begin
           reg:=TStringList.Create;
           reg.Add(vlog1);
           Registration(reg);
            end;
            end;

           except  on E : Exception do
           begin
           WriteMemo(E.message,clRed);
           Synchronize(Memo);
           end;
        end;



   except  on E : Exception do
   begin
   WriteMemo(E.message,clRed);
   Synchronize(Memo);
   end;
   end;





end;





procedure Start.CreateMaillist;
 var verror, verrtext, vlog1, response:String;
 pair1: TPair<string, TUserInfo>;
 domens:TStringList;
 i:Integer;
begin
  verrtext:='';
  verror:='';
  for pair1 in ResUserDirectory do
  try
      if Terminated then break;
      //ищем значение в словаре
       vlog1:= pair1.key;
       domens:=TStringList.create;
       domens.Text:= pair1.Value.Domens;
       //Если домен не задан для рассылки, то берем выбранный нна вкладке "Команды"
       if domens.Text='' then
         domens.Text:=vdomen;
       //Проходим по списку доменов для рассылки
       for I := 0 to domens.count-1 do
       try
       verrtext:='';
       verror:='';
       //Добавляем рассылку
       vdomen:=  domens[i];
       YaConnect.DomenDirectory.TryGetValue(vdomen,vtoken);
       response:=Request_ml_add(vdomen,vlog1);
       FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
       try
       Mail:=TMailResource.Create;
       Mail.ParseJSON(FJSONObject);

       if Mail.success='ok' then
       begin
       WriteMemo(vlog1+'@'+domens[i]+':рассылка успешно зарегистрирована!',clYellow);
       Synchronize(Memo);
       end
       else
       begin
       verror:=Mail.error;
       if verror='occupied' then
       begin
       verrtext:=vlog1+'@'+vdomen+':рассылка с таким логином уже была зарегистрирована ранее!' ;
       WriteMemo(verrtext,clGreen);
       Synchronize(Memo);
       end
       else
       begin
       verrtext:=vlog1+'@'+vdomen+':рассылка не зарегистрирована!Error:'+verror;
       WriteMemo(verrtext,clRed);
       Synchronize(Memo);
       end;
       end;
       finally
         Mail.Free;
       end;
       except  on E : Exception do
        begin
        WriteMemo(E.message,clRed);
        Synchronize(Memo);
        end;
        end;
except  on E : Exception do
begin
WriteMemo(E.message,clRed);
Synchronize(Memo);
end;
end;
end;

procedure Start.OpenMailBox(logname:String);
  var
  l,vUrl,response,temptoken:String;
  Mail:TMailResource;
  begin
         if  pos('@',logname)<>0  then l:= copy(logname,1,pos('@',logname)-1)
         else l:= logname;
         //Запрос на получение временного токена
         response:=request_get_oauth_token(vdomen,l);
         FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
         Mail:=TMailResource.Create;
         {передали значение в метод для парсинга}
         Mail.ParseJSON(FJSONObject);

         if Mail.success='ok' then
         begin
         //Парсим ответ, получаем временный токен
        // temptoken:= Mail.temptoken;
        // WriteMemo('Логин '+l+': временный токен '+temptoken,clYellow);
         Synchronize(Memo);
         vUrl:= TIdURI.URLEncode('https://passport.yandex.ru/passport?mode=oauth&type=trusted-pdd-partner&error_retpath=mail.yandex.ru&access_token='{+temptoken});
         ShellExecute( Handle, 'open',PChar(vUrl),'-private-window ',  nil, SW_RESTORE );
         end;
         Mail.Free;
  end;






  procedure Start.Registration(list:TStringList);
  var i:Integer;
  l,vUrl,ready,response:String;
  Mail:TMailResource;
  begin
   //Дорегистрация
    for i := 0 to list.Count-1 do
         begin   // for
         if Terminated then break;
         if  pos('@',list[i])<>0  then
         l:= copy(list[i],1,pos('@',list[i])-1)
         else l:= list[i];
          //Проверяем, пройдена ли дорегистрация
         ready:=request_edit(vdomen,l);
         FJSONObject:=TJSONObject.ParseJSONValue(ready) as TJSONObject;
         FJSONObject:=FJSONObject.Get('account').JsonValue as TJSONObject;
         Mail:=TMailResource.Create;
         {передали значение в метод для парсинга}
         Mail.ParseJSON(FJSONObject);
       //  ready:=Mail.ready;
         Mail.Free;
         if ready='no' then
         //Запрос на получение временного токена
         try
         try
         response:=request_Get_oauth_token(vdomen,l);
         FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
         Mail:=TMailResource.Create;
         {передали значение в метод для парсинга}
         Mail.ParseJSON(FJSONObject);
         if Mail.success='ok' then
         begin
         //Парсим ответ, получаем временный токен
        // WriteMemo('Логин '+l+': временный токен '+Mail.temptoken,clYellow);
         Synchronize(Memo);
         vUrl:= TIdURI.URLEncode('https://passport.yandex.ru/passport?mode=oauth&type=trusted-pdd-partner&error_retpath=mail.yandex.ru&access_token='{+Mail.temptoken});
         ShellExecute( Handle, 'open',PChar(vUrl),'-private-window ',  nil, SW_RESTORE );

            vBtnCap:='Continue';
            Synchronize(Unlockbutton);
            Suspend;
            WriteMemo('Логин '+list[i]+': необходимо пройти регистрацию!',clYellow);
            Synchronize(Memo);
            WriteMemo('Логин '+list[i]+': регистрация успешно завершена!',clYellow);
            Synchronize(Memo);
            end
           else
         //сли ответ не содержит  'success', значит  временный токен не получен
         begin
         WriteMemo('Логин '+list[i]+' временный токен не получен!Ошибка:"'+Mail.error+'"',clRed);
         Synchronize(Memo);
         end;
         except  on E : Exception do
         begin
         WriteMemo(E.message,clRed);
         Synchronize(Memo);
         end;
         end;
         finally
           Mail.Free;
         end
         else
            //Если нет, значит дорегистрация уже пройдена
            begin
            WriteMemo('Логин '+list[i]+': регистрация уже пройдена!',clGreen);
            Synchronize(Memo);
            end;
         end;   // for
         if vAct=2 then
          begin
         WriteMemo('Дорегистрация завершена',clYellow);
         Synchronize(Memo);
          end;

         end;






procedure Start.DelMailBox;
var i:Integer;
response,vlog1:String;
begin

  dellist:= TStringList.Create;
   for i := 0 to loglist.Count-1 do
   begin
   vlog1:=copy(loglist[i],1,pos('@',loglist[i])-1);
    if not UserDirectory.ContainsKey(vlog1) then
      begin
      dellist.add(vlog1);
      end;
   end;
    SelDelBoxes;
    WriteMemo('[Список ящиков для удаления:] ',clLime);
    Synchronize(Memo);
    WriteMemo(dellist.Text,clYellow);
    Synchronize(Memo);
    WriteMemo('[Удаление выбранных почтовых ящиков:]',clLime);
    Synchronize(Memo);


    for i := 0 to dellist.Count-1 do
       begin
       vlog1:= dellist[i];
       response:=request_del(vdomen,vlog1);
       FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
       try
       Mail:=TMailResource.Create;
       {передали значение в метод для парсинга}
       Mail.ParseJSON(FJSONObject);
       if Mail.success='ok' then
       begin
       WriteMemo('Почтовый ящик '+vlog1+' удален!',clYellow);
       Synchronize(Memo);
       end
       else
       begin
       WriteMemo('Почтовый ящик '+vlog1+' невозможно удалить!Ошибка '+Mail.error,clRed);
       Synchronize(Memo);
       end;
       finally
         Mail.Free;
       end;
       end;
       WriteMemo('[Удаление ящиков завершено!]',clLime);
       Synchronize(Memo);
end;



procedure Start.Execute;
begin
vdomen:= YaConnect.FDomen;
vpass:= YaConnect.FPass;
vtoken:= YaConnect.FToken;
vlog:=YaConnect.Flog;
vText:=YaConnect.FText;
vBtnCap:=YaConnect.FBtnCap;
case YaConnect.getAct of
 0:begin
 vAct:=0;
 Save;
 end;
 1:begin
 vAct:=1;
 Restore;
 end;
 2:begin
 vAct:=2;
 Sync;
 end;
 3:begin
 vAct:=3;
 Enter;
 end;
 4:begin
 vAct:=4;
 Reload;
 end;
 5:begin
 vAct:=5;
 LoadFromExcel;
 end;
end;
end;

procedure Start.Restore;
var st,filelist,filelist2,subscribers:TStringList;
FRootFolder,iname,fname,html, vmaillist,response,ml:string;
  tsr : tsearchrec;
  i,j,k:integer;
begin
try
   FRootFolder:=ExtractFilePath(ParamStr(0));
   filelist:=TStringList.Create;
   subscribers:=TStringList.Create;
   WriteMemo('[ Восстановление почты для домена '+vdomen+ ' ]',clLime);
   Synchronize(Memo);
   if FindFirst(FRootFolder+vdomen+'/'+'m'+'/' + '*.inf',faAnyFile,tsr) = 0 then
    repeat
      filelist.Add(tsr.name);
    until FindNext(tsr) <> 0;
   FindClose(tsr);
   ResUserDirectory := TDictionary<string, TUserInfo>.Create;
   st:=TStringList.Create;
   WriteMemo('[ Восстановление почтовых ящиков и рассылок ]',clLime);
   Synchronize(Memo);
   for i :=0 to filelist.Count-1 do
   begin
   st.LoadFromFile(FRootFolder+vdomen+'/'+'m'+'/'+filelist[i]);
   html:= st.Text;
   FJSONObject:=TJSONObject.ParseJSONValue(html) as TJSONObject;
   FJSONObject:=FJSONObject.Get('account').JsonValue as TJSONObject;

   Mail:=TMailResource.Create;
   {передали значение в метод для парсинга}
   Mail.ParseJSON(FJSONObject);
  //iname:=Mail.iname;
  // fname:=Mail.fname;
   vmaillist:= Mail.maillist;
   Mail.Free;
   ResUserDirectory.Clear;
   ResUserDirectory.add(copy(filelist[i],1,pos('.inf',filelist[i])-1),TUserInfo.Create(iname,fname));
   if vmaillist='no' then
   //Создание почтовых ящиков
   CreateMailBox
   else
   //Создание рассылок
   CreateMaillist;
   end;
   //Подписать на рассылку
   filelist2:=TStringList.Create;
   if FindFirst(FRootFolder+vdomen+'/'+'ml'+'/' + '*.inf',faAnyFile,tsr) = 0 then
    repeat
      filelist2.Add(tsr.name);
    until FindNext(tsr) <> 0;
   FindClose(tsr);
   st.Clear;

   for j :=0 to filelist2.Count-1 do
   begin
      st.LoadFromFile(FRootFolder+vdomen+'\'+'ml'+'\'+filelist2[j]);
      html:= st.Text;
      FJSONObject:=TJSONObject.ParseJSONValue(html) as TJSONObject;
        if Assigned(FJSONObject) then //парсинг прошел успешно - считываем login
          EnumArray('subscribers',subscribers)
        else //парсинг провалился - выводим сообщение
          raise Exception.Create('Файл не содержит JSON-данные');
        WriteMemo('[ Восстановление подписок на рассылки ]',clLime);
        Synchronize(Memo);
        for k := 0 to subscribers.Count-1 do
        begin

       ml:=copy(filelist2[j],1,pos('.inf',filelist2[j])-1);
       response:=request_ml_subscribe(vdomen,ml,subscribers[k]);
       FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
       try
       Mail:=TMailResource.Create;
       {передали значение в метод для парсинга}
       Mail.ParseJSON(FJSONObject);
       if Mail.success='ok' then
       begin
       WriteMemo('Пользователь '+subscribers[k]+' подписан на рассылку '+ml+'!',clYellow);
       Synchronize(Memo);
       end
       else
       begin
       if Mail.error='already_subscribed' then
        begin
          WriteMemo('Пользователь '+subscribers[k]+' уже был подписан на рассылку '+ml+'!',clRed);
          Synchronize(Memo);
        end
        else
        begin
        WriteMemo('Пользователь '+subscribers[k]+' не подписан на рассылку '+ml+'!Ошибка '+Mail.error,clRed);
        Synchronize(Memo);
        end;

       end;
       finally
         Mail.Free;
       end;
        end;
   end;





finally
  WriteMemo('[ Восстановление в домене '+vdomen+' завершено! ]',clLime);
   Synchronize(Memo);
   vBtnCap:='Start';
   Synchronize(Unlockbutton);
end;

end;

procedure Start.ListUser;
begin
YaConnect.cbLog.Items.Clear;
YaConnect.cbLog.Items:=loglist;
YaConnect.cbLog.Items.EndUpdate;
end;

procedure Start.Reload;
begin
 try
  GetYaLogins;
  Synchronize(ListUser);
 except
  on E:Exception  do
  begin
    WriteMemo(E.Classname+E.Message,clRed);
    Synchronize(Memo);
  end;
 end;

end;

procedure Start.Enter;
begin
   try
   OpenMailbox(vlog);
   except
   on E:Exception  do
   begin
    WriteMemo(E.Classname+E.Message,clRed);
    Synchronize(Memo);
   end;
   end;
end;

procedure Start.Save;
var
  vRootFolder,vlog1, vml1:string;
  st: TStringList;
  i,j:integer;
begin

        try
        WriteMemo('[Сохранение данных в файл]',clLime);
        Synchronize(Memo);
        vRootFolder:=ExtractFilePath(ParamStr(0));
        if not DirectoryExists(vRootFolder+vdomen) then
               CreateDir(vRootFolder+vdomen);
        if not DirectoryExists(vRootFolder+vdomen+'\'+'m') then
               CreateDir(vRootFolder+vdomen+'\'+'m');
        if not DirectoryExists(vRootFolder+vdomen+'\'+'ml') then
               CreateDir(vRootFolder+vdomen+'\'+'ml');
        st:=TStringList.Create;
        //Получение списка почтовых ящиков
        getYaLogins;
        try
        for i := 0 to loglist.Count-1 do
         begin
        st.Clear;
        vlog1:= copy(loglist[i],1,pos('@',loglist[i])-1);
        st.text:=request_edit(vdomen,vlog1);
        st.SaveToFile(vRootFolder+vdomen+'\'+'m'+'\'+vlog1+'.inf');
        WriteMemo('[ Данные пользователя '+vlog1+' сохранены в файл ]',clLime);
        Synchronize(Memo);
         end;
        except
        on E:Exception do
        begin
        WriteMemo('[ Данные пользователя '+vlog1+' не удалось сохранить в файл ]',clRed);
        Synchronize(Memo);
        end;
        end;
        //Получение списка  подписчиков
        GetMaillist;
        try
         for j := 0 to maillists.Count-1 do
         begin

        st.Clear;
        vml1:=copy(maillists[j],1,pos('@',maillists[j])-1);
        st.text:=request_ml_subscribers(vdomen,vml1);
        st.SaveToFile(vRootFolder+vdomen+'\'+'ml'+'\'+vml1+'.inf');
        WriteMemo('[ Данные рассылки '+vml1+' сохранены в файл ]',clLime);
        Synchronize(Memo);
         end;
        except
        on E:Exception do
        begin
        WriteMemo('[ Данные рассылки '+vml1+' не удалось сохранить в файл ]',clRed);
        Synchronize(Memo);
        end;
         end;


        finally
        st.Free;
        WriteMemo('[ Сохранение завершено ]',clLime);
        Synchronize(Memo);
        vBtnCap:='Start';
        Synchronize(Unlockbutton);
        end;
end;

procedure Start.Sync;
var
     i:Integer;
     Buffer: array[0..Pred(MAX_COMPUTERNAME_LENGTH+1)] of Char;
     Size: cardinal;
     vlog1,vlog2,vlog3 : String;
     pair1: TPair<string, TUserInfo>;
begin
try
  CriticalSection:=TCriticalSection.Create;
  CriticalSection.Enter;
  Size := SizeOf(Buffer);
  hostname := Buffer;
  //Получение списка логинов WMI запросом (wmilogins)
  WriteMemo('[ Получение списка логинов на сервере '+hostname+' ]',clLime);
  Synchronize(Memo);
  GetComputerName(Buffer, Size);
  try
    CoInitialize(nil);
    try
      GetWin32_UserAccountInfo;
    finally
      CoUninitialize;
    end;
  except
    on E:EOleException do
    begin
       WriteMemo(Format('EOleException %s %x', [E.Message,E.ErrorCode]),clRed);
       Synchronize(Memo);
    end;
    on E:Exception do
    begin
       WriteMemo(E.Classname+E.Message,clRed);
       Synchronize(Memo);
    end;
  end;
  for vlog1 in UserDirectory.Keys do
   begin
   if  finished then break;
   WriteMemo(vlog1,clYellow);
   Synchronize(Memo);
   end;
   WriteMemo('[ Всего учетных записей на сервере '+hostname+':] '+InttoStr(UserDirectory.count),clLime);
   Synchronize(Memo);
   //Получение списка логинов в почтовом домене
   try
      WriteMemo('[ Получение списка логинов в почтовом домене '+vdomen+' ]:',clLime);
      Synchronize(Memo);
      GetYaLogins;
      WriteMemo('[ Отсортированный список учетных записей в домене '+vdomen+' ]: ',clLime);
      Synchronize(Memo);
      for i := 0 to loglist.Count-1 do
      begin
      WriteMemo(loglist[i],clYellow);
      Synchronize(Memo);
      end;
      WriteMemo('[ Всего учетных записей в почтовом домене '+vdomen+' ]: '+InttoStr(loglist.Count),clLime);
      Synchronize(Memo);
    except  on E : Exception do
    begin
    WriteMemo(E.message,clRed);
    Synchronize(Memo);
    end;
    end;
   //Получение списка почтовых ящиков, для которых нет логинов на сервере
   if visDel then
   Synchronize(DelMailBox);
   //Получение списка логинов, для которых нужно создать почтовые ящики
   ResUserDirectory := TDictionary<string, TUserInfo>.Create;
   for pair1  in UserDirectory    do
   begin
   vlog2:=pair1.key;
   if ((loglist.IndexOf(UPPERCASE(vlog2+'@'+vdomen))=-1) )  then
    begin
    ResUserDirectory.add(vlog2,TUserInfo.Create(pair1.Value.IName, pair1.Value.FName));
    end;
   end;
    if ResUserDirectory.Count=0 then
    begin
    WriteMemo('[ Для всех учетных записей в домене '+vdomen+' существуют почтовые ящики! ]',clLime);
    Synchronize(Memo)
    end
    else
    begin
    WriteMemo('[ Список учетных записей, для которых будут созданы почтовые ящики ]:',clLime);
    Synchronize(Memo);
    for vlog3 in ResUserDirectory.Keys do
    begin
    if  finished then break;
    WriteMemo(vlog3,clYellow);
    Synchronize(Memo);
    end;
    //Создание почтовых ящиков
     WriteMemo('[ Создание почтовых ящиков ]',clLime);
     Synchronize(Memo);
    CreateMailBox;
    // Registration(reslist);
    WriteMemo('[ Создание почтовых ящиков в домене '+vdomen+' завершено! ]',clLime);
    Synchronize(Memo);
    end;
    if visAll and (not Finished) then
    Registration(loglist);
    if visPass and (not Finished) then
    ChangePass(loglist);
    vBtnCap:='Start';
    Synchronize(Unlockbutton);
    finally
     CriticalSection.Leave;
end;
end;


end.
