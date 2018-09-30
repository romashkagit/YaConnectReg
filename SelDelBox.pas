unit SelDelBox;

interface

uses

  Winapi.Windows,Winapi.Messages,SysUtils,Variants,Classes,Graphics,Controls,
  Forms,Dialogs,StdCtrls,CheckLst,Menus,ExtCtrls;

type
  TForm2 = class(TForm)
    clbDel: TCheckListBox;
    pmSel: TPopupMenu;
    miSelect: TMenuItem;
    miUnselect: TMenuItem;
    pnlSelTop: TPanel;
    lblSel: TLabel;
    pnlSelBot: TPanel;
    btnOK: TButton;
    btnCancel: TButton;
    procedure miSelectClick(Sender: TObject);
    procedure miUnselectClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}


procedure TForm2.btnOKClick(Sender: TObject);
begin
Close;
ModalResult:= mrOK;
end;

procedure TForm2.miSelectClick(Sender: TObject);
var i:integer;
begin
for i := 0 to clbDel.Items.Count-1 do
 clbDel.Checked[i]:=true;
end;

procedure TForm2.miUnselectClick(Sender: TObject);
var i:integer;
begin
for i := 0 to clbDel.Items.Count-1 do
 clbDel.Checked[i]:=false;
end;

end.

