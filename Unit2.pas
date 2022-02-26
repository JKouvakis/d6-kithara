unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ShellCtrls, FileCtrl, Grids, StrUtils, DB,
  ADODB;

type
  TForm2 = class(TForm)
    ProgressBar1: TProgressBar;
    Label1: TLabel;
    Button1: TButton;
    Tree: TShellTreeView;
    Names: TListBox;
    Files: TFileListBox;
    StringGrid1: TStringGrid;
    Button2: TButton;
    ADOQuery1: TADOQuery;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

uses Unit1;

{$R *.dfm}

procedure TForm2.Button1Click(Sender: TObject);
begin
  Form1.Enabled := True;
  Form2.Close;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
  Tree.Root:=ExtractFilePath(Application.ExeName)+'Data\';
  Tree.FullExpand;
end;

procedure TForm2.FormShow(Sender: TObject);
var
  i,j:integer;
  cpath:string;
begin
  for i:=1 to Tree.Items.Count-1 do
    begin
      cpath:=Tree.Folders[i].PathName;
      Files.ApplyFilePath(cpath);
      If Files.Items.Count>0 Then
        For j:=0 to Files.Items.Count-1 do
          begin
            Names.Items.Add(cpath+'\'+Files.Items.Strings[j]);
          end;
    end;
  ProgressBAr1.Max:=Names.Items.Count;
end;

function cleartags(s: string):string;
var
  new : string;
begin
  new:=S;
  repeat
    if pos('<',new)=1 then
      new := Midstr(new,pos('>',new)+1,length(new));
    If pos('<',new)>1 then
      new := Leftstr(new,pos('<',new)-1) + midstr(new,pos('>',new)+1,length(new));
  until pos('<',new)=0;
  Result:=new;
end;

procedure TForm2.Button2Click(Sender: TObject);
var
  i,j:integer;
  S,title,artist,poem,music,contain,filenm:string;
  tab,crd,getme:boolean;
  F: TextFile;
begin
  For i:=0 to Names.Items.Count-1 do
    begin
      tab:=false;crd:=false;getme:=false;
      title:='';artist:='';poem:='';music:='';filenm:='';
      contain:='0';
      ProgressBar1.Position:=i;
      AssignFile(F, Names.Items.Strings[i]);
      Reset(F);
      repeat
        ReadLn(F,S);
        if (getme=true) then
          begin
            music:=S;
            getme:=false;
          end;
        if (PoS('class=ti',S)<>0) or (Pos('class="ti"',S)<>0) then
          begin
            title:=S;
          end;
        if (PoS('class=ar',S)<>0) or (Pos('class="ar"',S)<>0) then
          begin
            artist:=S;
          end;
        if (PoS('class=cr',S)<>0) or (Pos('class="cr"',S)<>0) then
          begin
            music:=S;
            getme:=true;
          end;
        if (PoS('class=ch',S)<>0) or (Pos('class="ch"',S)<>0) then
          begin
            crd:=true;
          end;
        if (PoS('class=ta',S)<>0) or (Pos('class="ta"',S)<>0) then
          begin
            tab:=true;
          end;
      until EOF(F);

      StringGrid1.RowCount:=i;

      filenm:=Names.Items.Strings[i];
      filenm:=Midstr(filenm,length(ExtractFilePath(Application.ExeName)+'Data\')+1,length(filenm));
      title:=cleartags(title);
      artist:=cleartags(artist);
      music:=cleartags(music);
      if (pos('/',music)>0) then
        begin
          poem:=midstr(music,pos('/',music)+1,length(music));
          music:=leftstr(music,pos('/',music)-1);
        end
      else
        begin
          poem:=music
        end;
      If ((crd=true) and (tab=true)) then contain:='(+t)(+x)'
      else if (tab=true) then contain:='(+t)'
      else if (crd=true) then contain:='(+x)';

      ADOQuery1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;'+
        'Data Source='+ExtractFilePath(Application.ExeName)+'Data\data.mdb;'+
        'Persist Security Info=False;'+
        'Jet OLEDB:Database Password=""';
      ADOQuery1.SQL.Text:='INSERT into Table1' +
        ' (Id,FileName,Title,Artist,Music,Poem,Contain) VALUES' +
        ' ('+inttostr(i)+',"'+FileNm+'","'+Title+
        '","'+Artist+'","'+Music+'","'+Poem+'","'+Contain+'") ';
//        ' (:Id,:FileName,:Title,:Artist,:Music,:Poem,:Contain) ';
//      ADOQuery1.FieldValues['Id']:=inttostr(i);
//      ADOQuery1.FieldValues['FileName']:=filenm;
//      ADOQuery1.FieldValues['Title']:=title;
//      ADOQuery1.FieldValues['Artist']:=artist;
//      ADOQuery1.FieldValues['Music']:=music;
//      ADOQuery1.FieldValues['Poem']:=poem;
//      ADOQuery1.FieldValues['Contain']:=contain;
      ADOQuery1.ExecSQL;
//      StringGrid1.Cells[0,i]:=filenm;
//      StringGrid1.Cells[1,i]:=title;
//      StringGrid1.Cells[2,i]:=artist;
//      StringGrid1.Cells[3,i]:=music;
//      StringGrid1.Cells[4,i]:=poem;
//      StringGrid1.Cells[5,i]:=contain;

    end;
end;

end.
