unit u_assifparser_testform;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ComCtrls,
  u_assifparser_class;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Edit1: TEdit;
    ListBox1: TListBox;
    TreeView1: TTreeView;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
var
  Parser:TAssiFormulaParser;
begin
  Parser:=TAssiFormulaParser.Create;
  TreeView1.Items.Clear;
  if Parser.Execute(Edit1.Text) then
  begin
     Parser.UploadToTree(TreeView1);
     ShowMessage(Parser.GetResultAsText);
  end
  else begin
     ShowMessage('Error');
  end;
  Parser.Free;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  Parser:TAssiFormulaParser;
begin
  Parser:=TAssiFormulaParser.Create;
  TreeView1.Items.Clear;
  if Parser.Parse(Edit1.Text) then
  begin
     Parser.UploadToTree(TreeView1);
  end
  else begin
       ShowMessage('Error');
  end;
  Parser.Free;
end;

procedure TForm1.ListBox1Click(Sender: TObject);
begin
  if ListBox1.ItemIndex<>-1 then
  Edit1.Text:=ListBox1.Items.Strings[ListBox1.ItemIndex];
end;

end.

