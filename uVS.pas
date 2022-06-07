unit uVS;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  StdCtrls, Windows, ComObj;

type

  { TfrmVS }

  TfrmVS = class(TForm)
    Button1: TButton;
    Button2: TButton;
    CheckBox1: TCheckBox;
    CheckBox10: TCheckBox;
    CheckBox11: TCheckBox;
    CheckBox12: TCheckBox;
    CheckBox13: TCheckBox;
    CheckBox14: TCheckBox;
    CheckBox15: TCheckBox;
    CheckBox16: TCheckBox;
    CheckBox17: TCheckBox;
    CheckBox18: TCheckBox;
    CheckBox19: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox20: TCheckBox;
    CheckBox21: TCheckBox;
    CheckBox22: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    CheckBox6: TCheckBox;
    CheckBox7: TCheckBox;
    CheckBox8: TCheckBox;
    CheckBox9: TCheckBox;
    cbLanguage: TComboBox;
    ImageList1: TImageList;
    Label1: TLabel;
    edVSexe: TLabeledEdit;
    edSaveDir: TLabeledEdit;
    od: TOpenDialog;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    sd1: TSaveDialog;
    sbComponents: TScrollBox;
    sd: TSelectDirectoryDialog;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CheckBox17Change(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private

  public

  end;

var
  frmVS: TfrmVS;

implementation

{$R *.lfm}

{ TfrmVS }


function ShellZip(ZipFile, SourceFolder:string): boolean;
const
  emptyzip: array[0..23] of byte  = (80,75,5,6,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
  SHCONTCH_NOPROGRESSBOX = 4;
  SHCONTCH_AUTORENAME = 8;
  SHCONTCH_RESPONDYESTOALL = 16;
  SHCONTF_INCLUDEHIDDEN = 128;
  SHCONTF_FOLDERS = 32;
  SHCONTF_NONFOLDERS = 64;
var
  ms: TMemoryStream;
  shellobj: variant;
  SrcFldr, DestFldr, ShellFldrItems: variant;
  ZipFileV, SourceFolderV: Variant;
begin
  shellobj := CreateOleObject('Shell.Application');

  ZipFileV := ZipFile;
  DestFldr := shellobj.NameSpace(ZipFileV);

  SourceFolderV := SourceFolder;
  SrcFldr := shellobj.NameSpace(SourceFolderV);

  if FileExists(ZipFile) then DeleteFile(PChar(ZipFile));

  ShellFldrItems := SrcFldr.Items;

  ms := TMemoryStream.Create;
  ms.WriteBuffer(emptyzip, sizeof(emptyzip));
  ms.SaveToFile(ZipFile);
  ms.Free;

  try
     DestFldr.CopyHere(ShellFldrItems, 0);
  finally
     ZipFileV := Unassigned;
     SourceFolderV := Unassigned;
     SrcFldr := Unassigned;
     DestFldr := Unassigned;
     shellobj := Unassigned;
  end;
end;

procedure TfrmVS.SpeedButton2Click(Sender: TObject);
begin
  if sd.Execute then
     edSaveDir.Text:=sd.FileName;
end;

procedure TfrmVS.SpeedButton1Click(Sender: TObject);
begin
  if od.Execute then
     begin
       edVSexe.Text:=od.FileName;
     end;
end;

procedure TfrmVS.CheckBox17Change(Sender: TObject);
var
  i: integer;
  Component: TComponent;
begin
  for i := 1 to 16 do
    begin
      Component := frmVS.FindComponent('CheckBox'  + IntToStr(i));
      if Component <> nil then
         begin
           TCheckBox(Component).Enabled := not CheckBox17.Checked;
         end;
    end;
end;

procedure TfrmVS.Button1Click(Sender: TObject);
var
  i: integer;
  cmdLine: string;
  Component: TComponent;
  zipFile: string;
begin
  cmdLine := '';
  for i := 1 to 16 do
    begin
      Component := frmVS.FindComponent('CheckBox'  + IntToStr(i));
      if Component <> nil then
         begin
           if TCheckBox(Component).Checked then
              cmdLine := cmdLine + ' --add ' + TCheckBox(Component).Caption;
         end;
    end;
  if RadioButton1.Checked then cmdLine := cmdLine + ' ' + RadioButton1.Caption;
  if RadioButton2.Checked then cmdLine := cmdLine + ' ' + RadioButton2.Caption;
  if CheckBox19.Checked then cmdLine := cmdLine + ' ' + CheckBox19.Caption;
  if CheckBox20.Checked then cmdLine := cmdLine + ' ' + CheckBox20.Caption;
  if CheckBox21.Checked then cmdLine := cmdLine + ' ' + CheckBox21.Caption;

  if CheckBox18.Checked then
     begin
       cmdLine := cmdLine + ' ' + CheckBox18.Caption + ' ' + cbLanguage.Items[cbLanguage.ItemIndex];
     end;

  cmdLine := cmdLine + ' --layout ' + edSaveDir.Text;

  ExecuteProcess(edVSexe.Text, cmdLine);

  if (CheckBox22.Checked) then
     begin
       zipFile := ExtractFilePath(edSaveDir.Text) + ExtractFileName(edSaveDir.Text) + '.zip';
       ShellZip(zipFile, edSaveDir.Text);
     end;
end;

procedure TfrmVS.Button2Click(Sender: TObject);
var
  i: integer;
  cmdLine: string;
  Component: TComponent;
  zipFile: string;
begin
  cmdLine := '';
  for i := 1 to 16 do
    begin
      Component := frmVS.FindComponent('CheckBox'  + IntToStr(i));
      if Component <> nil then
         begin
           if TCheckBox(Component).Checked then
              cmdLine := cmdLine + ' --add ' + TCheckBox(Component).Caption;
         end;
    end;
  if RadioButton1.Checked then cmdLine := cmdLine + ' ' + RadioButton1.Caption;
  if RadioButton2.Checked then cmdLine := cmdLine + ' ' + RadioButton2.Caption;
  if CheckBox19.Checked then cmdLine := cmdLine + ' ' + CheckBox19.Caption;
  if CheckBox20.Checked then cmdLine := cmdLine + ' ' + CheckBox20.Caption;
  if CheckBox21.Checked then cmdLine := cmdLine + ' ' + CheckBox21.Caption;

  if CheckBox18.Checked then
     begin
       cmdLine := cmdLine + ' ' + CheckBox18.Caption + ' ' + cbLanguage.Items[cbLanguage.ItemIndex];
     end;

  cmdLine := cmdLine + ' --layout ' + edSaveDir.Text;

  ShowMessage(cmdLine);
end;

end.

