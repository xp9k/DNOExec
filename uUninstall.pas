unit uUninstall;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  ComCtrls, Menus, VirtualTrees, Windows, Registry, StrUtils, DnoTypes;

type

  { TfrmUninstall }

  TfrmUninstall = class(TForm)
    Button1: TButton;
    Button2: TButton;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    Panel1: TPanel;
    Panel2: TPanel;
    pm: TPopupMenu;
    ProgressBar1: TProgressBar;
    StatusBar1: TStatusBar;
    VST: TVirtualStringTree;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure pmPopup(Sender: TObject);
    procedure VSTBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
    procedure VSTChecking(Sender: TBaseVirtualTree; Node: PVirtualNode;
      var NewState: TCheckState; var Allowed: Boolean);
    procedure VSTCompareNodes(Sender: TBaseVirtualTree; Node1,
      Node2: PVirtualNode; Column: TColumnIndex; var Result: Integer);
    procedure VSTGetNodeDataSize(Sender: TBaseVirtualTree;
      var NodeDataSize: Integer);
    procedure VSTGetPopupMenu(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; const P: TPoint; var AskParent: Boolean;
      var pMenu: TPopupMenu);
    procedure VSTGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: String);
    procedure VSTInitNode(Sender: TBaseVirtualTree; ParentNode,
      Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
    procedure VSTResize(Sender: TObject);
  private
    UninstallProgress: integer;
    procedure ScanForApplications;
    procedure ParseUninstallString(var Uninstall: PUninstall);
  public

  end;

var
  frmUninstall: TfrmUninstall;

implementation
  uses
    main;

{$R *.lfm}

{ TfrmUninstall }



procedure TfrmUninstall.Button1Click(Sender: TObject);
procedure ExecuteNodes(Node: PVirtualNode);
  var
    _Node: PVirtualNode;
    Uninstall: PUninstall;
  begin
    _Node := Node;
    while Assigned(_Node) do
      begin
        if  (_Node^.CheckState = csCheckedNormal) then
          begin
            Uninstall := VST.GetNodeData(_Node);
            inc(UninstallProgress);
            ProgressBar1.Position := UninstallProgress;
            StatusBar1.SimpleText := 'Удаляется ' + Uninstall^.DisplayName + ' ' + Uninstall^.Version;
            Application.ProcessMessages;
            ExecuteProcess(RawByteString(Uninstall^.FilePath), RawByteString(Uninstall^.Args));
//            sleep(2000);
          end;
        _Node := VST.GetNextSibling(_Node);
      end;
  end;
var
  Node: PVirtualNode;
begin
  UninstallProgress := 0;
  ProgressBar1.Position:=0;
  ProgressBar1.Max := VST.CheckedCount;
  Node := VST.GetFirst;
  ExecuteNodes(Node);
  ProgressBar1.Position:=ProgressBar1.Max;
  StatusBar1.SimpleText := '';
  ScanForApplications;
end;

procedure TfrmUninstall.Button2Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmUninstall.FormShow(Sender: TObject);
begin
  ScanForApplications;
end;

procedure TfrmUninstall.MenuItem1Click(Sender: TObject);
procedure ExecuteNodes(Node: PVirtualNode);
  var
    _Node: PVirtualNode;
    Uninstall: PUninstall;
  begin
    _Node := Node;
    if Assigned(_Node) then
      begin
        Uninstall := VST.GetNodeData(_Node);
        inc(UninstallProgress);
        ProgressBar1.Position := UninstallProgress;
        StatusBar1.SimpleText := 'Удаляется ' + Uninstall^.DisplayName + ' ' + Uninstall^.Version;
        Application.ProcessMessages;
        ExecuteProcess(RawByteString(Uninstall^.UninstallString), '');
      end;
  end;
var
  Node: PVirtualNode;
begin
  UninstallProgress := 0;
  ProgressBar1.Position:=0;
  ProgressBar1.Max := VST.CheckedCount;
  Node := VST.FocusedNode;
  if Assigned(Node) then
    begin
      ExecuteNodes(Node);
      ProgressBar1.Position:=ProgressBar1.Max;
      StatusBar1.SimpleText := '';
      ScanForApplications;
    end;
end;

procedure TfrmUninstall.MenuItem2Click(Sender: TObject);
var
  NodeData: PUninstall;
  FilePath: string;
  p: integer;
begin
  NodeData := VST.GetNodeData(VST.FocusedNode);
  p := pos('/', NodeData^.UninstallString);
  if p > 0 then
    FilePath := Copy(NodeData^.UninstallString, 1, p - 1).Replace('"', '').Trim
  else
    FilePath :=NodeData^.UninstallString.Replace('"', '').Trim;
  FilePath := ExtractFileDir(FilePath);
  ShellExecute(0, 'open', 'explorer.exe', PChar(FilePath), '', SW_SHOW);
end;

procedure TfrmUninstall.MenuItem3Click(Sender: TObject);
procedure UncheckNode(Node: PVirtualNode);
  var
    _Node: PVirtualNode;
  begin
    _Node := Node;
    while Assigned(_Node) do
      begin
        _Node^.CheckState := csUncheckedNormal;
        _Node := VST.GetNextSibling(_Node);
      end;
  end;
var
  Node: PVirtualNode;
begin
  Node := VST.GetFirst;
  VST.BeginUpdate;
  UncheckNode(Node);
  VST.EndUpdate;
end;

procedure TfrmUninstall.MenuItem4Click(Sender: TObject);
begin
  ScanForApplications;
end;

procedure TfrmUninstall.pmPopup(Sender: TObject);
begin
  self.MenuItem1.Visible := Assigned(VST.FocusedNode);
  self.MenuItem2.Visible := Assigned(VST.FocusedNode);
end;

procedure TfrmUninstall.VSTBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
begin
  if Node^.Index mod 2 = 0 then
    begin
      TargetCanvas.Brush.Color := $F0F0F0;
      TargetCanvas.FillRect(CellRect);
    end;
end;

procedure TfrmUninstall.VSTChecking(Sender: TBaseVirtualTree;
  Node: PVirtualNode; var NewState: TCheckState; var Allowed: Boolean);
var
  NodeData: PUninstall;
begin
  NodeData := VST.GetNodeData(Node);
  if NodeData^.UninstallString = '' then Allowed := false;
end;

procedure TfrmUninstall.VSTCompareNodes(Sender: TBaseVirtualTree; Node1,
  Node2: PVirtualNode; Column: TColumnIndex; var Result: Integer);
var
  Data1,
  Data2: PUninstall;
begin
  Data1 := PUninstall(Sender.GetNodeData(Node1));
  Data2 := PUninstall(Sender.GetNodeData(Node2));
  case Column of
    0: Result := Result + CompareText(Data1^.DisplayName, Data2^.DisplayName);
    1: Result := Result + CompareText(Data1^.Version, Data2^.Version);
    2: Result := Result + CompareText(Data1^.UninstallString, Data2^.UninstallString);
  end;
end;

procedure TfrmUninstall.VSTGetNodeDataSize(Sender: TBaseVirtualTree;
  var NodeDataSize: Integer);
begin
  NodeDataSize:=sizeof(TUninstall);
end;

procedure TfrmUninstall.VSTGetPopupMenu(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; const P: TPoint;
  var AskParent: Boolean; var pMenu: TPopupMenu);
begin
  if Node <> nil then
   begin
    VST.Selected[Node] := true;
    VST.FocusedNode := Node;
    pMenu := pm;
   end;
   //else
   //  pMenu := nil;
end;

procedure TfrmUninstall.VSTGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
  Column: TColumnIndex; TextType: TVSTTextType; var CellText: String);
var
  NodeData: PUninstall;
begin
  NodeData := PUninstall(VST.GetNodeData(Node));
  case Column of
    0:
       begin
         CellText := NodeData^.DisplayName;
       end;
    1:
       begin
         CellText := NodeData^.Version;
       end;
    2:
       begin
         CellText := NodeData^.UninstallString;
       end;
  end;
end;

procedure TfrmUninstall.VSTInitNode(Sender: TBaseVirtualTree; ParentNode,
  Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
  procedure EnumTree(Node, UninstallNode: PVirtualNode; CheckFor: string);
    var
      _Node: PVirtualNode;
      App: PApp;
    begin
      _Node := Node;
      while Assigned(_Node) do
        begin
          App := frmMain.VST.GetNodeData(_Node);
          if (_Node^.CheckState = csCheckedNormal) or (_Node^.CheckState = csMixedNormal) then
             if ContainsText(CheckFor, App^.Name) then UninstallNode^.CheckState := csCheckedNormal;
          Application.ProcessMessages;
          if (_Node^.ChildCount > 0) then EnumTree(_Node^.FirstChild, UninstallNode, CheckFor);
          _Node := VST.GetNextSibling(_Node);
        end;
    end;
var
  Uninstall: PUninstall;
  App: PApp;
  AppNode: PVirtualNode;
begin
  Node^.CheckType := ctCheckBox;
  Uninstall := VST.GetNodeData(Node);
  AppNode := frmMain.VST.GetFirst;
  App := frmMain.VST.GetNodeData(AppNode);
  EnumTree(AppNode, Node, Uninstall^.DisplayName);
end;

procedure TfrmUninstall.VSTResize(Sender: TObject);
begin
  VST.Header.Columns[2].Width := VST.Width - (VST.Header.Columns[0].Width + VST.Header.Columns[1].Width + GetSystemMetrics(SM_CXVSCROLL) + 5);
end;

procedure TfrmUninstall.ScanForApplications;

procedure ScanKey(RootKey: HKEY; Key: string);
var
  reg: TRegistry;
  i: Integer;
  keys: TStringList;
  Node: PVirtualNode;
  NodeData: PUninstall;
begin
  keys :=  TStringList.Create;
  reg := TRegistry.Create;
  try
     reg.RootKey := RootKey;
     reg.OpenKey(Key, false);
     reg.GetKeyNames(keys);
     for i := 0 to keys.Count - 1 do
      begin
         reg.CloseKey;
         if reg.OpenKey(Key + '\' + keys[i], false) then
           if reg.ReadString('DisplayName') <> '' then
            begin
             Node := VST.AddChild(nil);
             NodeData := VST.GetNodeData(Node);
             NodeData^.DisplayName := reg.ReadString('DisplayName');
             NodeData^.Version := reg.ReadString('DisplayVersion');
             if reg.ValueExists('QuietUninstallString') then
               NodeData^.UninstallString := reg.ReadString('QuietUninstallString')
             else
               NodeData^.UninstallString := reg.ReadString('UninstallString');
            end;
         ParseUninstallString(NodeData);
      end;
  finally
    reg.Free;
    keys.free;
  end;
end;

begin
  VST.Clear;
  VST.BeginUpdate;

  ScanKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall');
  ScanKey(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall');
  ScanKey(HKEY_LOCAL_MACHINE, 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall');

  VST.EndUpdate;
end;


procedure TfrmUninstall.ParseUninstallString(var Uninstall: PUninstall);
var
  args: TStringArray;
  Delimiter: char;
begin
  if Uninstall^.UninstallString.Contains('MsiExec') then
   begin
     Uninstall^.UninstallString := Uninstall^.UninstallString.Replace('/I', '/X');
     args := Uninstall^.UninstallString.Split('/');
     if Length(args) > 1 then
      begin
        Uninstall^.FilePath := args[0];
        args := Copy(args, 1, Length(args) - 1);
        Uninstall^.Args := '/' + AnsiString.Join('/', args);
        if not ContainsText(Uninstall^.Args, '/qn') then Uninstall^.Args += ' /qn';
        if not ContainsText(Uninstall^.Args, '/norestart') then Uninstall^.Args += ' /norestart';
      end;
   end
  else
   begin
    if ContainsText(Uninstall^.UninstallString, '--') then Delimiter := '-' else Delimiter := '/';
    args := Uninstall^.UninstallString.Split(Delimiter);
    if Length(args) > 1 then
     begin
      Uninstall^.FilePath := args[0];
      args := Copy(args, 1, Length(args) - 1);
      Uninstall^.Args := Delimiter + AnsiString.Join(Delimiter, args);
     end;
    if not ContainsText(Uninstall^.Args, Delimiter + 'quiet') then Uninstall^.Args += ' ' + Delimiter + 'quiet';
    if not ContainsText(Uninstall^.Args, Delimiter + 'S') then Uninstall^.Args += ' ' + Delimiter + 'S';
    if not ContainsText(Uninstall^.Args, Delimiter + 'SILENT') then Uninstall^.Args += ' ' + Delimiter + 'SILENT';
    if not ContainsText(Uninstall^.Args, Delimiter + 'AUTO') then Uninstall^.Args += ' ' + Delimiter + 'AUTO';
    if not ContainsText(Uninstall^.Args, Delimiter + 'NORESTART') then Uninstall^.Args += ' ' + Delimiter + 'NORESTART';
   end;
end;

end.

