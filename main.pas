unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  ComCtrls, Buttons, Menus, VirtualTrees, StrUtils, Dom, XMLRead,
  XMLWrite, Process, ShlObj, ActiveX, Windows, ComObj, uVS;

type

  PExecution = ^TExecution;
  TExecution = record
    Filename: string;
    Args: string;
    Hidden: Boolean;
  end;

  PLink = ^TLink;
  TLink = record
    Filename: string;
    TargetName: string;
    WorkingDirectory: string;
  end;

  PApp = ^TApp;
  TApp = record
    Name: string;
    Version: string;
    Exclusive: boolean;
    Checked: boolean;
    Parent: PApp;
    Executions: TList;
    Links: TList;
    State: Integer;
  end;

  { TfrmMain }

  TfrmMain = class(TForm)
    BevelExecutions: TBevel;
    BevelLinks: TBevel;
    btInstall: TButton;
    btSaveApplication: TBitBtn;
    btSaveExecution: TBitBtn;
    btSaveLink: TBitBtn;
    cbChecked: TCheckBox;
    cbExclusive: TCheckBox;
    cbExecutions: TComboBox;
    cbLinks: TComboBox;
    edArguments: TLabeledEdit;
    edFilename: TLabeledEdit;
    edLinkFilename: TLabeledEdit;
    edLinkTargetname: TLabeledEdit;
    edLinkWorkDir: TLabeledEdit;
    edName: TLabeledEdit;
    edVersion: TLabeledEdit;
    ImageList1: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    N1: TMenuItem;
    OpenDialog1: TOpenDialog;
    odExecution: TOpenDialog;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TScrollBox;
    Panel8: TPanel;
    Panel9: TPanel;
    PanelApplication: TPanel;
    PanelExecutions: TPanel;
    PanelLinks: TPanel;
    pbLabel: TLabel;
    Panel1: TPanel;
    ProgressBar1: TProgressBar;
    SaveDialog1: TSaveDialog;
    SpeedButton1: TSpeedButton;
    SpeedButton10: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton12: TSpeedButton;
    SpeedButton13: TSpeedButton;
    SpeedButton14: TSpeedButton;
    SpeedButton15: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton9: TSpeedButton;
    Splitter1: TSplitter;
    VST: TVirtualStringTree;
    procedure btSaveApplicationClick(Sender: TObject);
    procedure btSaveExecutionClick(Sender: TObject);
    procedure btInstallClick(Sender: TObject);
    procedure btSaveLinkClick(Sender: TObject);
    procedure cbExecutionsChange(Sender: TObject);
    procedure cbLinksChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem5Click(Sender: TObject);
    procedure MenuItem6Click(Sender: TObject);
    procedure SpeedButton10Click(Sender: TObject);
    procedure SpeedButton11Click(Sender: TObject);
    procedure SpeedButton12Click(Sender: TObject);
    procedure SpeedButton13Click(Sender: TObject);
    procedure SpeedButton14Click(Sender: TObject);
    procedure SpeedButton15Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure VSTBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
    procedure VSTFocusChanged(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex);
    procedure VSTGetNodeDataSize(Sender: TBaseVirtualTree;
      var NodeDataSize: Integer);
    procedure VSTGetText(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
      var CellText: String);
    procedure VSTInitNode(Sender: TBaseVirtualTree; ParentNode,
      Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
    procedure VSTMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure VSTResize(Sender: TObject);
  private

  public
    SelectedNode: PVirtualNode;
    ProgessBarPosition: integer;
    Installing: boolean;
    function GetAppVersionStr(Filename: string): string;
    procedure CreateLink(Filename, TargetName, WorkingDirectory: string);
    function GetDesktopDir: string;
    procedure LoadXMLConfig(Filename: string);
    procedure SaveXMLConfig(Filename: string);
    procedure LoadNodes(XMLNode: TDOMNode; VSTParentNode: PVirtualNode);
    procedure SaveNodes(Doc: TXMLDocument; Node: PVirtualNode; XMLNode: TDOMNode);
    function InsertNodes(ParentNode: PVirtualNode; App: PApp): PVirtualNode;
    procedure RunExecution(Execution: PExecution);
    procedure RunExecutionByProcess(Execution: PExecution);
    procedure RunApplication(Application: PApp);
    procedure RunCreateLink(Link: PLink);
    function ReplaceEnvs(InputString: string): string;
    function ApplicationHasExecutions(Application: PApp): boolean;
    Procedure ClearComponents;
    procedure ShowArgumentsHelp;
    procedure CheckcbExecutions;
    procedure CheckcbLinks;
  end;

const
  RIGHT_MARGIN = 24;
  CONFIG_PANEL_WIDTH = 450;

  STATE_NONE = 0;
  STATE_NOT_FOUND = -1;
  STATE_INSTALLED = 1;

var
  frmMain: TfrmMain;

implementation

{$R *.lfm}

{ TfrmMain }

function MyBoolToStr(Val: Boolean): string;
begin
  if val then result := 'True' else Result := 'False';
end;

procedure TfrmMain.btInstallClick(Sender: TObject);
procedure ExecuteNodes(Node: PVirtualNode);
  var
    _Node: PVirtualNode;
    App: PApp;
  begin
    _Node := Node;
    while Assigned(_Node) and Installing do
      begin
        if  ((_Node^.CheckState = csCheckedNormal) or (_Node^.CheckState = csMixedNormal)) then
          begin
            VST.Selected[_Node] := True;
            App := VST.GetNodeData(_Node);
            pbLabel.Caption:= 'Установка ' + App^.Name;
            inc(ProgessBarPosition);
            ProgressBar1.Position := ProgessBarPosition;
            pbLabel.BringToFront;
            pbLabel.SendToBack;
            Application.ProcessMessages;
            RunApplication(App);
            if App^.State <> STATE_NOT_FOUND then App^.State := STATE_INSTALLED;
            Application.ProcessMessages;
            if (_Node^.ChildCount > 0) then ExecuteNodes(_Node^.FirstChild);
          end;
        ;
        _Node := VST.GetNextSibling(_Node);
      end;
  end;
var
  Node: PVirtualNode;
  App: PApp;
  i: integer;
begin
  if Installing then begin
    ProgressBar1.Position := 0;
    ProgessBarPosition := 0;
    btInstall.Caption := 'Установить';
    Installing := False;
  end
  else
  begin
    Installing := True;
    btInstall.Caption := 'Отмена';
    ProgressBar1.Max := VST.CheckedCount;
    ProgressBar1.Position := 0;
    ProgessBarPosition := 0;

    Node := VST.GetFirst;
    App := VST.GetNodeData(Node);
    ExecuteNodes(Node);

    ProgressBar1.Position := ProgressBar1.Max;
    pbLabel.Caption:= 'Установка завершена';
    Installing := False;
    btInstall.Caption := 'Установить';
  end;
//  if Installing then Installing := not Installing;
end;

procedure TfrmMain.btSaveLinkClick(Sender: TObject);
var
  NodeData: PApp;
  Link: PLink;
begin
  if not Assigned(SelectedNode) then exit;

  if cbLinks.Items.Count > 0 then
    begin
      VST.BeginUpdate;
      NodeData := VST.GetNodeData(SelectedNode);
      Link := NodeData^.Links[cbLinks.ItemIndex];
      Link^.Filename := edLinkFilename.Text;
      Link^.TargetName := edLinkTargetname.Text;
      Link^.WorkingDirectory := edLinkWorkDir.Text;
      VST.EndUpdate;
    end;
end;

procedure TfrmMain.cbExecutionsChange(Sender: TObject);
var
  NodeData: PApp;
  Execution: PExecution;
begin
  if cbExecutions.Items.Count > 0 then
    begin
      NodeData := VST.GetNodeData(SelectedNode);
      Execution := NodeData^.Executions[cbExecutions.ItemIndex];
      edFilename.Text := Execution^.Filename;
      edArguments.Text := Execution^.Args;
    end;
  CheckcbExecutions;
end;

procedure TfrmMain.cbLinksChange(Sender: TObject);
var
  NodeData: PApp;
  Link: PLink;
begin
  if cbLinks.Items.Count > 0 then
    begin
      NodeData := VST.GetNodeData(SelectedNode);
      Link := NodeData^.Links[cbLinks.ItemIndex];
      edLinkFilename.Text := Link^.Filename;
      edLinkTargetname.Text := Link^.TargetName;
      edLinkWorkDir.Text := Link^.WorkingDirectory;
    end;
end;

procedure TfrmMain.FormActivate(Sender: TObject);
begin
  if Application.HasOption('a', 'autorun') then
    begin
      btInstallClick(self);
    end
end;

procedure TfrmMain.btSaveApplicationClick(Sender: TObject);
var
  NodeData: PApp;
begin
  if not Assigned(SelectedNode) then exit;
  VST.BeginUpdate;
  NodeData := VST.GetNodeData(SelectedNode);
  NodeData^.Name := edName.Text;
  NodeData^.Version := edVersion.Text;
  NodeData^.Checked := cbChecked.Checked;
  NodeData^.Exclusive := cbExclusive.Checked;
  if (NodeData^.Exclusive) then
      SelectedNode^.CheckType := ctRadioButton
  else
      SelectedNode^.CheckType := ctTriStateCheckBox;
  if NodeData^.Checked then SelectedNode^.CheckState := csCheckedNormal else SelectedNode^.CheckState := csUncheckedNormal;
  VST.EndUpdate;
end;

procedure TfrmMain.btSaveExecutionClick(Sender: TObject);
var
  NodeData: PApp;
  Execution: PExecution;
begin
  if not Assigned(SelectedNode) then exit;

  if cbExecutions.Items.Count > 0 then
    begin
      VST.BeginUpdate;
      NodeData := VST.GetNodeData(SelectedNode);
      Execution := NodeData^.Executions[cbExecutions.ItemIndex];
      Execution^.Filename := edFilename.Text;
      Execution^.Args := edArguments.Text;
      VST.EndUpdate;
    end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ConfigFilename: string;
begin
  if Application.HasOption('c', 'config') then
   begin
     ConfigFilename := Application.GetOptionValue('c', 'config');
   end
  else
     ConfigFilename := ExtractFileDir(ParamStr(0)) + '\' + 'Config.xml';
  if FileExists(ConfigFilename) then
   LoadXMLConfig(ConfigFilename)
  else
   VSTMouseDown(self, mbRight, [], 0, 0);
  VST.FullExpand;

  pbLabel.Parent := ProgressBar1;
  Panel2.Constraints.MinWidth := Panel3.Constraints.MinWidth + Panel4.Constraints.MinWidth;
  CheckcbExecutions;
  CheckcbLinks;

  Installing := False;
end;

procedure TfrmMain.FormResize(Sender: TObject);
begin
  VST.Header.Columns[0].Width := VST.ClientWidth - VST.Header.Columns[1].Width - 20;
end;

procedure TfrmMain.MenuItem2Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    begin
      LoadXMLConfig(OpenDialog1.FileName);
      VST.FullExpand;
      pbLabel.Parent := ProgressBar1;
    end;
end;

procedure TfrmMain.MenuItem3Click(Sender: TObject);
begin
  if SaveDialog1.Execute then
    begin
      SaveXMLConfig(SaveDialog1.FileName);
    end;
end;

procedure TfrmMain.MenuItem5Click(Sender: TObject);
begin
  frmVS.ShowModal;
end;

procedure TfrmMain.MenuItem6Click(Sender: TObject);
var
  ConfigFilename: string;
begin
 ConfigFilename := ExtractFileDir(ParamStr(0)) + '\' + 'Config.xml';
  if FileExists(ConfigFilename) then
    begin
    if Application.MessageBox('Файл уже существует. Заменить?', 'Внимание', MB_ICONQUESTION + MB_YESNO) = IDYES then
      SaveXMLConfig(ConfigFilename);
    end
  else
    SaveXMLConfig(ConfigFilename);
end;


procedure TfrmMain.SpeedButton10Click(Sender: TObject);
var
  NodeData: PApp;
  index: integer;
begin
  if cbLinks.Items.Count > 0 then
    begin
      index := cbLinks.ItemIndex;
      NodeData := VST.GetNodeData(SelectedNode);
      NodeData^.Links.Remove(NodeData^.Links[cbLinks.ItemIndex]);
      cbLinks.Items.Delete(cbLinks.ItemIndex);
    end;
  if cbLinks.Items.Count > 0 then
    if index = 0 then
        cbLinks.ItemIndex := index
      else
        cbLinks.ItemIndex := index - 1;
  cbLinksChange(self);
  CheckcbLinks;
end;

procedure TfrmMain.SpeedButton11Click(Sender: TObject);
begin
  PanelExecutions.Visible := not PanelExecutions.Visible;
  if PanelExecutions.Visible then SpeedButton11.ImageIndex := 8 else SpeedButton11.ImageIndex := 7;
end;

procedure TfrmMain.SpeedButton12Click(Sender: TObject);
begin
  PanelLinks.Visible := not PanelLinks.Visible;
  if PanelLinks.Visible then SpeedButton12.ImageIndex := 8 else SpeedButton12.ImageIndex := 7;
end;

procedure TfrmMain.SpeedButton13Click(Sender: TObject);
var
  Filename: string;
begin
  odExecution.InitialDir := ExtractFileDir(ParamStr(0));
  if odExecution.Execute then
    begin
      Filename := odExecution.FileName;
      Filename := ReplaceStr(Filename, ExtractFileDir(ParamStr(0)), '{src}');
      edFilename.Text := Filename;
    end;
end;

procedure TfrmMain.SpeedButton14Click(Sender: TObject);
begin
  ShowMessage('Ярлык будет помещен на рабочий стол');
end;

procedure TfrmMain.SpeedButton15Click(Sender: TObject);
begin
  ShowArgumentsHelp;
end;

procedure TfrmMain.SpeedButton1Click(Sender: TObject);
begin
  VST.MoveTo(VST.FocusedNode, VST.GetPreviousSibling(VST.FocusedNode), amInsertBefore, false);
end;

procedure TfrmMain.SpeedButton2Click(Sender: TObject);
begin
  VST.MoveTo(VST.FocusedNode, VST.GetNextSibling(VST.FocusedNode), amInsertAfter, false);
end;

procedure TfrmMain.SpeedButton3Click(Sender: TObject);
begin
  if not Assigned(VST.FocusedNode) then exit;

  VST.BeginUpdate;
  VST.MoveTo(VST.FocusedNode, VST.GetNextSibling(VST.FocusedNode), amAddChildLast, false);
  VST.EndUpdate;
end;

procedure TfrmMain.SpeedButton4Click(Sender: TObject);
begin
  if VST.FocusedNode = nil then exit;

  VST.BeginUpdate;
  VST.MoveTo(VST.FocusedNode, VST.FocusedNode^.Parent, amInsertBefore, false);
  VST.EndUpdate;
end;

procedure TfrmMain.SpeedButton5Click(Sender: TObject);
var
  NodeData: PApp;
  Execution: PExecution;
begin
  if not Assigned(SelectedNode) then exit;
  new(Execution);
  NodeData := VST.GetNodeData(SelectedNode);
  NodeData^.Executions.Add(Execution);
  cbExecutions.Items.Add('Команда #' + IntToStr(cbExecutions.Items.Count + 1));
  cbExecutions.ItemIndex := cbExecutions.Items.Count - 1;
  cbExecutionsChange(self);
end;

procedure TfrmMain.SpeedButton6Click(Sender: TObject);
var
  NodeData: PApp;
  index: integer;
begin
  if cbExecutions.Items.Count > 0 then
    begin
      index := cbExecutions.ItemIndex;
      NodeData := VST.GetNodeData(SelectedNode);
      NodeData^.Executions.Remove(NodeData^.Executions[cbExecutions.ItemIndex]);
      cbExecutions.Items.Delete(cbExecutions.ItemIndex);
    end;
  if cbExecutions.Items.Count > 0 then
    if index = 0 then
        cbExecutions.ItemIndex := index
      else
        cbExecutions.ItemIndex := index - 1;
  cbExecutionsChange(self);
  CheckcbExecutions;
end;

procedure TfrmMain.SpeedButton7Click(Sender: TObject);
var
  NodeData: PApp;
  Node: PVirtualNode;
begin
  if VST.FocusedNode <> nil then
    Node := VST.AddChild(VST.FocusedNode^.Parent)
  else
    Node := VST.AddChild(nil);
  NodeData := VST.GetNodeData(Node);
  NodeData^.Executions := TList.Create;
  NodeData^.Links := TList.Create;
  NodeData^.Name := 'Новое Имя';
  SelectedNode := Node;
  ClearComponents;
  CheckcbExecutions;
  CheckcbLinks;
  VST.FocusedNode := Node;
end;

procedure TfrmMain.SpeedButton8Click(Sender: TObject);
begin
  VST.DeleteNode(VST.FocusedNode);
  ClearComponents;
  SelectedNode:=nil;
  CheckcbExecutions;
  CheckcbLinks;
end;

procedure TfrmMain.SpeedButton9Click(Sender: TObject);
var
  NodeData: PApp;
  Link: PLink;
begin
  if not Assigned(SelectedNode) then exit;
  new(Link);
  NodeData := VST.GetNodeData(SelectedNode);
  NodeData^.Links.Add(Link);
  cbLinks.Items.Add('Ярлык #' + IntToStr(cbLinks.Items.Count + 1));
  cbLinks.ItemIndex := cbLinks.Items.Count - 1;
  cbLinksChange(self);
  CheckcbLinks;
end;

procedure TfrmMain.VSTBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);

var
  NodeData: PApp;
begin
  NodeData := VST.GetNodeData(Node);
  case NodeData^.State of
    STATE_NOT_FOUND:
       begin
         TargetCanvas.Brush.Color := clRed;
         TargetCanvas.FillRect(CellRect);
         TargetCanvas.Font.Style:=[fsStrikeOut];
       end;
    STATE_INSTALLED:
       begin
         TargetCanvas.Brush.Color := clLime;
         TargetCanvas.FillRect(CellRect);
         TargetCanvas.Font.Style:=[fsStrikeOut];
       end;
  end;
end;

procedure TfrmMain.VSTFocusChanged(Sender: TBaseVirtualTree; Node: PVirtualNode;
  Column: TColumnIndex);
var
  NodeData: PApp;
  i: integer;
begin
  ClearComponents;
  SelectedNode := Node;
  NodeData := VST.GetNodeData(Node);
  if NodeData = nil then exit;

  edName.Text := NodeData^.Name;
  edVersion.text := NodeData^.Version;

  for i := 0 to NodeData^.Executions.Count - 1 do
      begin
           cbExecutions.Items.Add('Команда #' + IntToStr(i + 1));
      end;
  if cbExecutions.Items.Count > 0 then cbExecutions.ItemIndex := 0;
  cbExclusive.Checked := NodeData^.Exclusive;
  cbChecked.Checked := NodeData^.Checked;
  cbExecutionsChange(self);

  for i := 0 to NodeData^.Links.Count - 1 do
      begin
           cbLinks.Items.Add('Ярлык #' + IntToStr(i + 1));
      end;
  if cbLinks.Items.Count > 0 then cbLinks.ItemIndex := 0;
  cbLinksChange(self);
  CheckcbExecutions;
  CheckcbLinks;
end;

procedure TfrmMain.VSTGetNodeDataSize(Sender: TBaseVirtualTree;
  var NodeDataSize: Integer);
begin
  NodeDataSize := SizeOf(TApp);
end;


procedure TfrmMain.VSTGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: String);
var
  App: PApp;
begin
  App := PApp(VST.GetNodeData(Node));
  case Column of
    0:
       begin
         CellText := App^.Name;
       end;
    1:
       begin
         CellText := App^.Version;
       end;
  end;
end;

procedure TfrmMain.VSTInitNode(Sender: TBaseVirtualTree;
  ParentNode, Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
var
  App: PApp;
  vNode: PVirtualNode;
begin
  vNode := Sender.GetNodeData(Node);
  App := PApp(vNode);
  if (App^.Exclusive) then
      Node^.CheckType := ctRadioButton
  else
      Node^.CheckType := ctTriStateCheckBox;
  if App^.Checked then Node^.CheckState := csCheckedNormal;
end;

procedure TfrmMain.VSTMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbRight then
    begin
     Panel2.Visible := not Panel2.Visible;
     if Panel2.Visible then
       begin
        frmMain.Width := frmMain.Width + Panel2.Width;
        frmMain.Menu := MainMenu1;
       end
     else
       begin
        frmMain.Width := frmMain.Width - Panel2.Width;
        frmMain.Menu := nil;
       end;
    end;
end;

procedure TfrmMain.VSTResize(Sender: TObject);
begin
  FormResize(self);
end;

function TfrmMain.GetAppVersionStr(Filename: string): string;
var
  Rec: Cardinal;
begin
  Rec := GetFileVersion(Filename);
  if Rec <> $fffffff then
    Result := Format('%d.%d', [LongRec(Rec).Hi, LongRec(Rec).Lo])
  else
    Result := '';
end;

procedure TfrmMain.CreateLink(Filename, TargetName, WorkingDirectory: string);
var
  IObject : IUnknown;
  ISLink : IShellLink;
  IPFile : IPersistFile;
  LinkName : WideString;
begin
  IObject := CreateComObject(CLSID_ShellLink) ;
  ISLink := IObject as IShellLink;
  IPFile := IObject as IPersistFile;
  with ISLink do
    begin
      SetPath(pChar(TargetName)) ;
      SetWorkingDirectory(pChar(WorkingDirectory)) ;
      LinkName := GetDesktopDir + '\' + Filename + '.lnk';
      IPFile.Save(PWChar(LinkName), false) ;
    end
end;

function TfrmMain.GetDesktopDir: string;
var
 dt: array [0..266] of char;
begin
  SHGetSpecialFolderPath(0, dt, CSIDL_DESKTOPDIRECTORY, false);
  result := dt;
end;

procedure TfrmMain.LoadXMLConfig(Filename: string);
var
  Doc: TXMLDocument;
begin
  VST.Clear;
  ReadXMLFile(Doc, Filename);
  try
    LoadNodes(Doc.DocumentElement, nil);
  finally
    Doc.Free;
  end;
end;

procedure TfrmMain.SaveXMLConfig(Filename: string);
var
  Doc: TXMLDocument;
  Node: PVirtualNode;
  RootNode: TDOMNode;
  Attribute: TDOMAttr;
begin
  try
    Doc := TXMLDocument.create;
    RootNode := Doc.CreateElement('Config');
    Doc.Appendchild(RootNode);

    Attribute := Doc.CreateAttribute('Width');
    Attribute.Value := IntToStr(VST.Width);
    RootNode.Attributes.SetNamedItem(Attribute);

    Attribute := Doc.CreateAttribute('Height');
    Attribute.Value := IntToStr(frmMain.Height - MainMenu1.Height);
    RootNode.Attributes.SetNamedItem(Attribute);

    Node := VST.GetFirst;
    SaveNodes(Doc, Node, RootNode);

    WriteXMLFile(Doc, Filename);
  finally
    RootNode.Free;
    Doc.Free;
  end;
end;

procedure TfrmMain.LoadNodes(XMLNode: TDOMNode; VSTParentNode: PVirtualNode);
  function GetNodeAttribute(Node: TDOMNode; Attribute: string): Variant;
  var
    NodeAttribute: TDOMNode;
  begin
    NodeAttribute := Node.Attributes.GetNamedItem(Attribute);
    if (Assigned(NodeAttribute)) and (NodeAttribute.NodeValue <> '') then
      Result := NodeAttribute.NodeValue
    else
      Result := nil;
  end;

  function GetExecutions(Node: TDOMNode): TList;
  var
    i: integer;
    Execution: PExecution;
  begin
    Result := TList.Create;
    for i := 0 to Node.ChildNodes.Count - 1 do
      with Node do
        begin
          if ChildNodes[i].NodeName = 'Execution' then
             begin
               New(Execution);
               Execution^.Filename := GetNodeAttribute(ChildNodes[i], 'Filename');
               Execution^.Args := GetNodeAttribute(ChildNodes[i], 'Arguments');
               Execution^.Hidden := GetNodeAttribute(ChildNodes[i], 'Hidden') = True;
               Result.Add(Execution);
             end;
        end;
  end;

  function GetLinks(Node: TDOMNode): TList;
  var
    i: integer;
    Link: PLink;
  begin
    Result := TList.Create;
    for i := 0 to Node.ChildNodes.Count - 1 do
      with Node do
        begin
          if ChildNodes[i].NodeName = 'Link' then
             begin
               New(Link);
               Link^.Filename := GetNodeAttribute(ChildNodes[i], 'Filename');
               Link^.TargetName := GetNodeAttribute(ChildNodes[i], 'TargetName');
               Link^.WorkingDirectory := GetNodeAttribute(ChildNodes[i], 'WorkingDirectory');
               Result.Add(Link);
             end;
        end;
  end;

var
  i, j: Integer;
  NodeData: PApp;
  Node: PVirtualNode;
  tmpNode: TDOMNode;
  value: integer;
begin
  VST.BeginUpdate;

  if XMLNode.NodeName = 'Config' then
    begin
     TryStrToInt(GetNodeAttribute(XMLNode, 'Width'), value);
     frmMain.Width:=value;
     TryStrToInt(GetNodeAttribute(XMLNode, 'Height'), value);
     frmMain.Height:=value;
    end;

  For i := 0 to XMLNode.ChildNodes.Count - 1 do
    begin
    tmpNode := XMLNode.ChildNodes[i];
    Node := VSTParentNode;
    if tmpNode.NodeName = 'Application' then
      begin
        Node := VST.AddChild(VSTParentNode);
        NodeData := PApp(VST.GetNodeData(Node));

        NodeData^.State := STATE_NONE;
        NodeData^.Name := GetNodeAttribute(tmpNode, 'Name');
        NodeData^.Version := GetNodeAttribute(tmpNode, 'Version');
        NodeData^.Exclusive := GetNodeAttribute(tmpNode, 'Exclusive') = True;
        NodeData^.Checked := GetNodeAttribute(tmpNode, 'Checked') = True;
        NodeData^.Executions := GetExecutions(tmpNode);
        NodeData^.Links := GetLinks(tmpNode);

        if not ApplicationHasExecutions(NodeData) then NodeData^.State := STATE_NOT_FOUND;

        if NodeData^.Version = '' then
          if NodeData^.Executions.Count = 1 then
            if (ExtractFileExt(PExecution(NodeData^.Executions[0])^.Filename) <> '.msi') then
                 NodeData^.Version := GetAppVersionStr( ReplaceEnvs(PExecution(NodeData^.Executions[0])^.Filename) );

      end;
      if tmpNode.FindNode('Application') <> nil  then
           LoadNodes(tmpNode, Node);
    end;
  VST.EndUpdate;
end;

procedure TfrmMain.SaveNodes(Doc: TXMLDocument; Node: PVirtualNode; XMLNode: TDOMNode);
  function GetExecutionXMLNode(Doc: TXMLDocument; Execution: PExecution): TDOMNode;
  var
    NodeData: PApp;
    XMLNode: TDOMNode;
    Attribute: TDOMAttr;
  begin
    NodeData := VST.GetNodeData(Node);
    XMLNode := Doc.CreateElement('Execution');
    Attribute := Doc.CreateAttribute('Filename');
    Attribute.Value := Execution^.Filename;
    XMLNode.Attributes.SetNamedItem(Attribute);
    Attribute := Doc.CreateAttribute('Arguments');
    Attribute.Value := Execution^.Args;
    XMLNode.Attributes.SetNamedItem(Attribute);
    Attribute := Doc.CreateAttribute('Hidden');
    Attribute.Value := MyBoolToStr(Execution^.Hidden);
    XMLNode.Attributes.SetNamedItem(Attribute);

    Result := XMLNode;
  end;
  function GetLinkXMLNode(Doc: TXMLDocument; Link: PLink): TDOMNode;
  var
    NodeData: PApp;
    XMLNode: TDOMNode;
    Attribute: TDOMAttr;
  begin
    NodeData := VST.GetNodeData(Node);
    XMLNode := Doc.CreateElement('Link');
    Attribute := Doc.CreateAttribute('Filename');
    Attribute.Value := Link^.Filename;
    XMLNode.Attributes.SetNamedItem(Attribute);
    Attribute := Doc.CreateAttribute('TargetName');
    Attribute.Value := Link^.TargetName;
    XMLNode.Attributes.SetNamedItem(Attribute);
    Attribute := Doc.CreateAttribute('WorkingDirectory');
    Attribute.Value := Link^.WorkingDirectory;
    XMLNode.Attributes.SetNamedItem(Attribute);

    Result := XMLNode;
  end;
var
  _Node, ChildNode: PVirtualNode;
  _XMLNode, tmpNode: TDOMNode;
  NodeData: PApp;
  Attribute: TDOMAttr;
  i, j: integer;
begin
  _Node := Node;
  while Assigned(_Node) do
    begin
      NodeData := VST.GetNodeData(_Node);
      _XMLNode := Doc.CreateElement('Application');
      Attribute := Doc.CreateAttribute('Name');
      Attribute.Value:=NodeData^.Name;
      _XMLNode.Attributes.SetNamedItem(Attribute);
      Attribute := Doc.CreateAttribute('Version');
      Attribute.Value:=NodeData^.Version;
      _XMLNode.Attributes.SetNamedItem(Attribute);
      Attribute := Doc.CreateAttribute('Exclusive');
      Attribute.Value:=MyBoolToStr(NodeData^.Exclusive);
      _XMLNode.Attributes.SetNamedItem(Attribute);
      Attribute := Doc.CreateAttribute('Checked');
      Attribute.Value:=MyBoolToStr(NodeData^.Checked);
      _XMLNode.Attributes.SetNamedItem(Attribute);

      for i := 0 to NodeData^.Executions.Count - 1 do
        begin
          tmpNode := GetExecutionXMLNode(Doc, NodeData^.Executions[i]);
          _XMLNode.AppendChild(tmpNode);
        end;

      for i := 0 to NodeData^.Links.Count - 1 do
        begin
          tmpNode := GetLinkXMLNode(Doc, NodeData^.Links[i]);
          _XMLNode.AppendChild(tmpNode);
        end;

      XMLNode.Appendchild(_XMLNode);
      ChildNode := _Node^.FirstChild;
      if (_Node^.ChildCount > 0) then
        SaveNodes(Doc, ChildNode, _XMLNode);
      _Node := VST.GetNextSibling(_Node);
    end;
end;

function TfrmMain.InsertNodes(ParentNode: PVirtualNode; App: PApp): PVirtualNode;
var
 Node: PVirtualNode;
 NodeData: PApp;
begin
  Node := VST.AddChild(ParentNode);
  NodeData := PApp(VST.GetNodeData(Node));
  NodeData^ := App^;
  Result := Node;
end;

procedure TfrmMain.RunExecution(Execution: PExecution);
function isQuottedStr(StrToCheck: string): boolean;
var
 len: integer;
begin
 len := strlen(PChar(StrToCheck));
 if (StrToCheck[1] = '"') and (StrToCheck[len] = '"') then
   Result := true
 else
   Result := false;
end;

function QuoteStr(StrToQuote: string): string;
begin
 Result := '"' + StrToQuote + '"';
end;

function UnQuoteStr(StrToUnQuote: string): string;
begin
 Result := ReplaceStr(StrToUnQuote, '"', '');
end;

function SystemPath: string;
begin
    SetLength(Result, MAX_PATH);
    SetLength(Result, GetSystemDirectory(@Result[1], MAX_PATH));
end;
var

  Filename, Arguments, FileToCheck: string;
begin
  FileToCheck := Execution^.Filename;
  FileToCheck := ReplaceEnvs(FileToCheck);
  if isQuottedStr(FileToCheck) then
    FileToCheck := UnQuoteStr(FileToCheck);

  if ExtractFileExt(FileToCheck) = '.msi' then
    begin
      Filename := SystemPath + '\' + 'msiexec.exe';
      if isQuottedStr(Execution^.Filename) then
        Arguments := ReplaceEnvs('/i' + Execution^.Filename + ' ' + Execution^.Args)
      else
        Arguments := ReplaceEnvs('/i' + QuoteStr(Execution^.Filename) + ' ' + Execution^.Args);
    end
  else
    begin
      Arguments := ReplaceEnvs(Execution^.Args);
      Filename := ReplaceEnvs(Execution^.Filename);
    end;

  //if FileExists(FileToCheck) then
  //  begin
      ExecuteProcess(RawByteString(Filename), RawByteString(Arguments));
  //  end
  //else
  //  ShowMessage('Файл не найден: ' + Execution^.Filename);
end;

procedure TfrmMain.RunExecutionByProcess(Execution: PExecution);
function SystemPath: string;
begin
    SetLength(Result, MAX_PATH);
    SetLength(Result, GetSystemDirectory(@Result[1], MAX_PATH));
end;
var
  AProcess: TProcess;
  ShowWindowOptions: TShowWindowOptions;
  Output, Filename, Arguments: string;
begin
  AProcess := TProcess.Create(nil);

  AProcess.Options := [poWaitOnExit];

  if Execution^.Hidden then
    ShowWindowOptions := swoHIDE;

  AProcess.ShowWindow := ShowWindowOptions;

  if ExtractFileExt(Execution^.Filename) = '.msi' then
    begin
      Arguments := '/i' + Execution^.Filename + ' ' + Execution^.Args;
      Filename := SystemPath + '\' + 'msiexec.exe';

      AProcess.Parameters.Add('/i' + Execution^.Filename);
      AProcess.Parameters.Add(Arguments);
      AProcess.Executable := Filename;
    end
  else
    begin
      Arguments := Execution^.Args;
      Filename := Execution^.Filename;
      AProcess.Parameters.Add(Arguments);
      AProcess.CurrentDirectory := ExtractFileDir(Filename);
      AProcess.Executable := Filename;
    end;

  if FileExists(ReplaceEnvs(Execution^.Filename)) then
    begin
      AProcess.Execute;
    end
  else
    ShowMessage('Файл не найден: ' + Execution^.Filename);

  AProcess.Free;
end;

procedure TfrmMain.RunApplication(Application: PApp);
var
  i: integer;
begin
  for i := 0 to Application^.Executions.Count - 1 do
    begin
     RunExecution(PExecution(Application^.Executions[i]));
//      RunExecutionByProcess(PExecution(Application^.Executions[i]));
    end;

  for i := 0 to Application^.Links.Count - 1 do
    begin
     RunCreateLink(PLink(Application^.Links[i]));
    end;
end;

procedure TfrmMain.RunCreateLink(Link: PLink);
begin
  CreateLink(ReplaceEnvs(Link^.Filename), ReplaceEnvs(Link^.TargetName), ReplaceEnvs(Link^.WorkingDirectory));
end;

function TfrmMain.ReplaceEnvs(InputString: string): string;
var
 dt: string;
begin
  dt := GetDesktopDir;

  Result := InputString;
  Result := ReplaceStr(Result, '{src}', ExtractFileDir(ParamStr(0)));
  Result := ReplaceStr(Result, '{sd}', SysUtils.GetEnvironmentVariable('SystemDrive'));
  Result := ReplaceStr(Result, '{tmp}', SysUtils.GetEnvironmentVariable('TMP'));
  Result := ReplaceStr(Result, '{dt}', dt);
  Result := ReplaceStr(Result, '{appdata}', SysUtils.GetEnvironmentVariable('appdata'));
end;

function TfrmMain.ApplicationHasExecutions(Application: PApp): boolean;
var
  i: integer;
  Execution: PExecution;
begin
  Result := True;
  for i := 0 to Application^.Executions.Count - 1 do
    begin
      Execution := Application^.Executions[i];
      if (not FileExists(ReplaceEnvs(Execution^.Filename)) and (Execution^.Filename.StartsWith('{src}'))) then
        begin
          Result := false;
          break;
        end;
    end;
end;

procedure TfrmMain.ClearComponents;
begin
  edName.Text:='';
  edVersion.Text:='';
  cbExclusive.Checked:=false;
  cbChecked.Checked:=false;
  edFilename.Text:='';
  edArguments.Text:='';
  cbExecutions.Clear;
  cbLinks.Clear;
  edLinkFilename.Text:='';
  edLinkTargetname.Text:='';
  edLinkWorkDir.Text:='';
end;

procedure TfrmMain.ShowArgumentsHelp;
var
 info: string;
begin
 info := 'Inno Setup: /S' + #10#13 +
         'MSI: /qn' + #10#13 +
         'WiX Toolset: /quiet /norestart';
 ShowMessage(info);
end;

procedure TfrmMain.CheckcbExecutions;
begin
  edFilename.Enabled := cbExecutions.Items.Count > 0;
  edArguments.Enabled := cbExecutions.Items.Count > 0;
  btSaveExecution.Enabled := cbExecutions.Items.Count > 0;
end;

procedure TfrmMain.CheckcbLinks;
begin
  edLinkFilename.Enabled := cbLinks.Items.Count > 0;
  edLinkTargetname.Enabled := cbLinks.Items.Count > 0;
  edLinkWorkDir.Enabled := cbLinks.Items.Count > 0;
  btSaveLink.Enabled := cbLinks.Items.Count > 0;
end;

end.

