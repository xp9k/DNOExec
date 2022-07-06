unit DnoTypes;

{$mode objfpc}{$H+}

interface

uses
    Classes;

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
  UpdateURL: string;
  Exclusive: boolean;
  Checked: boolean;
  Parent: PApp;
  Executions: TList;
  Links: TList;
  State: Integer;
end;

PUninstall = ^TUninstall;
TUninstall = record
  DisplayName: string;
  Version: string;
  UninstallString: string;
  FilePath: string;
  Args: string;
end;


implementation

end.

