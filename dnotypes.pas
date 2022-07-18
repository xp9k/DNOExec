unit DnoTypes;

{$mode objfpc}{$H+}

interface

uses
    Classes, Windows;

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
  RegRootKey: HKEY;
  RegPath: string;
  DisplayName: string;
  Version: string;
  UninstallString: string;
  FilePath: string;
  Args: string;
  ArgsList: array of string;
end;

const
  HKEYNames: array[0..6] of string =
    ('HKEY_CLASSES_ROOT', 'HKEY_CURRENT_USER', 'HKEY_LOCAL_MACHINE', 'HKEY_USERS',
    'HKEY_PERFORMANCE_DATA', 'HKEY_CURRENT_CONFIG', 'HKEY_DYN_DATA');

function HKEYToStr(const Key: HKEY): string;


implementation

function HKEYToStr(const Key: HKEY): string;
begin
  if (key > HKEY_CLASSES_ROOT + 6) then
    Result := ''
  else
    Result := HKEYNames[key - HKEY_CLASSES_ROOT];
end;

end.

