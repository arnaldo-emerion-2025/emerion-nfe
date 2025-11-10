unit uIniFileUtils;

interface

uses
  IniFiles, SysUtils, Forms;

procedure AtualizaIni;
procedure ValidaIni;

implementation

procedure ValidaIni;
var
  IniFile: string;
  arqIni: TIniFile;
begin
  IniFile := ChangeFileExt(Application.ExeName, '.ini');
  arqIni := TIniFile.Create(IniFile);
  arqIni.Free;
end;

procedure AtualizaIni;
var
  IniFile: string;
  arqIni: TIniFile;
begin
  IniFile := ChangeFileExt(Application.ExeName, '.ini');

  arqIni := TIniFile.Create(IniFile);
  try
    if (arqIni.readString('Geral', 'TipoImpressao', '') = '') then
    begin
      arqIni.WriteString('Geral', 'TipoImpressao', 'FAST');
    end;

    if (arqIni.readString('Geral', 'DialogoImpressao', '') = '') then
    begin
      arqIni.WriteString('Geral', 'DialogoImpressao', 'SIM');
    end;

    if (arqIni.readString('Geral', 'FileImpressao', '') = '') then
    begin
      arqIni.WriteString('Geral', 'FileImpressao', arqIni.readString('Geral',
        'PathSchemas', '') + '\DANFeRetrato.fr3');
    end;
  finally

    arqIni.Free
  end;

end;

end.
