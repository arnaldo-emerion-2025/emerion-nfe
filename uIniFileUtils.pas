unit uIniFileUtils;

interface

uses
  IniFiles, SysUtils, Forms;

type
  TWebService = class
    uf: String;
    ambiente: Integer;
  end;

  TWebServiceFazenda = class
    urlBase: String;
    calcularTributos: String;
    gerarXml: String;
    validarXml: String;
  end;

  TNfeEmerionIni = class
    webservice: TWebService;
    tipoSistema: String;
    utilizarNovoImposto: Boolean;
    webserviceFazenda: TWebServiceFazenda;

    constructor Create;
  end;

procedure AtualizaIni;
procedure ValidaIni;

var
  nfeEmerionIni: TNfeEmerionIni;

implementation

var
  Ini: TIniFile;

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
      arqIni.WriteString('Geral', 'FileImpressao', arqIni.readString('Geral', 'PathSchemas', '') + '\DANFeRetrato.fr3');
    end;
  finally

    arqIni.Free
  end;

end;

{ TNfeEmerionIni }

constructor TNfeEmerionIni.Create;
begin
  inherited;
  webservice := TWebService.Create;
  webserviceFazenda := TWebServiceFazenda.Create;
end;

initialization

begin
  try
    Ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
    nfeEmerionIni := TNfeEmerionIni.Create;
    try
      nfeEmerionIni.webservice.uf := Ini.readString('WebService', 'UF', '');
      nfeEmerionIni.webservice.ambiente := Ini.readString('WebService', 'Ambiente', '').ToInteger;

      nfeEmerionIni.tipoSistema := Ini.readString('TIPO_SISTEMA', 'TIPO_ARQUIVO', '');
      nfeEmerionIni.utilizarNovoImposto := SameText(Ini.readString('TIPO_SISTEMA', 'UTILIZAR_NOVO_IMPOSTO', 'Nao'.ToUpper), 'SIM');
      nfeEmerionIni.webserviceFazenda.urlBase := Ini.readString('WEBSERVICE_FAZENDA', 'URL_BASE', '');
      nfeEmerionIni.webserviceFazenda.calcularTributos := Ini.readString('WEBSERVICE_FAZENDA', 'CALCULAR_TRIBUTOS', '');
      nfeEmerionIni.webserviceFazenda.gerarXml := Ini.readString('WEBSERVICE_FAZENDA', 'GERAR_XML', '');
      nfeEmerionIni.webserviceFazenda.validarXml := Ini.readString('WEBSERVICE_FAZENDA', 'VALIDAR_XML', '');
    finally
      Ini.Free;
    end;
  except
    on E: Exception do
      raise Exception.Create('Failed to read INI file: ' + E.Message);
  end;
end;

end.
