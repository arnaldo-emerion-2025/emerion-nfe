unit uWebServiceTributos;

interface

uses
  System.SysUtils,
  System.IniFiles,
  System.Classes,
  System.Contnrs,
  Forms,
  uRegimeGeralModel,
  Dialogs,
  Winapi.Windows,
  Winapi.ActiveX,
  uWebServiceUtils;

type
  TWebServiceTributos = class
  public
    class procedure calcularNovosImpostos();
    class function generateRegimeGeralData(): string;
  end;

var
  APIBaseURL: string = '';
  endpointCalcularTributos: string = '';
  endpointGerarXML: string = '';
  endpointValidarXML: string = '';

implementation

var
  Ini: TIniFile;

  { TWebServiceTributos }

class procedure TWebServiceTributos.calcularNovosImpostos;
var
  jsonArquivoEntrada, responseString, xmlGenerated: String;
begin
  jsonArquivoEntrada := generateRegimeGeralData;

  responseString := TRestClientHelper.PostJSON(APIBaseURL + endpointCalcularTributos, [], jsonArquivoEntrada);

  ShowMessage(responseString);

  xmlGenerated := TRestClientHelper.PostJSON(APIBaseURL + endpointGerarXML, ['tipo=nfe'], responseString);

  ShowMessage(xmlGenerated);
end;

class function TWebServiceTributos.generateRegimeGeralData: string;
var
  entrada: TRegimeGeral;
  item: TItem;
begin
  entrada := TRegimeGeral.Create;
  entrada.id := '1';
  entrada.versao := '1.0.0';
  entrada.dataHoraEmissao := '2026-01-01T09:50:05-03:00';
  entrada.municipio := 4314902;
  entrada.uf := 'RS';

  // create an item and set properties
  item := TItem.Create;
  item.numero := 1;
  item.ncm := '24021000';
  item.nbs := '';
  item.quantidade := 222;
  item.unidade := 'KG';
  item.cst := '000';
  item.baseCalculo := 200;
  item.cClassTrib := '000001';

  item.tributacaoRegular.cst := '000';
  item.tributacaoRegular.cClassTrib := '000000';

  item.impostoSeletivo.cst := '000';
  item.impostoSeletivo.baseCalculo := 1111;
  item.impostoSeletivo.cClassTrib := '000000';
  item.impostoSeletivo.unidade := 'KG';
  item.impostoSeletivo.quantidade := 222;
  item.impostoSeletivo.impostoInformado := 0;

  entrada.itens.Add(item);

  Result := entrada.ToJSONString;
end;

initialization

begin
  try
    Ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
    try
      APIBaseURL := Ini.ReadString('WEBSERVICE_FAZENDA', 'URL_BASE', '');
      endpointCalcularTributos := Ini.ReadString('WEBSERVICE_FAZENDA', 'CALCULAR_TRIBUTOS', '');
      endpointGerarXML := Ini.ReadString('WEBSERVICE_FAZENDA', 'GERAR_XML', '');
      endpointValidarXML := Ini.ReadString('WEBSERVICE_FAZENDA', 'VALIDAR_XML', '');
    finally
      Ini.Free;
    end;
  except
    on E: Exception do
      raise Exception.Create('Failed to read INI file: ' + E.Message);
  end;
end;

end.
