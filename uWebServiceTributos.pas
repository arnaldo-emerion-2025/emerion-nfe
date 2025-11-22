unit uWebServiceTributos;

interface

uses
  System.SysUtils,
  System.IniFiles,
  Forms,
  uRegimeGeralModel,
  Dialogs,
  Winapi.Windows,
  Winapi.ActiveX,
  ComObj;

type
  TRestClientHelper = class
  public
    class procedure calcularNovosImpostos();
    class function PostJSON(const Url, ContentType, Body: string): string;
    class function generateRegimeGeralData(): string;
  end;

var
  APIBaseURL: string = '';
  endpointCalcularTributos: string = '';
  endpointGerarXML: string = '';
  endpointValidarXML: string = '';

const
  WinHttpRequestOption_SslErrorIgnoreFlags = 3;
  WinHttpRequestOption_EnableRedirects = 6;
  WinHttpRequestOption_SecurityFlags = 4;
  WinHttpRequestOption_SecureProtocols = 9;

  SECURITY_FLAG_IGNORE_CERT_CN_INVALID = $00001000;
  SECURITY_FLAG_IGNORE_CERT_DATE_INVALID = $00002000;

  WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2 = $00000800;

implementation

var
  Ini: TIniFile;

  { TRestClientHelper }

class procedure TRestClientHelper.calcularNovosImpostos;
var
  jsonArquivoEntrada, responseString: String;
begin
  jsonArquivoEntrada := generateRegimeGeralData;

  responseString := PostJSON(APIBaseURL + endpointCalcularTributos, 'application/json', jsonArquivoEntrada);

  ShowMessage(responseString);
end;

class function TRestClientHelper.generateRegimeGeralData: string;
var
  entrada: TRegimeGeral;
  item: TItem;
  jsonText: string;
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

class function TRestClientHelper.PostJSON(const Url, ContentType, Body: string): string;
var
  WinHttp: OleVariant;
begin
  WinHttp := CreateOleObject('WinHttp.WinHttpRequest.5.1');

  // Open connection - TRUE for async off (we want sync)
  WinHttp.Open('POST', Url, False);

  // Force TLS 1.2 (if supported by the OS)
  WinHttp.Option[WinHttpRequestOption_SslErrorIgnoreFlags] := 0;
  WinHttp.Option[WinHttpRequestOption_EnableRedirects] := True;
  WinHttp.Option[WinHttpRequestOption_SecurityFlags] := SECURITY_FLAG_IGNORE_CERT_CN_INVALID or
    SECURITY_FLAG_IGNORE_CERT_DATE_INVALID;

  // Set headers
  WinHttp.SetRequestHeader('Content-Type', ContentType);
  WinHttp.SetRequestHeader('Accept', 'application/json');

  // Send body
  WinHttp.Send(Body);

  // Return response as string
  Result := WinHttp.ResponseText;
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
