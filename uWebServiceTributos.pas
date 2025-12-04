unit uWebServiceTributos;

interface

uses
  ulkJSON,
  System.SysUtils,
  System.IniFiles,
  System.Classes,
  System.Contnrs,
  Forms,
  uRegimeGeralModel,
  Dialogs,
  Winapi.Windows,
  Winapi.ActiveX,
  uWebServiceUtils,
  uGeneratedXMLUtil,
  ACBrNFe.Classes;

type
  TWebServiceTributos = class
  public
    class function calcularNovosImpostos(items: TDetCollection; jsonObj: TlkJSONobject): TInfNFeNovosTributos;
  private
    class function generateRegimeGeralData(items: TDetCollection; jsonObj: TlkJSONobject): string;
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

class function TWebServiceTributos.calcularNovosImpostos(items: TDetCollection; jsonObj: TlkJSONobject)
  : TInfNFeNovosTributos;
var
  jsonArquivoEntrada, responseString, xmlGenerated: String;
  response: THttpResponse;
  Inf: TInfNFeNovosTributos;
begin
  jsonArquivoEntrada := generateRegimeGeralData(items, jsonObj);

  {
    First step defined
    Calcular Tributos da RTC
  }
  responseString := TRestClientHelper.PostJSON(APIBaseURL + endpointCalcularTributos, [], jsonArquivoEntrada).Body;

  {
    Second step defined
    Gerar XML dos Grupos de Tributacao da RTC
  }
  xmlGenerated := TRestClientHelper.PostJSON(APIBaseURL + endpointGerarXML, ['tipo=nfe'], responseString).Body;

  {
    Third step defined
    Validar XML Gerado
  }
  response := TRestClientHelper.PostJSON(APIBaseURL + endpointValidarXML, ['tipo=nfe', 'subtipo=grupo'], xmlGenerated,
    'application/xml');

  if (response.Status <> 200) then
    raise Exception.Create('Erro enquanto consultando calculadora do Sefaz: ' + response.Body);

  Inf := TInfNFeNovosTributos.Create;
  Inf.LoadFromXML(xmlGenerated);

  Result := Inf;
end;

class function TWebServiceTributos.generateRegimeGeralData(items: TDetCollection; jsonObj: TlkJSONobject): string;
var
  entrada: TRegimeGeral;
  item: TItem;
  det: TDetCollectionItem;
  i: Integer;
  detArray: TlkJSONlist;
  detItem: TlkJSONobject;
begin
  entrada := TRegimeGeral.Create;
  entrada.id := '1';
  entrada.versao := '1.0.0';
  entrada.dataHoraEmissao := '2026-01-01T09:50:05-03:00';
  entrada.municipio := 3550308;
  entrada.uf := 'SP';

  detArray := jsonObj.Field['det'] as TlkJSONlist;

  for i := 0 to items.Count - 1 do
  begin
    det := items[i];
    detItem := detArray.Child[i] as TlkJSONobject;

    item := TItem.Create;
    item.numero := det.Prod.nItem;
    item.ncm := det.Prod.ncm;
    item.quantidade := det.Prod.qCom;
    item.unidade := det.Prod.uCom;
    item.cst := detItem.Field['imposto'].Field['IBSCBS'].Field['cst'].Value;
    item.baseCalculo := det.Prod.vProd;
    item.cClassTrib := detItem.Field['imposto'].Field['IBSCBS'].Field['cClassTrib'].Value;;
    entrada.itens.Add(item);
  end;

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
