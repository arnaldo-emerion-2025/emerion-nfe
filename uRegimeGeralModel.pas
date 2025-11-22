unit uRegimeGeralModel;

interface

uses
  SysUtils, Classes, Contnrs, ulkJSON;

type
  TTributacaoRegular = class(TObject)
  public
    cst: string;
    cClassTrib: string;
    function ToJSON: TlkJSONObject;
  end;

  TImpostoSeletivo = class(TObject)
  public
    cst: string;
    baseCalculo: Double;
    cClassTrib: string;
    unidade: string;
    quantidade: Double;
    impostoInformado: Double;
    function ToJSON: TlkJSONObject;
  end;

  TItem = class(TObject)
  public
    numero: Integer;
    ncm: string;
    nbs: string;
    quantidade: Double;
    unidade: string;
    cst: string;
    baseCalculo: Double;
    cClassTrib: string;
    tributacaoRegular: TTributacaoRegular;
    impostoSeletivo: TImpostoSeletivo;
    constructor Create;
    destructor Destroy; override;
    function ToJSON: TlkJSONObject;
  end;

  TRegimeGeral = class(TObject)
  public
    id: string;
    versao: string;
    dataHoraEmissao: string;
    municipio: Integer;
    uf: string;
    itens: TObjectList;
    constructor Create;
    destructor Destroy; override;
    function ToJSON: TlkJSONObject;
    function ToJSONString: string;
  end;

implementation

{ TTributacaoRegular }

function TTributacaoRegular.ToJSON: TlkJSONObject;
begin
  Result := TlkJSONObject.Create;
  Result.Add('cst', cst);
  Result.Add('cClassTrib', cClassTrib);
end;

{ TImpostoSeletivo }

function TImpostoSeletivo.ToJSON: TlkJSONObject;
begin
  Result := TlkJSONObject.Create;
  Result.Add('cst', cst);
  Result.Add('baseCalculo', baseCalculo);
  Result.Add('cClassTrib', cClassTrib);
  Result.Add('unidade', unidade);
  Result.Add('quantidade', quantidade);
  Result.Add('impostoInformado', impostoInformado);
end;

{ TItem }

constructor TItem.Create;
begin
  inherited Create;
  tributacaoRegular := TTributacaoRegular.Create;
  impostoSeletivo := TImpostoSeletivo.Create;
end;

destructor TItem.Destroy;
begin
  impostoSeletivo.Free;
  tributacaoRegular.Free;
  inherited Destroy;
end;

function TItem.ToJSON: TlkJSONObject;
begin
  Result := TlkJSONObject.Create;
  Result.Add('numero', numero);
  Result.Add('ncm', ncm);
  Result.Add('quantidade', quantidade);
  Result.Add('unidade', unidade);
  Result.Add('cst', cst);
  Result.Add('baseCalculo', baseCalculo);
  Result.Add('cClassTrib', cClassTrib);
  Result.Add('tributacaoRegular', tributacaoRegular.ToJSON);
  Result.Add('impostoSeletivo', impostoSeletivo.ToJSON);
end;

{ TRegimeGeral }

constructor TRegimeGeral.Create;
begin
  inherited Create;
  itens := TObjectList.Create(True);
end;

destructor TRegimeGeral.Destroy;
begin
  itens.Free;
  inherited Destroy;
end;

function TRegimeGeral.ToJSON: TlkJSONObject;
var
  i: Integer;
  jList: TlkJSONlist;
  itemObj: TItem;
begin
  Result := TlkJSONObject.Create;
  Result.Add('id', id);
  Result.Add('versao', versao);
  Result.Add('dataHoraEmissao', dataHoraEmissao);
  Result.Add('municipio', municipio);
  Result.Add('uf', uf);

  jList := TlkJSONlist.Create;
  for i := 0 to itens.Count - 1 do
  begin
    itemObj := TItem(itens[i]);
    jList.Add(itemObj.ToJSON);
  end;
  Result.Add('itens', jList);
end;

function TRegimeGeral.ToJSONString: string;
var
  jObj: TlkJSONObject;
begin
  jObj := ToJSON;
  try
    Result := TlkJSON.GenerateText(jObj);
  finally
    jObj.Free;
  end;
end;

end.
