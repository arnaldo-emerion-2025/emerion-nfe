unit uJsonHelper;

interface

uses
  ulkJSON,
  pcnConversaoNFe,
  dialogs,
  SysUtils;

procedure readJson(const jsonString: String);
function getTipoConsumidor(const jsonObj: TlkJSONobject): TpcnConsumidorFinal;
function getStringSafe(const jsonObj: TlkJSONobject; const field: String): String;
function getNumberSafe(const jsonObj: TlkJSONobject; const field: String): Extended;

implementation

procedure readJson(const jsonString: String);
var
  jsonObj: TlkJSONobject;
begin
  jsonObj := TlkJSON.ParseText(jsonString) as TlkJSONobject;
  try
    ShowMessage(jsonObj.Field['inicio'].Field['SeqNFe'].Value);
  finally
    jsonObj.Free;
  end;
end;

function getTipoConsumidor(const jsonObj: TlkJSONobject): TpcnConsumidorFinal;
var
  ok: Boolean;
begin
  if not (jsonObj.Field['ide'].Count > 0) then
    raise Exception.Create('NFe IDE info not available');

  Result := StrToConsumidorFinal(ok, jsonObj.Field['ide'].Field['indFinal'].Value)
end;

function getStringSafe(const jsonObj: TlkJSONobject; const field: String): String;
var
 val: TlkJSONbase;
begin
   val := jsonObj.Field['nfat'];
   if(val <> nil) then
     Result := val.Value
   else
     Result := '';
end;

function getNumberSafe(const jsonObj: TlkJSONobject; const field: String): Extended;
var
 val: TlkJSONbase;
begin
   val := jsonObj.Field['nfat'];
   if(val <> nil) then
     Result := val.Value
   else
     Result := 0;
end;

end.
