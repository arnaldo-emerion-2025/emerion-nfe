unit uJsonHelper;

interface

uses
  ulkJSON,
  dialogs;

procedure readJson(const jsonString: String);

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

end.
