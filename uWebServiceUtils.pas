unit uWebServiceUtils;

interface

uses
  ComObj, System.SysUtils;

type
  THttpResponse = record
    Status: Integer;
    Body: string;
  end;

type
  TRestClientHelper = class
  public
    class function EncodeURLParam(const S: string): string;
    class function BuildQueryString(const Params: array of string): string;
    class function PostJSON(const Url: String; const Params: array of string; Body: string; contentType : String = 'application/json'): THttpResponse;
  end;

const
  WinHttpRequestOption_SslErrorIgnoreFlags = 3;
  WinHttpRequestOption_EnableRedirects = 6;
  WinHttpRequestOption_SecurityFlags = 4;
  WinHttpRequestOption_SecureProtocols = 9;

  SECURITY_FLAG_IGNORE_CERT_CN_INVALID = $00001000;
  SECURITY_FLAG_IGNORE_CERT_DATE_INVALID = $00002000;

  WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2 = $00000800;

implementation

class function TRestClientHelper.EncodeURLParam(const S: string): string;
var
  I: Integer;
begin
  Result := '';
  for I := 1 to Length(S) do
  begin
    if S[I] in ['A'..'Z', 'a'..'z', '0'..'9', '-', '_', '.', '~'] then
      Result := Result + S[I]
    else
      Result := Result + '%' + IntToHex(Ord(S[I]), 2);
  end;
end;

class function TRestClientHelper.BuildQueryString(const Params: array of string): string;
var
  I, P: Integer;
  Key, Value: string;
begin
  Result := '';

  for I := Low(Params) to High(Params) do
  begin
    P := Pos('=', Params[I]);
    if P > 0 then
    begin
      Key := Copy(Params[I], 1, P - 1);
      Value := Copy(Params[I], P + 1, MaxInt);

      if Result <> '' then
        Result := Result + '&';

      Result := Result +
        EncodeURLParam(Key) + '=' + EncodeURLParam(Value);
    end;
  end;
end;

class function TRestClientHelper.PostJSON(const Url: String; const Params: array of string; Body: string; contentType : String = 'application/json'): THttpResponse;
var
  WinHttp: OleVariant;
  FullUrl, Query: string;
begin
  Query := BuildQueryString(Params);

  if Query <> '' then
    FullUrl := Url + '?' + Query
  else
    FullUrl := Url;

  WinHttp := CreateOleObject('WinHttp.WinHttpRequest.5.1');

  // Open connection - TRUE for async off (we want sync)
  WinHttp.Open('POST', FullUrl, False);

  // Force TLS 1.2 (if supported by the OS)
  WinHttp.Option[WinHttpRequestOption_SslErrorIgnoreFlags] := 0;
  WinHttp.Option[WinHttpRequestOption_EnableRedirects] := True;
  WinHttp.Option[WinHttpRequestOption_SecurityFlags] := SECURITY_FLAG_IGNORE_CERT_CN_INVALID or
    SECURITY_FLAG_IGNORE_CERT_DATE_INVALID;

  // Set headers
  WinHttp.SetRequestHeader('Content-Type', contentType);

  // Send body
  WinHttp.Send(Body);

  // Return response as string
  Result.Status := WinHttp.Status;
  Result.Body := WinHttp.ResponseText;
end;

end.
