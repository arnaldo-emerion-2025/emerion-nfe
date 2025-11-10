program NFeEmerion2;

uses
  Forms,
  Uprincipal in 'Uprincipal.pas' {Form1},
  uMsg in 'uMsg.pas',
  uFuncoes in 'uFuncoes.pas',
  uIniFileUtils in 'uIniFileUtils.pas',
  uJsonHelper in 'uJsonHelper.pas',
  uNFeJson in 'uNFeJson.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
