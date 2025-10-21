unit Uprincipal;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Variants,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  OleCtrls,
  SHDocVw,
  ACBrNFe,
  IniFiles,
  ShellAPI,
  midaslib,
  clipbrd,
  DB,
  DBClient,
  ExtCtrls,
  ACBrNFeDANFEClass,
  ACBrBase,
  ACBrValidador,
  DSCommonServer,
  DSHTTPCommon,
  DSHTTPWebBroker,
  Wininet,
  StrUtils,
  uFuncoes,
  ACBrUtil,
  ComCtrls,
  ACBrNFeDANFEFR,
  frxClass,
  frxExportPDF,
  WideStrings,
  DBXFirebird,
  SqlExpr,
  SimpleDS, ACBrNFeDANFEFRDM, ACBrDFe, pcnConversao, System.Math,
  ACBrDFeReport, ACBrDFeDANFeReport, frxExportXLS, frxExportImage,
  frxExportHTML;

type
  TForm1 = class(TForm)
    ClientDataSet1: TClientDataSet;
    Timer1: TTimer;
    Memo1: TMemo;
    Memo2: TMemo;
    Label1: TLabel;
    ACBrValidador1: TACBrValidador;
    frxReport1: TfrxReport;
    SQLConnection1: TSQLConnection;
    WBResposta: TWebBrowser;
    frxPDFExport1: TfrxPDFExport;
    procedure ACBrNFe1GerarLog(const ALogLine: string; var Tratado: Boolean);

  type
    TCliente = record
      isDenegada: Boolean;
      ie: String;
      cgc: String;
    end;

  procedure FormCreate(Sender: TObject);

  procedure Timer1Timer(Sender: TObject);

  private
    { Private declarations }
    strxNome: string;
    strDenegada: String;
    function validaCliente(cgc, uf: String): TCliente;
    function RecuperaChaveEnviando: string;
    procedure ReescreveChaveEnviada(strChave, strProtocolo: string;
      strDEPCSefaz: string = '');
    procedure VerificaInternet;
    function EnviaNFe2(RChave, RProtocolo: String; SN: Boolean;
      TipoEnvio: integer = 3): Boolean;
    function EnviaNFe2DB(TipoEnvio: integer = 3): Boolean;
    function LimpaStr(str: String): String;
    procedure ContigenciaDPEC;
    procedure DPEC_SEFAZ;
    procedure ConfirmaEnvioDPEC(strEnviado: string);
    procedure GravaProtcoloDpec(strPTo: string);
    procedure ValidaINI;

    function SelectCabecalho(strTAB, ID_TAB: string): string;
    function SelectItem(strItem, ID_TAB: string): string;
    function SelectTitulos(strTitulos, ID_TAB: string): string;
    function LeftZero(Const Text: string; Const Tam: word;
      Const RetQdoVazio: String = ' '): string;

    // Antigos botões do NFEEMERION
    procedure bConfigClick;
    procedure bEnviarNfeClick;
    procedure bEnviarNfeClick2;
    procedure bCancelaNFeClick;
    procedure bCancelaNFeEventoClick;
    procedure bComplementoClick;
    procedure bConsultaNFeClick;
    procedure bDanfeClick;
    procedure bInutilizaClick;
    procedure bEnviarSNClick;
    procedure bComplementoSNClick;
    procedure bConsultaDisponibilidade;
    procedure AtualizaIni;

  public

    VAux, VERRO, Chave: string;
    Foi: Boolean;
    // ========= variaveis Certificado VCert
    VCertCaminho, VCertSenha, VCertNumSerie: string;

    // ========= variaveis Config. Gerais   VCGerais
    VCGeraisLogo, VCGeraisCaminhoArquivosXML, VCGeraisCaminhoArquivoLeitura, //
      VCGeraisCaminhoArquivoRetorno: string;
    VCGeraisCaminhoArquivoDownload: string;
    VCGeraisCaminhoArquivoCancelada, VCGeraisCaminhoArquivoDANFE, //
      VCGeraisCaminhoArquivoSchemas: string;
    VCGeraisCaminhoArquivoTipoImpressao,
      VCGeraisCaminhoArquivoFileImpressao: string;
    VCGeraisDanfe, VCGeraisFormaemissao: integer;
    VCGeraisSalvar: Boolean;

    // ========= Variaveis WebSevice VWebS
    VWebSUF: string;
    VWebSAmbiente: integer;
    VWebSVisualizar: Boolean;
    VSIte: string;

    // ========= Variaveis PROXY VProxy
    VProxyHost, VProxyUsuario, VProxySenha, VProxyPorta: string;

    VNumLote: integer;

    ParNF: integer;
    ACAO: string;
    vchave: string;

    // Chave Recuperada do Sefaz
    RecChave: String;
    ACBrNFe1: TACBrNFe;
    ACBrNFeDANFEFR1: TACBrNFeDANFEFR;

    procedure GravarConfiguracao;
    procedure LerConfiguracao;
    procedure LoadXML(MyMemo: TMemo; MyWebBrowser: TWebBrowser);
    { Public declarations }
    function RecuperarXMLFATPED(RChave, RProtocolo: string; SN: Boolean;
      TipoEnvio: integer = 3): Boolean;
    function ValidaCNPJ(sCNPJ: string; MostraMsg: Boolean = true): Boolean;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses
  pcnNFe,
  // ACBrNFebNotasFiscais,
  ConvUtils,
  ACBrNFeWebServices,
  pcnConversaoNFe, ACBrDFeSSL, blcksock;

function VALIDAcpf(cpf: string): Boolean;
var
  i: integer;
  Want: char;
  Wvalid: Boolean;
  Wdigit1, Wdigit2: integer;
begin
  Result := False;
  if trim(cpf) <> '' then
  begin

    Wdigit1 := 0;
    Wdigit2 := 0;
    Want := cpf[1];
    // variavel para testar se o cpf é repetido como 111.111.111-11
    Delete(cpf, ansipos('.', cpf), 1); // retira as mascaras se houver
    Delete(cpf, ansipos('.', cpf), 1);
    Delete(cpf, ansipos('-', cpf), 1);

    Wvalid := False;
    // testar se o cpf é repetido como 111.111.111-11
    for i := 1 to length(cpf) do
    begin
      if cpf[i] <> Want then
      begin
        Wvalid := true;
        // se o cpf possui um digito diferente ele passou no primeiro teste
        break
      end;
    end;
    // se o cpf é composto por numeros repetido retorna falso
    if not Wvalid then
    begin
      Result := False;
      exit;
    end;

    // executa o calculo para o primeiro verificador
    for i := 1 to 9 do
    begin
      Wdigit1 := Wdigit1 + (strtoint(cpf[10 - i]) * (i + 1));
    end;
    Wdigit1 := ((11 - (Wdigit1 mod 11)) mod 11) mod 10;
    { formula do primeiro verificador
      soma=1°*2+2°*3+3°*4.. até 9°*10
      digito1 = 11 - soma mod 11
      se digito > 10 digito1 =0
    }

    // verifica se o 1° digito confere
    if IntToStr(Wdigit1) <> cpf[10] then
    begin
      Result := False;
      exit;
    end;

    for i := 1 to 10 do
    begin
      Wdigit2 := Wdigit2 + (strtoint(cpf[11 - i]) * (i + 1));
    end;
    Wdigit2 := ((11 - (Wdigit2 mod 11)) mod 11) mod 10;
    { formula do segundo verificador
      soma=1°*2+2°*3+3°*4.. até 10°*11
      digito1 = 11 - soma mod 11
      se digito > 10 digito1 =0
    }

    // confere o 2° digito verificador
    if IntToStr(Wdigit2) <> cpf[11] then
    begin
      Result := False;
      exit;
    end;

    // se chegar até aqui o cpf é valido
    Result := true;
  end;
end;

function TForm1.ValidaCNPJ(sCNPJ: string; MostraMsg: Boolean = true): Boolean;
begin
  ACBrValidador1.Documento := sCNPJ;
  Result := ACBrValidador1.Validar;
end;

procedure TForm1.ValidaINI;
var
  IniFile: string;
  arqIni: TIniFile;
begin

  IniFile := ChangeFileExt(Application.ExeName, '.ini');

  arqIni := TIniFile.Create(IniFile);

  try

  finally
    arqIni.Free;
  end;

end;

procedure TForm1.VerificaInternet;
var
  Flags: DWord;
begin
  { if not InternetGetConnectedState(@Flags, 0) then
    ShowMessage('Você não está conectado à Internet.'); }
end;

function TForm1.RecuperaChaveEnviando: string;
var
  arq: TIniFile;
  AuxStr: string;
begin
  AuxStr := ExtractFilePath(Application.ExeName) + 'Config.ini';
  arq := TIniFile.Create(AuxStr);
  try
    AuxStr := trim(arq.readString('PASSNFE', 'CHAVE', ''));
    Result := AuxStr;
  finally
    FreeAndnil(arq);
  end;

end;

function TForm1.RecuperarXMLFATPED(RChave, RProtocolo: string; SN: Boolean;
  TipoEnvio: integer = 3): Boolean;
var
  Arquivo: TextFile;
  Linha: string;
  ArrumaDuplicada: TextFile;
  id_estsip: string;
  ArquivoSL: Tstringlist;
  VSNprod, VSNAliq: Real;
  Teste: Boolean;
begin

  EnviaNFe2(RChave, RProtocolo, SN, TipoEnvio);

  try
    try
      Memo2.Lines.Clear;
      ACBrNFe1.Enviar(VNumLote, False);
    finally
      Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
        IntToStr(VNumLote) + '.txt');
    end;
  except
    on E: Exception do
    begin
      // Validando retorno para NFe Denegada
      if (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 110) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 205) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 301) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 302) then
      begin
        strDenegada := 'S';
        ValidaRejeicao(IntToStr(ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.cStat) + ' - ' + ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.xMotivo);
      end;

      // Teste := ACBrNFe1.NotasFiscais.Items[0].SaveToFile(ACBrNFe1.NotasFiscais.Items[0].NomeArq);

      if ((ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 539) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 502) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 204) or
        (pos(IntToStr(VNumLote) + '->Rejeição: Duplicidade de NF-e', E.Message)
        > 0)) then
        Result := true
      else
        Result := False;
    end;
  end;

  { if Result then
    begin
    ArquivoSL := Tstringlist.Create;
    ArquivoSL.Clear;
    ArquivoSL.LoadFromFile(ACBrNFe1.NotasFiscais.Items[0].NomeArq);
    AssignFile(ArrumaDuplicada, ACBrNFe1.NotasFiscais.Items[0].NomeArq);
    Rewrite(ArrumaDuplicada);
    write(ArrumaDuplicada, '<?xml version="1.0" encoding="UTF-8"?><nfeProc versao="3.10" xmlns="http://www.portalfiscal.inf.br/nfe">');
    write(ArrumaDuplicada, ArquivoSL.Text);
    write(ArrumaDuplicada, RProtocolo + '</nfeProc>');
    CloseFile(ArrumaDuplicada);
    ArquivoSL.Free;
    end
    else
    DeleteFile(pchar(ACBrNFe1.NotasFiscais.Items[0].NomeArq));
  }

  ACBrNFe1.NotasFiscais.Clear;
  ACBrNFe1.NotasFiscais.LoadFromFile(VCGeraisCaminhoArquivoRetorno + '\' +
    RChave + '-nfe.xml');
  ACBrNFe1.Consultar;

  // if (ACBrNFe1.WebServices.Consulta.Protocolo = RProtocolo) then
  if (pos(ACBrNFe1.WebServices.Consulta.Protocolo, RProtocolo) > 0) then
  begin
    // ShowMessage('Condicao Verdadeira');
  end;

  ACBrNFe1.NotasFiscais.Items[0].GravarXML(RChave + '-nfe.xml');
  ACBrNFe1.NotasFiscais.Items[0].GravarXML
    (IntToStr(ACBrNFe1.NotasFiscais.Items[0].NFe.Ide.nNF) + ' - NF-e- ' + RChave
    + '.xml');

end;

procedure TForm1.ReescreveChaveEnviada(strChave, strProtocolo: string;
  strDEPCSefaz: string = '');
var
  arq: TIniFile;
  AuxStr: string;
begin
  AuxStr := ExtractFilePath(Application.ExeName) + 'Config.ini';
  arq := TIniFile.Create(AuxStr);
  try
    arq.WriteString('PASSNFE', 'CHAVE', strChave);
    arq.WriteString('PASSNFE', 'PROTOCOLO', strProtocolo);
    arq.WriteString('PASSNFE', 'DENEGADA', strDenegada);
    arq.WriteString('PASSNFE', 'DEPC_SEFAZ', strDEPCSefaz);

  finally
    FreeAndnil(arq);
  end;

end;

function TForm1.SelectCabecalho(strTAB: string; ID_TAB: string): string;
begin
  Result := ' Select FatPed.CodEmp, ' + //
    ' FatPced.DteRes, ' + //
    ' FatPed.NumRes, ' + //
    ' FatPed.SeqLib, ' + //
    ' FatPed.SeqFat, ' + //
    ' FatPed.QtiFat, ' + //
    ' FatPed.DteFat, ' + //
    ' FatPed.UfeFat, ' + //
    ' FatPed.NroNfs, ' + //
    ' FatPed.CodPfa, ' + //
    ' FatPed.TipPfa, ' + //
    ' FatPed.CodCf1, ' + //
    ' FatPed.CodCf2, ' + //
    ' FatPed.CodCli, ' + //
    ' FatPed.FlgSai, ' + //
    ' FatPed.FlgEnt, ' + //
    ' FatPed.TipFrt, ' + //
    ' FatPed.EspFat, ' + //
    ' FatPed.MarFat, ' + //
    ' FatPed.IntFin, ' + //
    ' FatPed.DesNat, ' + //
    ' FatPed.InsSub, ' + //
    ' FatPed.CodTra, ' + //
    ' FatPed.TraSda, ' + //
    ' FatPed.NomTra, ' + //
    ' FatPed.CgcTra, ' + //
    ' FatPed.InsTra, ' + //
    ' FatPed.TenTra, ' + //
    ' FatPed.EndTra, ' + //
    ' FatPed.RefTra, ' + //
    ' FatPed.NumTra, ' + //
    ' FatPed.BaiTra, ' + //
    ' FatPed.CidTra, ' + //
    ' FatPed.UfeTra, ' + //
    ' FatPed.CepTra, ' + //
    ' FatPed.NroFat, ' + //
    ' FatPed.PlcTra, ' + //
    ' FatPed.UfePlc, ' + //
    ' FatPed.TefCli, ' + //
    ' FatPed.EnfCli, ' + //
    ' FatPed.RffCli, ' + //
    ' FatPed.NrfCli, ' + //
    ' FatPed.BafCli, ' + //
    ' FatPed.CifCli, ' + //
    ' FatPed.UffCli, ' + //
    ' FatPed.CefCli, ' + //
    ' FatPed.TenCli, ' + //
    ' FatPed.EndCli, ' + //
    ' FatPed.RefCli, ' + //
    ' FatPed.NumCli, ' + //
    ' FatPed.BaiCli, ' + //
    ' FatPed.CidCli, ' + //
    ' FatPed.UfeCli, ' + //
    ' FatPed.CepCli, ' + //
    ' FatPed.CgeCli, ' + //
    ' FatPed.IneCli, ' + //
    ' FatPed.InfLiq, ' + //
    ' FatPed.InfBrt, ' + //
    ' FatPed.AltVol, ' + //
    ' FatPed.LotNfe, ' + //
    ' FatPed.EnvNfe, ' + //
    ' FatPed.SeqNfe, ' + //
    ' FatPed.DteNfe, ' + //
    ' FatPed.RecNfe, ' + //
    ' FatPed.ProNfe, ' + //
    ' FatPed.DopNfe, ' + //
    ' FatPed.HreNfe, ' + //
    ' FatPed.UsuNfe, ' + //
    ' FatPed.DtePnf, ' + //
    ' FatPed.HrePnf, ' + //
    ' FatPed.ImpNfe, ' + //
    ' FatPed.RetNfe, ' + //
    ' FatPed.FlgAtu, ' + //
    ' FatPed.Id_FinUff, ' + //
    ' FatPed.Id_FinCif, ' + //
    ' FatPed.Id_FinUfe, ' + //
    ' FatPed.Id_FinCie, ' + //
    ' FatPed.Id_TraUfe, ' + //
    ' FatPed.Id_TraCie, ' + //
    ' FatPed.TrbPis, ' + //
    ' FatPed.PerPis, ' + //
    ' FatPed.TrbCof, ' + //
    ' FatPed.PerCof, ' + //
    ' FatPed.TotFat, ' + //
    ' FatPed.TotDsr, ' + //
    ' FatPed.TotFrt, ' + //
    ' FatPed.TotSeg, ' + //
    ' FatPed.TotDes, ' + //
    ' FatPed.TotIpi, ' + //
    ' FatPed.TotPis, ' + //
    ' FatPed.TotCof, ' + //
    ' FatPed.BasIcm, ' + //
    ' FatPed.TotIcm, ' + //
    ' FatPed.BasSub, ' + //
    ' FatPed.TotSub, ' + //
    ' FatPed.TotGer, ' + //
    ' FatPed.Ob1Fat, ' + //
    ' FatPed.Ob2Fat, ' + //
    ' FatPed.Ob3Fat, ' + //
    ' FatPed.Ob4Fat, ' + //
    ' FatPed.Ob5Fat, ' + //
    ' FatPed.Ob6Fat, ' + //
    ' FatPed.Ob7Fat, ' + //
    ' FatPed.Ob8Fat, ' + //
    ' FatPed.NfePis, ' + //
    ' FatPed.NfeCof, ' + //
    ' FatPed.Id_EstSip, ' + //
    ' FatPed.FlgDenegada, ' + //
    ' FinCli.NomCli, ' + //
    ' FinCli.Em1Cli, ' + //
    ' fatped.flgimp, ' + //
    ' fatped.flgnfe, ' + //
    ' fatped.id_fatped, ' + //
    ' FatPed.SitFat, ' + //
    ' FatPed.ENVDPEC, ' + //
    ' FatPed.USUDPEC, ' + //
    ' FatPed.JustDPEC, ' + //
    ' FatPed.ProtDPEC, ' + //
    ' FatPed.Libera_Resp, ' + //
    ' FatPed.LocEmb, ' + //
    ' FatPed.UFEmb, ' + //
    ' FinCli.Pt1Cli||FinCli.Fo1Cli TelCli, ' + //
    ' 0.0 TotImpII ' + //
    ' From FatPed LEFT ' + //
    ' JOIN FinCli ON (FatPed.CodCli = FinCli.CodCli) ' + //
    ' Where FatPed.id_fatped = ' + ID_TAB + //
    ' Order by FatPed.NroNfs '; //

end;

function TForm1.SelectItem(strItem: string; ID_TAB: string): string;
begin
  Result := ' Select FatPe2.NroPe2,' + //
    ' ESTPRO.CBAPRO cEANTRIB, ' + //
    ' ESTPRO.CBAEMB cEAN, ' + //
    ' ESTPRO.CBAEMB DESIMP, ' + //
    ' FatPe2.CodClp,' + //
    ' FatPe2.CodGru,' + //
    ' FatPe2.CodSub,' + //
    ' FatPe2.CodPro,' + //
    ' FatPe2.REFPE2,' + //
    ' FatPe2.DesPe2,' + //
    ' FatPe2.ObsPe2,' + //
    ' FatPe2.ClsIpi,' + //
    ' FatPe2.CodCfo,' + //
    ' FatPe2.CodSt1,' + //
    ' FatPe2.CodSt2,' + //
    ' FatPe2.CodUnd,' + //
    ' FatPe2.QtpPe2,' + //
    ' FatPe2.VlqPe2 VluPe2,' + //
    ' FatPe2.TotPe2,' + //
    ' FatPe2.IcmPe2,' + //
    ' FatPe2.BscIcm,' + //
    ' FatPe2.RedIcm,' + //
    ' FatPe2.BasIcm,' + //
    ' FatPe2.TotIcm,' + //
    ' FatPe2.IpiPe2,' + //
    ' FatPe2.CSTIPI,' + //
    ' FatPe2.TrbIpi,' + //
    ' FatPe2.BscIpi,' + //
    ' FatPe2.RedIpi,' + //
    ' FatPe2.BasIpi,' + //
    ' FatPe2.TotIpi,' + //
    ' FatPe2.IcmSub,' + //
    ' FatPe2.MrgSub,' + //
    ' FatPe2.BaseSb,' + //
    ' FatPe2.BasSub,' + //
    ' FatPe2.TotSub,' + //
    ' FatPe2.TotDsr,' + //
    ' FatPe2.TotFrt,' + //
    ' FatPe2.TotSeg,' + //
    ' FatPe2.TotDes,' + //
    ' FatPe2.BASPIS,' + //
    ' FatPe2.CSTPIS,' + //
    ' FatPe2.TOTPIS,' + //
    ' FatPe2.ALIQPIS,' + //
    ' FatPe2.BASCOF,' + //
    ' FatPe2.CSTCOF,' + //
    ' FatPe2.AliqCof,' + //
    ' FatPe2.TotCof,' + //
    ' FatPe2.NUMPEDCOMPRA,' + //
    ' FatPe2.NUMITEMCOMPRA,' + //
    ' Estite.VPFITE' + //
    ' From FatPe2' + //
    ' Join ESTPRO on CODGRU = fatpe2.codgru and codsub = fatpe2.codsub and codpro = fatpe2.codpro and codclp = fatpe2.codclp '
    + ' Join ESTITE on CODGRU = fatpe2.codgru and codsub = fatpe2.codsub and codpro = fatpe2.codpro and codclp = fatpe2.codclp '
    +
  //
    ' Where FatPe2.id_FatPed = ' + ID_TAB + //
    ' Order by FatPe2.NroPe2';
end;

function TForm1.SelectTitulos(strTitulos, ID_TAB: string): string;
begin
  Result := ' Select FatPe3.NroPe3,' + ' FatPe3.DtvPe3,' +
    ' FatPe3.VlpPe3 From FatPe3' + ' Where FatPe3.CodEmp = ' + ID_TAB +
    ' Order by FatPe3.NroPe3';

end;

procedure TForm1.GravaProtcoloDpec(strPTo: string);
var
  arq: TIniFile;
  AuxStr: string;
begin
  AuxStr := ExtractFilePath(Application.ExeName) + 'Config.ini';
  arq := TIniFile.Create(AuxStr);
  try
    arq.WriteString('DPEC', 'PROTOCOLO', strPTo);

  finally
    FreeAndnil(arq);
  end;

end;

procedure TForm1.GravarConfiguracao;
var
  IniFile: string;
  Ini: TIniFile;
  StreamMemo: TMemoryStream;
begin
  IniFile := ChangeFileExt(Application.ExeName, '.ini');

  Ini := TIniFile.Create(IniFile);
  try
    // ====================== CERTIFICADO ==============================================

    Ini.WriteString('Certificado', 'Caminho', VCertCaminho);
    Ini.WriteString('Certificado', 'Senha', VCertSenha);
    Ini.WriteString('Certificado', 'NumSerie', VCertNumSerie);

    // ====================== GERAL ====================================================

    Ini.WriteInteger('Geral', 'DANFE', VCGeraisDanfe);
    // 0=RETRATO  / 1=PAISAGEM
    Ini.WriteInteger('Geral', 'FormaEmissao', VCGeraisFormaemissao);
    // 0=NORMAL  / 1=CONTIGÊNCIA  / 2=SCAN  / 3=DPEC  / 4=FSDA
    Ini.WriteString('Geral', 'LogoMarca', VCGeraisLogo);
    Ini.WriteBool('Geral', 'Salvar', true);
    // Salvar os arquivos de envio e resposta
    Ini.WriteString('Geral', 'PathXML', VCGeraisCaminhoArquivosXML);
    if DirectoryExists(VCGeraisCaminhoArquivosXML) = False then
      ForceDirectories(VCGeraisCaminhoArquivosXML);

    Ini.WriteString('Geral', 'PathLeitura', VCGeraisCaminhoArquivoLeitura);
    if DirectoryExists(VCGeraisCaminhoArquivoLeitura) = False then
      ForceDirectories(VCGeraisCaminhoArquivoLeitura);

    Ini.WriteString('Geral', 'PathRetorno', VCGeraisCaminhoArquivoRetorno);
    if DirectoryExists(VCGeraisCaminhoArquivoRetorno) = False then
      ForceDirectories(VCGeraisCaminhoArquivoRetorno);

    // ====================== WEBSERVICE ===============================================

    Ini.WriteString('WebService', 'UF', VWebSUF);
    Ini.WriteInteger('WebService', 'Ambiente', VWebSAmbiente);
    // 0=produção  / 1=homologação
    Ini.WriteBool('WebService', 'Visualizar', true);

    // ====================== PROXY   ===============================================
    Ini.WriteString('Proxy', 'Host', VProxyHost);
    Ini.WriteString('Proxy', 'Porta', VProxyPorta);
    Ini.WriteString('Proxy', 'User', VProxyUsuario);
    Ini.WriteString('Proxy', 'Pass', VProxySenha);

  finally
    Ini.Free;
  end;

end;

procedure TForm1.LerConfiguracao;
var
  IniFile: string;
  Ini: TIniFile;
  Ok: Boolean;
begin
  IniFile := ChangeFileExt(Application.ExeName, '.ini');

  Ini := TIniFile.Create(IniFile);
  try

    // ACBrNFe1.Configuracoes.Certificados.Senha        := Ini.ReadString( 'Certificado','Senha'   ,'') ;
    // ===== Achando o Certificado pelo Caminho ou pelo numero de série
{$IFDEF ACBrNFeOpenSSL}
    ACBrNFe1.Configuracoes.Certificados.Certificado :=
      Ini.readString('Certificado', 'Caminho', '');
    edtNumSerie.Visible := False;
    Label25.Visible := False;
    sbtnGetCert.Visible := False;
{$ELSE}
    ACBrNFe1.Configuracoes.Certificados.NumeroSerie :=
      Ini.readString('Certificado', 'NumSerie', '');
{$ENDIF}
    // ===== Setando as configurações Gerais
    VCGeraisFormaemissao := Ini.ReadInteger('Geral', 'FormaEmissao', 0);
    VCGeraisSalvar := Ini.ReadBool('Geral', 'Salvar', true);
    VCGeraisCaminhoArquivosXML := Ini.readString('Geral', 'PathXML', '');
    VCGeraisCaminhoArquivoLeitura := Ini.readString('Geral', 'PathLeitura', '');
    VCGeraisCaminhoArquivoRetorno := Ini.readString('Geral', 'PathRetorno', '');
    VCGeraisCaminhoArquivoCancelada := Ini.readString('Geral',
      'PathCancelada', '');
    VCGeraisCaminhoArquivoDANFE := Ini.readString('Geral', 'PathDANFE', '');
    VCGeraisCaminhoArquivoSchemas := Ini.readString('Geral', 'PathSchemas', '');

    VCGeraisCaminhoArquivoTipoImpressao :=
      Ini.readString('Geral', 'TipoImpressao', 'RAVE');
    VCGeraisCaminhoArquivoFileImpressao :=
      Ini.readString('Geral', 'FileImpressao', '');

    // Validando as Pastas necessárias para o envio da NFe
    if not DirectoryExists(VCGeraisCaminhoArquivosXML) then
    begin
      messagebox(Handle, pchar('Caminho para a pasta XML (' +
        VCGeraisCaminhoArquivosXML +
        ') não encontrado. Verifique e tente novamente.'), 'NFE Emerion',
        MB_OK + MB_ICONEXCLAMATION);
      Application.Terminate;
    end;

    if not DirectoryExists(VCGeraisCaminhoArquivoLeitura) then
    begin
      messagebox(Handle, pchar('Caminho para a pasta Leitura (' +
        VCGeraisCaminhoArquivoLeitura +
        ') não encontrado. Verifique e tente novamente.'), 'NFE Emerion',
        MB_OK + MB_ICONEXCLAMATION);
      Application.Terminate;
    end;

    if not DirectoryExists(VCGeraisCaminhoArquivoRetorno) then
    begin
      messagebox(Handle, pchar('Caminho para a pasta Retorno (' +
        VCGeraisCaminhoArquivoRetorno +
        ') não encontrado. Verifique e tente novamente.'), 'NFE Emerion',
        MB_OK + MB_ICONEXCLAMATION);
      Application.Terminate;
    end;

    if not DirectoryExists(VCGeraisCaminhoArquivoCancelada) then
    begin
      messagebox(Handle, pchar('Caminho para a pasta Cancelada (' +
        VCGeraisCaminhoArquivoCancelada +
        ') não encontrado. Verifique e tente novamente.'), 'NFE Emerion',
        MB_OK + MB_ICONEXCLAMATION);
      Application.Terminate;
    end;

    if not DirectoryExists(VCGeraisCaminhoArquivoDANFE) then
    begin
      messagebox(Handle, pchar('Caminho para a pasta Danfe (' +
        VCGeraisCaminhoArquivoDANFE +
        ') não encontrado. Verifique e tente novamente.'), 'NFE Emerion',
        MB_OK + MB_ICONEXCLAMATION);
      Application.Terminate;
    end;

    if not DirectoryExists(VCGeraisCaminhoArquivoSchemas) then
    begin
      messagebox(Handle, pchar('Caminho para a pasta Schemas (' +
        VCGeraisCaminhoArquivoSchemas +
        ') não encontrado. Verifique e tente novamente.'), 'NFE Emerion',
        MB_OK + MB_ICONEXCLAMATION);
      Application.Terminate;
    end;
    // Validação
    ACBrNFe1.Configuracoes.Geral.FormaEmissao :=
      StrToTpEmis(Ok, IntToStr(VCGeraisFormaemissao + 1));
    ACBrNFe1.Configuracoes.Geral.Salvar := VCGeraisSalvar;
    ACBrNFe1.Configuracoes.Arquivos.PathSalvar := VCGeraisCaminhoArquivosXML;
    ACBrNFe1.Configuracoes.Arquivos.PathSchemas :=
      VCGeraisCaminhoArquivoSchemas;

    ACBrNFe1.Configuracoes.Geral.ModeloDF := moNFe;
    ACBrNFe1.Configuracoes.Geral.VersaoDF := ve400;
    // ACBrNFe1.Configuracoes.Geral.SSLLib := libCapicom;
    ACBrNFe1.Configuracoes.Geral.SSLLib := libWinCrypt;
    ACBrNFe1.Configuracoes.Geral.ForcarGerarTagRejeicao938 := fgtSempre;
    ACBrNFe1.SSL.SSLType := LT_TLSv1_2;

    ACBrNFe1.Configuracoes.Arquivos.PathInu := VCGeraisCaminhoArquivoCancelada;
    ACBrNFe1.Configuracoes.Arquivos.PathNFe := VCGeraisCaminhoArquivoRetorno;

    // ACBrNFe1.Configuracoes.Arquivos.PathCan := VCGeraisCaminhoArquivoRetorno;
    ACBrNFe1.Configuracoes.Arquivos.PathInu := VCGeraisCaminhoArquivoRetorno;
    // ACBrNFe1.Configuracoes.Arquivos.PathDPEC := VCGeraisCaminhoArquivoRetorno;
    // ACBrNFe1.Configuracoes.Arquivos.PathCCe := VCGeraisCaminhoArquivoRetorno;
    // ACBrNFe1.Configuracoes.Arquivos.PathMDe := VCGeraisCaminhoArquivoRetorno;
    ACBrNFe1.Configuracoes.Arquivos.PathEvento := VCGeraisCaminhoArquivoRetorno;
    ACBrNFe1.Configuracoes.Arquivos.DownloadDFe.PathDownload :=
      VCGeraisCaminhoArquivoDownload;

    // ====== Setando as configuraçoes WebService
    VSIte := Ini.readString('WebService', 'SITE', '');
    VWebSUF := Ini.readString('WebService', 'UF', '');
    VWebSAmbiente := Ini.ReadInteger('WebService', 'Ambiente', 0);
    VWebSVisualizar := Ini.ReadBool('WebService', 'Visualizar', False);

    // ACBrNFeDANFERave1.Site := VSIte;
    // ACBrNFeDANFERaveCB1.Site := VSIte;
    ACBrNFeDANFEFR1.Site := VSIte;

    ACBrNFe1.Configuracoes.WebServices.uf := VWebSUF;
    if VWebSAmbiente = 1 then
      ACBrNFe1.Configuracoes.WebServices.Ambiente := taHomologacao
    else
      ACBrNFe1.Configuracoes.WebServices.Ambiente := taProducao;

    ACBrNFe1.Configuracoes.WebServices.Visualizar := VWebSVisualizar;

    // ===== Setando as configurações de Proxy
    VProxyHost := Ini.readString('Proxy', 'Host', '');
    VProxyPorta := Ini.readString('Proxy', 'Porta', '');
    VProxyUsuario := Ini.readString('Proxy', 'User', '');
    VProxySenha := Ini.readString('Proxy', 'Pass', '');
    ACBrNFe1.Configuracoes.WebServices.ProxyHost := VProxyHost;
    ACBrNFe1.Configuracoes.WebServices.ProxyPort := VProxyPorta;
    ACBrNFe1.Configuracoes.WebServices.ProxyUser := VProxyUsuario;
    ACBrNFe1.Configuracoes.WebServices.ProxyPass := VProxySenha;

    // Setando as Configuraçoes Gerais
    VCGeraisDanfe := Ini.ReadInteger('Geral', 'DANFE', 0);
    VCGeraisLogo := Ini.readString('Geral', 'LogoMarca', '');
    if ACBrNFe1.DANFE <> nil then
    begin
      ACBrNFe1.DANFE.TipoDANFE := StrToTpImp(Ok, IntToStr(VCGeraisDanfe + 1));
      ACBrNFe1.DANFE.Logo := VCGeraisLogo;
    end;

    if trim(VCGeraisCaminhoArquivoFileImpressao) = '' then
    begin
      messagebox(0,
        'Não foi informado tipo de impressão. Verifique e tente novamente.',
        'Emissão de NFe', MB_OK + MB_ICONINFORMATION);
      Application.Terminate;
    end;

    // ACBrNFeDANFERave1.PathPDF := VCGeraisCaminhoArquivoDANFE;
    // ACBrNFeDANFERaveCB1.PathPDF := VCGeraisCaminhoArquivoDANFE;
    ACBrNFeDANFEFR1.PathPDF := VCGeraisCaminhoArquivoDANFE;

    { if (VCGeraisCaminhoArquivoTipoImpressao = 'RAVE') then
      begin
      // Antigo
      // ACBrNFeDANFERave1.RavFile := VCGeraisCaminhoArquivoSchemas + '\NotaFiscalEletronica.rav';

      // Confirma se existe arquivo.
      if fileexists(VCGeraisCaminhoArquivoFileImpressao) then
      begin
      ACBrNFeDANFERave1.RavFile := VCGeraisCaminhoArquivoFileImpressao;
      ACBrNFe1.DANFE := ACBrNFeDANFERave1;
      end
      else
      if fileexists(VCGeraisCaminhoArquivoSchemas + '\NotaFiscalEletronica.rav') then
      begin
      ACBrNFeDANFERave1.RavFile := VCGeraisCaminhoArquivoSchemas + '\NotaFiscalEletronica.rav';
      ACBrNFe1.DANFE := ACBrNFeDANFERave1;
      end
      else
      begin
      messagebox(0, pwidechar('Não localizado o arquivo de impressão do DANFE: ' +
      VCGeraisCaminhoArquivoTipoImpressao + '. Verifique e tente novamente.'), 'Emissão de NFe',
      MB_OK + MB_ICONINFORMATION);
      Application.Terminate;
      end;
      end
      else }
    // if (VCGeraisCaminhoArquivoTipoImpressao = 'FAST') then
    // begin
    // Antigo
    // ACBrNFeDANFERave1.RavFile := VCGeraisCaminhoArquivoSchemas + '\NotaFiscalEletronica.rav';

    // Confirma se existe arquivo.
    if fileexists(VCGeraisCaminhoArquivoFileImpressao) then
    begin
      ACBrNFeDANFEFR1.FastFile := VCGeraisCaminhoArquivoFileImpressao;
      ACBrNFe1.DANFE := ACBrNFeDANFEFR1;
    end
    else
    begin
      messagebox(0, pwidechar('Não localizado o arquivo de impressão do DANFE: '
        + VCGeraisCaminhoArquivoTipoImpressao +
        '. Verifique e tente novamente.'), 'Emissão de NFe',
        MB_OK + MB_ICONINFORMATION);
      Application.Terminate;
    end;
    // end;

  finally
    Ini.Free;
  end;

end;

function TForm1.LimpaStr(str: String): String;
const
  caracteres: array [0 .. 5] of char = ('/', '.', ',', '-', '\', ' ');
var
  i: integer;
begin
  str := trim(str);

  for i := 0 to 5 do
  begin
    str := ReplaceStr(str, caracteres[i], '');
  end;

  Result := str;

end;

procedure TForm1.LoadXML(MyMemo: TMemo; MyWebBrowser: TWebBrowser);
begin
  MyMemo.Lines.SaveToFile(ExtractFileDir(Application.ExeName) + 'temp.xml');
  MyWebBrowser.Navigate(ExtractFileDir(Application.ExeName) + 'temp.xml');
end;

procedure TForm1.bConfigClick;
begin
  LerConfiguracao;
  ACBrNFe1.NotasFiscais.Clear;
end;

procedure TForm1.bEnviarNfeClick;
var
  Arquivo: TextFile;
  Linha: string;
  i, j, colunai, colunaf: integer;
  achou: Boolean;
  chaveDup, protocoloDupl, aux: string;
  ArrumaDuplicada: TextFile;
  Cdi: integer;
  intVezes: integer;
  strNome: String;
  dataMaior: Boolean;
begin
  achou := False;
  dataMaior := strTODate('01/11/2025') > Date;
  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\EVNOTA' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-ERRO' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  except
    Application.Terminate;
  end;

  // Busca informações do Arquivo de Leitura
  try
    EnviaNFe2('', '', False);
  Except
    on E: Exception do
    begin
      ShowMessage(E.Message);
    end;
  end;

  try

    Foi := true;

    try
      Memo2.Lines.Clear;
      ACBrNFe1.NotasFiscais.GravarXML(VCGeraisCaminhoArquivosXML +
        '\xmlantesdeenviar.xml');
      // ACBrNFe1.NotasFiscais[0].XML;
      ACBrNFe1.Enviar(VNumLote, False);
    finally
      Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
        IntToStr(VNumLote) + '.txt');
    end;

    ACBrNFe1.NotasFiscais.ImprimirPDF;

    // Validando retorno para NFe Denegada
    if (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 110) or
      (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 205) or
      (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 301) or
      (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 302) then
    begin
      strDenegada := 'S';

      strNome := VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(VNumLote) +
        ' - Denegada- ' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items
        [0].chDFe + '.xml';

      CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' +
        ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
        '-nfe.xml'), pchar(strNome), true);

      ValidaRejeicao(IntToStr(ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat)
        + ' - ' + ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.xMotivo);
    end;

  except
    on E: Exception do
    begin
      // Validando retorno para NFe Denegada
      if (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 110) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 205) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 301) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 302) then
      begin
        strDenegada := 'S';

        strNome := VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(VNumLote) +
          ' - Denegada- ' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.
          Items[0].chDFe + '.xml';

        CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' +
          ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
          '-nfe.xml'), pchar(strNome), true);

        ValidaRejeicao(IntToStr(ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.cStat) + ' - ' + ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.xMotivo);
      end;

      Memo1.Lines.Text := (E.Message);
      Foi := False;
      chaveDup := '';

      // ==============================

      for i := 0 to Memo1.Lines.Count - 1 do
      begin
        colunai := pos('[', Memo1.Lines[i]);
        colunaf := pos(']', Memo1.Lines[i]);

        if ((colunai > 0) and (colunaf > colunai)) then
        begin
          chaveDup := '';
          Linha := Memo1.Lines[i];

          for j := colunai + 1 to colunaf - 1 do
            if Linha[j] in ['0' .. '9'] then
            begin
              chaveDup := chaveDup + Linha[j];
            end;

          if length(chaveDup) <> 44 then
          begin
            achou := False;
            Foi := False;
          end
          else
          begin
            achou := true;
          end;
        end;
      end;

      { if achou then
        messagebox(Handle, pchar(E.Message), 'Envio NFe', mb_OK); }

      if achou then
      begin
        ACBrNFe1.NotasFiscais.Clear;
        ACBrNFe1.WebServices.Consulta.NFeChave := chaveDup;
        ACBrNFe1.WebServices.Consulta.Executar;
        aux := UTF8Encode(ACBrNFe1.WebServices.Consulta.RetWS);
        colunai := pos('<protNFe', aux);
        colunaf := pos('</protNFe>', aux);
        protocoloDupl := '';

        for j := colunai to colunaf + 9 do
          protocoloDupl := protocoloDupl + aux[j];

        Clipboard.Open;
        try
          Clipboard.AsText := chaveDup;
        finally
          Clipboard.Close;
        end;

        if messagebox(Handle,
          pchar('Numeração de NFe já enviada anteriormente com a chave: ' +
          chaveDup + #13 + //
          'Caso seja a mesma NFe você pode optar por Sim para recuperar o XML.'
          + #13 +
          //
          'Em caso de dúvida poderá consultar no site ' + //
          'http://www.nfe.fazenda.gov.br' + #13 + //
          'Para facilitar a chave se encontra na área de transferência. Basta utilizar CTRL+V para colar a chave no campo de pesquisa.'),
          pchar('Envio de NFe'), MB_YESNO + MB_ICONQUESTION) = IDYES then
        begin

          if RecuperarXMLFATPED(chaveDup, protocoloDupl, False) then
          begin
            ACBrNFe1.NotasFiscais.LoadFromFile
              (ACBrNFe1.Configuracoes.Arquivos.PathSalvar + '\' +
              IntToStr(ParNF) + '-env-lot.xml');
            // ACBrNFe1.Configuracoes.Arquivos.PathNFe
            // if (ACBrNFe1.NotasFiscais.Count > 0) then
            Foi := true;
          end
          else
            Foi := False;
        end
        else
          Foi := False;
      end;

    end;
  end;

  // =============================================

  if not Foi then
  begin

    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
      IntToStr(VNumLote) + '.txt');
    ReescreveChaveEnviada('', '');

    Application.Terminate;
    SysUtils.Abort;
    Application.Terminate;

  end
  else
  begin

    if (ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].nProt <> '') or
      (chaveDup <> '') then
    begin

      Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Retorno.RetWS);

      /// LoadXML(Memo1, WBResposta);

      intVezes := 0;

      if (trim(chaveDup) = '') then
      begin
        strNome := VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(VNumLote) +
          ' - NF-e- ' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0]
          .chDFe + '.xml';
        // strNome := VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(VNumLote) + ' - NF-e- ' + chaveDup + '.xml';
        CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' +
          ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
          '-nfe.xml'), pchar(strNome), true);
        // showmessage(strNome);
      end;

      RecChave := RecuperaChaveEnviando;

      if RecChave <> ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe
      then
        ReescreveChaveEnviada(ifthen(achou, //
          chaveDup, //
          ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe), //
          ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].nProt);
    end
    else
    begin
      ReescreveChaveEnviada('', '');
    end;
  end;
end;

procedure TForm1.bEnviarNfeClick2;
var
  Arquivo: TextFile;
  Linha: string;
  i, j, colunai, colunaf: integer;
  achou: Boolean;
  chaveDup, protocoloDupl, aux: string;
  ArrumaDuplicada: TextFile;
  Cdi: integer;
  intVezes: integer;
  strNome: String;
begin
  achou := False;

  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\EVNOTA' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-ERRO' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  except
    Application.Terminate;
  end;

  // Busca informações do Arquivo de Leitura
  try
    EnviaNFe2('', '', False);
    ACBrNFe1.NotasFiscais.Validar;
    ACBrNFe1.NotasFiscais.Assinar;
  Except
    on E: Exception do
    begin
      ShowMessage(E.Message);
    end;
  end;
  // ACBrNFe1.NotasFiscais.GravarTXT;
  ACBrNFe1.NotasFiscais.Items[0].GravarXML('XML_' + IntToStr(ParNF) + '.xml');

  (* try

    Foi := true;

    try
    Memo2.Lines.Clear;
    ACBrNFe1.Enviar(VNumLote, False);
    finally
    Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' + IntToStr(VNumLote) + '.txt');
    end;

    ACBrNFe1.NotasFiscais.ImprimirPDF;

    except
    on E: Exception do
    begin
    // Validando retorno para NFe Denegada
    if (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 110) or (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 205) or
    (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 301) or (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 302) then
    begin
    strDenegada := 'S';

    strNome := VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(VNumLote) + ' - Denegada- ' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0]
    .chNFe + '.xml';

    CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe + '-nfe.xml'), pchar(strNome),
    true);

    ValidaRejeicao(IntToStr(ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat) + ' - ' + ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.xMotivo);
    end;

    Memo1.Lines.Text := (E.Message);
    Foi := False;
    chaveDup := '';

    // ==============================

    for i := 0 to Memo1.Lines.Count - 1 do
    begin
    colunai := pos('[', Memo1.Lines[i]);
    colunaf := pos(']', Memo1.Lines[i]);

    if ((colunai > 0) and (colunaf > colunai)) then
    begin
    chaveDup := '';
    Linha := Memo1.Lines[i];
    for j := colunai + 1 to colunaf - 1 do
    if Linha[j] in ['0' .. '9'] then
    begin
    chaveDup := chaveDup + Linha[j];
    end;

    if length(chaveDup) <> 44 then
    begin
    achou := False;
    Foi := False;
    end
    else
    begin
    achou := true;
    end;
    end;
    end;

    { if achou then
    messagebox(Handle, pchar(E.Message), 'Envio NFe', mb_OK); }

    if achou then
    begin
    ACBrNFe1.NotasFiscais.Clear;
    ACBrNFe1.WebServices.Consulta.NFeChave := chaveDup;
    ACBrNFe1.WebServices.Consulta.Executar;
    aux := UTF8Encode(ACBrNFe1.WebServices.Consulta.RetWS);
    colunai := pos('<protNFe', aux);
    colunaf := pos('</protNFe>', aux);
    protocoloDupl := '';

    for j := colunai to colunaf + 9 do
    protocoloDupl := protocoloDupl + aux[j];

    Clipboard.Open;
    try
    Clipboard.AsText := chaveDup;
    finally
    Clipboard.Close;
    end;

    if messagebox(Handle, pchar('Numeração de NFe já enviada anteriormente com a chave: ' + chaveDup + #13 + //
    'Caso seja a mesma NFe você pode optar por Sim para recuperar o XML.' + #13 +
    //
    'Em caso de dúvida poderá consultar no site ' + //
    'http://www.nfe.fazenda.gov.br' + #13 + //
    'Para facilitar a chave se encontra na área de transferência. Basta utilizar CTRL+V para colar a chave no campo de pesquisa.'),
    pchar('Envio de NFe'), MB_YESNO + MB_ICONQUESTION) = IDYES then
    begin

    if RecuperarXMLFATPED(chaveDup, protocoloDupl, False) then
    begin
    ACBrNFe1.NotasFiscais.LoadFromFile(ACBrNFe1.Configuracoes.Arquivos.GetPathNFe(now) + '\' + chaveDup + '-nfe.xml');
    Foi := true;
    end
    else
    Foi := False;
    end
    else
    Foi := False;
    end;

    end;
    end;

    // =============================================

    if not Foi then
    begin

    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' + IntToStr(VNumLote) + '.txt');
    ReescreveChaveEnviada('', '');

    Application.Terminate;
    SysUtils.Abort;
    Application.Terminate;

    end
    else
    begin

    if (ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].nProt <> '') or (chaveDup <> '') then
    begin

    Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Retorno.RetWS);

    /// LoadXML(Memo1, WBResposta);

    intVezes := 0;
    strNome := VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(VNumLote) + ' - NF-e- ' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0]
    .chNFe + '.xml';

    CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe + '-nfe.xml'), pchar(strNome), true);
    // showmessage(strNome);
    RecChave := RecuperaChaveEnviando;

    if RecChave <> ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe then
    ReescreveChaveEnviada(ifthen(achou, //
    chaveDup, //
    ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe), //
    ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].nProt);
    end
    else
    begin
    ReescreveChaveEnviada('', '');
    end;
    end; *)

end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  // ConexaoDB;

  // ConexaoSQLDataBase(SQLConnection1);

  VerificaInternet;

  AtualizaIni;

  ValidaINI;

  strDenegada := 'N';

  // Form1.Top := 2000;
  // Alterado
  ACBrNFe1 := TACBrNFe.Create(self);
  ACBrNFeDANFEFR1 := TACBrNFeDANFEFR.Create(self);
  ACBrNFe1.Configuracoes.Arquivos.PathSalvar :=
    ExtractFilePath(Application.ExeName);

  LerConfiguracao;

  // GravarConfiguracao;

  try
    ACAO := ParamStr(1);

    if ACAO = 'DISPONIBILIDADE' then
      bConsultaDisponibilidade;

    if ACAO <> 'DISPONIBILIDADE' then
    begin
      ParNF := strtoint(ParamStr(2));

      if ACAO = 'VALIDAR' then
        validaCliente(ParamStr(3), ParamStr(4)); // 3 - CGC, 4 - UF
    end;

    if ParNF <> 0 then
    begin
      bConfigClick;

      if ACAO = 'ENVIA' then
        bEnviarNfeClick
      else if ACAO = 'CANCELA' then
        bCancelaNFeClick
      else if ACAO = 'COMPLEMENTO' then
        bComplementoClick
      else if ACAO = 'CONSULTA' then
        bConsultaNFeClick
      else if ACAO = 'DANFE' then
        bDanfeClick
      else if ACAO = 'INUTIL' then
        bInutilizaClick
      else if ACAO = 'ENVIASN' then
        bEnviarSNClick
      else if ACAO = 'COMPLEMENTOSN' then
        bComplementoSNClick
      else if ACAO = 'DPEC' then
        ContigenciaDPEC
      else if ACAO = 'DPECSEFAZ' then
        DPEC_SEFAZ
      else if ACAO = 'XML' then
        bEnviarNfeClick2;

      Timer1.Enabled := true;
    end;
  except
    Application.Terminate;
  end;
  Application.Terminate;
end;

procedure TForm1.ACBrNFe1GerarLog(const ALogLine: string; var Tratado: Boolean);
begin

  Memo2.Lines.Text := Memo2.Lines.Text + ALogLine + #13#10 +
    StringOfChar('#', 20) + #13#10;

end;

procedure TForm1.AtualizaIni;
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

procedure TForm1.bCancelaNFeClick;
var
  Arquivo: TextFile;
  Linha, Chave, aux: string;
begin
  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\CNNOTA' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CN' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  except
    Application.Terminate;
  end;

  AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\CNNOTA' +
    IntToStr(ParNF) + '.txt');
  Reset(Arquivo);
  Readln(Arquivo, Linha);

  ACBrNFe1.NotasFiscais.Clear;
  if not fileexists(trim(Copy(Linha, 4, 400))) then
  begin
    Memo1.Lines.Clear;
    Memo1.Lines.Text := 'Arquivo não existe.2';
    Memo1.Lines.Add
      ('Arquivo do XML da NFe não foi encontrado para o cancelamento.');
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CN' +
      IntToStr(ParNF) + '.txt');
    Application.Terminate;
  end;

  ACBrNFe1.NotasFiscais.LoadFromFile(trim(Copy(Linha, 4, 400)));
  aux := Copy(Linha, (pos('NF-e- ', Linha) + length('NF-e- ')), 44);
  try

    try
      Memo2.Lines.Clear;
      ACBrNFe1.Cancelamento(trim(Copy(Linha, 405, 120)));
    finally
      Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
        IntToStr(ParNF) + '.txt');
    end;

  except
    on E: Exception do
    begin
      Memo1.Lines.Text := (E.Message);
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CN' +
        IntToStr(ParNF) + '.txt');
      if trim(E.Message) = trim('Rejeição: Cancelamento para NF-e já cancelada')
      then
        if fileexists(VCGeraisCaminhoArquivoRetorno + '\' + aux + '-nfe.xml')
        then
          MoveFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' + aux +
            '-nfe.xml'), pchar(VCGeraisCaminhoArquivoRetorno + '\' +
            IntToStr(ParNF) + ' Cancelamento - NF-e- ' + aux + '.xml'));

      Application.Terminate;
      SysUtils.Abort;
    end
  end;

  // Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Cancelamento.RetWS);
  // LoadXML(Memo1, WBResposta);
  CopyFile(pchar(VCGeraisCaminhoArquivosXML + '\' +
    ACBrNFe1.WebServices.Retorno.ChaveNFe + '-nfe.xml'),
    pchar(VCGeraisCaminhoArquivoCancelada + '\' + IntToStr(VNumLote) +
    ' Cancelamento - NF-e- ' + ACBrNFe1.WebServices.Retorno.ChaveNFe +
    '.xml'), False);
  CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' + aux + '-nfe.xml'),
    pchar(VCGeraisCaminhoArquivoCancelada + '\' + IntToStr(ParNF) +
    ' Cancelamento - NF-e- ' + aux + '.xml'), False);
  MoveFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' + aux + '-nfe.xml'),
    pchar(VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(ParNF) +
    ' Cancelamento - NF-e- ' + aux + '.xml'));

end;

procedure TForm1.bCancelaNFeEventoClick;
var
  idLote, VAux: String;
  Arquivo: TextFile;
  Linha, Chave, aux: string;
begin
  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\CNNOTA' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CN' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  except
    Application.Terminate;
  end;

  AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\CNNOTA' +
    IntToStr(ParNF) + '.txt');
  Reset(Arquivo);
  Readln(Arquivo, Linha);

  ACBrNFe1.NotasFiscais.Clear;
  if not fileexists(trim(Copy(Linha, 4, 400))) then
  begin
    Memo1.Lines.Clear;
    Memo1.Lines.Text := 'Arquivo não existe.2';
    Memo1.Lines.Add
      ('Arquivo do XML da NFe não foi encontrado para o cancelamento.');
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CN' +
      IntToStr(ParNF) + '.txt');
    Application.Terminate;
  end;

  ACBrNFe1.NotasFiscais.Clear;
  ACBrNFe1.NotasFiscais.LoadFromFile(trim(Copy(Linha, 4, 400)));

  try

    try
      Memo2.Lines.Clear;
      idLote := '1';
      VAux := '';
      ACBrNFe1.EventoNFe.Evento.Clear;
      ACBrNFe1.EventoNFe.idLote := strtoint(idLote);
      with ACBrNFe1.EventoNFe.Evento.Add do
      begin
        infEvento.dhEvento := now;
        infEvento.tpEvento := teCancelamento;
        infEvento.detEvento.xJust := VAux;
      end;

      ACBrNFe1.EnviarEvento(strtoint(idLote));

    finally
      Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
        IntToStr(ParNF) + '.txt');
    end;

  except
    on E: Exception do
    begin
      Memo1.Lines.Text := (E.Message);
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CN' +
        IntToStr(ParNF) + '.txt');
      if trim(E.Message) = trim('Rejeição: Cancelamento para NF-e já cancelada')
      then
        if fileexists(VCGeraisCaminhoArquivoRetorno + '\' + aux + '-nfe.xml')
        then
          MoveFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' + aux +
            '-nfe.xml'), pchar(VCGeraisCaminhoArquivoRetorno + '\' +
            IntToStr(ParNF) + ' Cancelamento - NF-e- ' + aux + '.xml'));

      Application.Terminate;
      SysUtils.Abort;
    end
  end;

end;

procedure TForm1.bComplementoClick;
var
  Arquivo: TextFile;
  Linha: string;
  i, j, colunai, colunaf: integer;
  achou: Boolean;
  chaveDup, protocoloDupl, aux: string;
  ArrumaDuplicada: TextFile;
begin
  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\CPNOTA' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  except
    Application.Terminate;
  end;

  // Busca arquivo de Leitura
  try
    EnviaNFe2('', '', False);
  Except
    on E: Exception do
    begin
      ShowMessage(E.Message);
    end;
  end;

  try
    Foi := true;
    try
      Memo2.Lines.Clear;
      ACBrNFe1.Enviar(VNumLote, False);
    finally
      Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
        IntToStr(VNumLote) + '.txt');
    end;
    ACBrNFe1.NotasFiscais.ImprimirPDF;
  except
    on E: Exception do
    begin

      // Validando retorno para NFe Denegada
      if (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 110) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 205) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 301) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 302) then
      begin
        strDenegada := 'S';
        ValidaRejeicao(IntToStr(ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.cStat) + ' - ' + ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.xMotivo);
      end;

      Memo1.Lines.Text := (E.Message);
      Foi := False;
      // ==============================
      for i := 0 to Memo1.Lines.Count - 1 do
      begin
        colunai := pos('[', Memo1.Lines[i]);
        colunaf := pos(']', Memo1.Lines[i]);

        if ((colunai > 0) and (colunaf > colunai)) then
        begin
          chaveDup := '';
          Linha := Memo1.Lines[i];

          for j := colunai + 1 to colunaf - 1 do
            if Linha[j] in ['0' .. '9'] then
            begin
              chaveDup := chaveDup + Linha[j];
            end;

          if length(chaveDup) <> 44 then
          begin
            achou := False;
            Foi := False;
          end
          else
          begin
            achou := true;
          end;
        end;
      end;

      if achou then
      begin
        ACBrNFe1.NotasFiscais.Clear;
        ACBrNFe1.WebServices.Consulta.NFeChave := chaveDup;
        ACBrNFe1.WebServices.Consulta.Executar;
        aux := UTF8Encode(ACBrNFe1.WebServices.Consulta.RetWS);
        colunai := pos('<protNFe', aux);
        colunaf := pos('</protNFe>', aux);
        protocoloDupl := '';
        for j := colunai to colunaf + 9 do
          protocoloDupl := protocoloDupl + aux[j];

        Clipboard.Open;
        try
          Clipboard.AsText := chaveDup;
        finally
          Clipboard.Close;
        end;

        if messagebox(Handle,
          pchar('Numeração de NFe já enviada anteriormente com a chave: ' +
          chaveDup + #13 +
          'Caso seja a mesma NFe você pode optar por Sim para recuperar o XML.'
          + #13 + 'Em caso de dúvida poderá consultar no site ' +
          'http://www.nfe.fazenda.gov.br' + #13 +
          'Para facilitar a chave se encontra na área de transferência. Basta utilizar CTRL+V para colar a chave no campo de pesquisa.'),
          pchar('Envio de NFe'), MB_YESNO + MB_ICONQUESTION) = IDYES then
        begin
          if RecuperarXMLFATPED(chaveDup, protocoloDupl, False) then
          begin
            ACBrNFe1.NotasFiscais.LoadFromFile
              (ACBrNFe1.Configuracoes.Arquivos.GetPathNFe(now) + '\' + chaveDup
              + '-nfe.xml');
            Foi := true;
          end
          else
            Foi := False;
        end
        else
          Foi := False;
      end;
    end;
  end;

  // =============================================
  // showmessage(Memo1.Lines.Text);
  if not Foi then
  begin
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
      IntToStr(VNumLote) + '.txt');
    ReescreveChaveEnviada('', '');
    Application.Terminate;
    SysUtils.Abort;
    Application.Terminate;
  end
  else
  begin
    if (ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].nProt <> '') or
      (chaveDup <> '') then
    begin
      Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Retorno.RetWS);
      /// LoadXML(Memo1, WBResposta);
      CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' +
        ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
        '-nfe.xml'), pchar(VCGeraisCaminhoArquivoRetorno + '\' +
        IntToStr(VNumLote) + ' - NF-e- ' +
        ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
        '.xml'), true);

      RecChave := RecuperaChaveEnviando;

      if RecChave <> ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe
      then
        ReescreveChaveEnviada(ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.
          Items[0].chDFe, ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items
          [0].nProt);
    end
    else
    begin
      ReescreveChaveEnviada('', '');
    end;
  end;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  Application.Terminate;
end;

function TForm1.validaCliente(cgc, uf: String): TCliente;
var
  IniFile: TIniFile;
  status: String;
begin

  cgc := LimpaStr(cgc);

  ACBrNFe1.WebServices.ConsultaCadastro.uf := uf;
  if length(cgc) > 11 then
    ACBrNFe1.WebServices.ConsultaCadastro.CNPJ := cgc
  else
    ACBrNFe1.WebServices.ConsultaCadastro.cpf := cgc;
  ACBrNFe1.WebServices.ConsultaCadastro.Executar;

  Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.ConsultaCadastro.RetWS);
  Memo2.Lines.Text :=
    UTF8Encode(ACBrNFe1.WebServices.ConsultaCadastro.RetornoWS);
  /// LoadXML(Memo1, WBResposta);

  if (ACBrNFe1.WebServices.ConsultaCadastro.RetConsCad.InfCad.Items[0].cSit = 0)
  then
  begin
    Result.isDenegada := true;
    status := 'Inapto à emissão de NFe';
  end
  else
  begin
    Result.isDenegada := False;
    status := 'Apto à emissão de NFe';
  end;

  Result.ie := ACBrNFe1.WebServices.ConsultaCadastro.RetConsCad.InfCad.
    Items[0].ie;
  Result.cgc := ACBrNFe1.WebServices.ConsultaCadastro.RetConsCad.InfCad.
    Items[0].CNPJ;

  // Procedure para exibir os dados do cliente em um formulário
  consultaCliente(Result.cgc, Result.ie, status);

  { //Escrevendo informações no config.ini
    iniFile := TIniFile.Create(ExtractFilePath(application.ExeName) + 'config.ini');
    iniFile.WriteString('validacao', 'ie', result.ie);
    iniFile.WriteString('validacao', 'cgc', result.cgc);
    iniFile.WriteBool('validacao', 'denegada', result.isDenegada); }

end;

procedure TForm1.bConsultaDisponibilidade;
var
  strAux: string;
begin
  ACBrNFe1.WebServices.StatusServico.Executar;

  // MemoResp.Lines.Text := ACBrNFe1.WebServices.StatusServico.RetWS;
  // Memo1.Lines.Text := ACBrNFe1.WebServices.StatusServico.RetornoWS;
  // LoadXML(ACBrNFe1.WebServices.StatusServico.RetornoWS, Memo1);

  strAux := 'Status Serviço';
  strAux := strAux + #10#13 + 'tpAmb: ' +
    TpAmbToStr(ACBrNFe1.WebServices.StatusServico.tpAmb);
  strAux := strAux + #10#13 + 'verAplic: ' +
    ACBrNFe1.WebServices.StatusServico.verAplic;
  strAux := strAux + #10#13 + 'cStat: ' +
    IntToStr(ACBrNFe1.WebServices.StatusServico.cStat);
  strAux := strAux + #10#13 + 'xMotivo: ' +
    ACBrNFe1.WebServices.StatusServico.xMotivo;
  strAux := strAux + #10#13 + 'cUF: ' +
    IntToStr(ACBrNFe1.WebServices.StatusServico.cUF);
  strAux := strAux + #10#13 + 'dhRecbto: ' +
    DateTimeToStr(ACBrNFe1.WebServices.StatusServico.dhRecbto);
  strAux := strAux + #10#13 + 'tMed: ' +
    IntToStr(ACBrNFe1.WebServices.StatusServico.TMed);
  strAux := strAux + #10#13 + 'dhRetorno: ' +
    DateTimeToStr(ACBrNFe1.WebServices.StatusServico.dhRetorno);
  strAux := strAux + #10#13 + 'xObs: ' +
    ACBrNFe1.WebServices.StatusServico.xObs;
end;

procedure TForm1.bConsultaNFeClick;
var
  Arquivo: TextFile;
  Linha, aux: string;
begin
  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\CSNOTA' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CS' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  except
    Application.Terminate;
  end;

  AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\CSNOTA' +
    IntToStr(ParNF) + '.txt');
  Reset(Arquivo);
  Readln(Arquivo, Linha);

  ACBrNFe1.NotasFiscais.Clear;
  if not fileexists(trim(Copy(Linha, 1, 900))) then
  begin
    Memo1.Lines.Clear;
    Memo1.Lines.Text := 'Arquivo não existe.';
    Memo1.Lines.Add
      ('Arquivo do XML da NFe não foi encontrado para a Consulta ao SEFAZ.');
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CS' +
      IntToStr(ParNF) + '.txt');
    Application.Terminate;
  end;

  ACBrNFe1.NotasFiscais.Clear;
  ACBrNFe1.NotasFiscais.LoadFromFile(trim(Copy(Linha, 4, 890)));

  { ACBrNFe1.WebServices.Consulta.NFeChave := trim(Copy(Linha, 4, 400));
    ACBrNFe1.WebServices.Consulta.Executar;

    Memo1.Lines.Text :=  UTF8Encode(ACBrNFe1.WebServices.Consulta.RetWS);
    //memoRespWS.Lines.Text :=  UTF8Encode(ACBrNFe1.WebServices.Consulta.RetornoWS);
    LoadXML(Memo1, WBResposta); }

  try
    ACBrNFe1.Consultar;
  except
    on E: Exception do
    begin
      Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Consulta.RetWS);
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-CS' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
      SysUtils.Abort;
    end
  end;

  Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Consulta.RetWS);
  Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\' + IntToStr(ParNF) +
    ' Consulta - NF-e- ' + IntToStr(ParNF) + '.xml');
  Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(ParNF) +
    ' Consulta - NF-e- ' + IntToStr(ParNF) + '.TXT');

end;

procedure TForm1.bDanfeClick;
var
  Arquivo: TextFile;
  Linha: string;
begin
  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\DANFE' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;

      Memo1.Lines.Text := 'Arquivo de Localização de NFE NÃO encontrado';

      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
        IntToStr(ParNF) + '.txt');
      messagebox(0, pchar('Arquivo não localizado: ' +
        VCGeraisCaminhoArquivoLeitura + '\DANFE' + IntToStr(ParNF) + '.txt' +
        '. Favor verifique se a pasta existe e tente novamente.'),
        'Reimpressão do Danfe', MB_OK + MB_ICONEXCLAMATION);

      Application.Terminate;

    end;
  except
    on E: Exception do
    begin
      messagebox(0, pchar('Erro ao verificar os arquivos: ' + E.Message),
        'Reimpressão do Danfe', MB_OK + MB_ICONEXCLAMATION);
      Application.Terminate;
    end;
  end;

  AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\DANFE' + IntToStr(ParNF)
    + '.txt');
  Reset(Arquivo);

  Readln(Arquivo, Linha);
  // Linha := 'x:\XML\' + Copy(Linha, (pos('NF-e- ', Linha) + length('NF-e- ')), 44) + '-nfe.xml';
  try
    if fileexists(Linha) then
    begin
      ACBrNFe1.NotasFiscais.LoadFromFile(Linha);

      { if ACBrNFe1.NotasFiscais.Items[0].NFe.Ide.tpEmis = teDPEC then
        begin

        Readln(Arquivo, Linha);

        if trim(Linha) = '' then
        begin
        ACBrNFe1.WebServices.ConsultaDPEC.NFeChave := ACBrNFe1.NotasFiscais.Items[0].NFe.infNFe.ID;
        ACBrNFe1.WebServices.ConsultaDPEC.Executar;
        Linha := ACBrNFe1.WebServices.ConsultaDPEC.nRegDPEC + ' ' + DateTimeToStr
        (ACBrNFe1.WebServices.ConsultaDPEC.dhRegDPEC);

        end;
        ACBrNFe1.DANFE.ProtocoloNFe := Linha;
        end; }

      ACBrNFeDANFEFR1.MostraSetup := true;
      ACBrNFe1.NotasFiscais.Imprimir;
      ACBrNFe1.NotasFiscais.ImprimirPDF;

    end
    else
    begin
      messagebox(0, pchar('XML não encontrado em: ' + Linha +
        '. Favor verifique.'), 'Reimpressão da Danfe', MB_OK + MB_ICONWARNING);
    end;
  except
    on E: Exception do
    begin
      messagebox(0, pchar('Erro na Geração da Danfe : ' + E.Message),
        'Reimpressão do Danfe', MB_OK + MB_ICONEXCLAMATION);
    end;

  end;
  CloseFile(Arquivo);
end;

procedure TForm1.bInutilizaClick;
var
  Arquivo: TextFile;
  Linha: string;
  CNPJ, justificativa: string;
  numeroinicial, numerofinal, ano, modelo, serie: integer;
begin
  AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\INUTILIZAR' +
    IntToStr(ParNF) + '.txt');
  Reset(Arquivo);
  Readln(Arquivo, Linha);
  CNPJ := Linha;

  Readln(Arquivo, Linha);
  justificativa := Linha;

  Readln(Arquivo, Linha);
  ano := strtoint(Linha);

  Readln(Arquivo, Linha);
  modelo := strtoint(Linha);

  Readln(Arquivo, Linha);
  serie := strtoint(Linha);

  Readln(Arquivo, Linha);
  numeroinicial := strtoint(Linha);
  numerofinal := strtoint(Linha);
  CloseFile(Arquivo);
  try
    try
      Memo2.Lines.Clear;
      ACBrNFe1.WebServices.Inutiliza(CNPJ, justificativa, ano, modelo, serie,
        numeroinicial, numerofinal);
    finally
      Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
        IntToStr(numeroinicial) + '.txt');
    end;

  except
    on E: Exception do
    begin
      Memo1.Lines.Text := (E.Message);
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-IN' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
      SysUtils.Abort;
    end
  end;
  Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Inutilizacao.RetWS);
  // Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\' + IntToStr(ParNF) + ' INUTILIZADA - NF-e.xml');
  Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(ParNF) +
    ' INUTILIZADA - NF-e.xml');
end;

procedure TForm1.bEnviarSNClick;
var
  Arquivo: TextFile;
  Linha: string;
  i, j, colunai, colunaf: integer;
  achou: Boolean;
  chaveDup, protocoloDupl, aux: string;
  ArrumaDuplicada: TextFile;
  Cdi: integer;
  VSNprod, VSNAliq: Real;
begin

  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\EVNOTA' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-ERRO' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  except
    Application.Terminate;
  end;
  AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\EVNOTA' +
    IntToStr(ParNF) + '.txt');
  Reset(Arquivo);
  Readln(Arquivo, Linha);

  // Busca informações do Arquivo de Leitura
  try
    EnviaNFe2('', '', true);
  Except
    on E: Exception do
    begin
      ShowMessage(E.Message);
    end;
  end;

  try
    Foi := true;

    try
      Memo2.Lines.Clear;
      ACBrNFe1.Enviar(VNumLote, False);
    finally
      Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
        IntToStr(VNumLote) + '.txt');
    end;

    ACBrNFe1.NotasFiscais.ImprimirPDF;
  except
    on E: Exception do
    begin

      // Validando retorno para NFe Denegada
      if (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 110) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 205) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 301) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 302) then
      begin
        strDenegada := 'S';
        ValidaRejeicao(IntToStr(ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.cStat) + ' - ' + ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.xMotivo);
      end;

      Memo1.Lines.Text := (E.Message);
      Foi := False;
      // ==============================
      for i := 0 to Memo1.Lines.Count - 1 do
      begin
        colunai := pos('[', Memo1.Lines[i]);
        colunaf := pos(']', Memo1.Lines[i]);
        if ((colunai > 0) and (colunaf > colunai)) then
        begin
          chaveDup := '';
          Linha := Memo1.Lines[i];
          for j := colunai + 1 to colunaf - 1 do
            if Linha[j] in ['0' .. '9'] then
              chaveDup := chaveDup + Linha[j];
          achou := true;
        end;
      end;

      if achou then
        messagebox(Handle, pchar(E.Message), 'Envio NFe', MB_OK);

      if achou then
      begin
        ACBrNFe1.NotasFiscais.Clear;
        ACBrNFe1.WebServices.Consulta.NFeChave := chaveDup;
        ACBrNFe1.WebServices.Consulta.Executar;
        aux := UTF8Encode(ACBrNFe1.WebServices.Consulta.RetWS);
        colunai := pos('<protNFe', aux);
        colunaf := pos('</protNFe>', aux);
        protocoloDupl := '';
        for j := colunai to colunaf + 9 do
          protocoloDupl := protocoloDupl + aux[j];

        Clipboard.Open;
        try
          Clipboard.AsText := chaveDup;
        finally
          Clipboard.Close;
        end;

        if messagebox(Handle,
          pchar('Numeração de NFe já enviada anteriormente com a chave: ' +
          chaveDup + #13 +
          'Caso seja a mesma NFe você pode optar por Sim para recuperar o XML.'
          + #13 + 'Em caso de dúvida poderá consultar no site ' +
          'http://www.nfe.fazenda.gov.br' + #13 +
          'Para facilitar a chave se encontra na área de transferência. Basta utilizar CTRL+V para colar a chave no campo de pesquisa.'),
          pchar('Envio de NFe'), MB_YESNO + MB_ICONQUESTION) = IDYES then
        begin

          if RecuperarXMLFATPED(chaveDup, protocoloDupl, true) then
          begin
            ACBrNFe1.NotasFiscais.LoadFromFile
              (ACBrNFe1.Configuracoes.Arquivos.GetPathNFe(now) + '\' + chaveDup
              + '-nfe.xml');
            Foi := true;
          end
          else
            Foi := False;
        end
        else
          Foi := False;
      end;
    end;
  end;

  // =============================================

  if not Foi then
  begin
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
      IntToStr(VNumLote) + '.txt');
    ReescreveChaveEnviada('', '');
    Application.Terminate;
    SysUtils.Abort;
    Application.Terminate;

  end
  else
  begin
    if (ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].nProt <> '') or
      (chaveDup <> '') then
    begin
      Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Retorno.RetWS);
      /// LoadXML(Memo1, WBResposta);
      CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' +
        ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
        '-nfe.xml'), pchar(VCGeraisCaminhoArquivoRetorno + '\' +
        IntToStr(VNumLote) + ' - NF-e- ' +
        ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
        '.xml'), true);

      RecChave := RecuperaChaveEnviando;
      if RecChave <> ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe
      then
        ReescreveChaveEnviada(ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.
          Items[0].chDFe, ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items
          [0].nProt);
    end
    else
    begin
      ReescreveChaveEnviada('', '');
    end;
  end;

end;

procedure TForm1.bComplementoSNClick;
var
  Arquivo: TextFile;
  Linha: string;
  i, j, colunai, colunaf: integer;
  achou: Boolean;
  chaveDup, protocoloDupl, aux: string;
  ArrumaDuplicada: TextFile;
  VSNprod, VSNAliq: Real;
begin
  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\CPNOTA' + IntToStr(ParNF) +
      '.txt') = False then
    begin
      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-ERRO-CP'
        + IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  except
    Application.Terminate;
  end;
  AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\CPNOTA' +
    IntToStr(ParNF) + '.txt');
  Reset(Arquivo);
  Readln(Arquivo, Linha);

  // Busca informações do Arquivo de Leitura
  try
    EnviaNFe2('', '', true);
  Except
    on E: Exception do
    begin
      ShowMessage(E.Message);
    end;
  end;

  try
    Foi := true;

    try
      Memo2.Lines.Clear;
      ACBrNFe1.Enviar(VNumLote, False);
    finally
      Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
        IntToStr(VNumLote) + '.txt');
    end;

    ACBrNFe1.NotasFiscais.ImprimirPDF;
  except
    on E: Exception do
    begin

      // Validando retorno para NFe Denegada
      if (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 110) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 205) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 301) or
        (ACBrNFe1.NotasFiscais.Items[0].NFe.procNFe.cStat = 302) then
      begin
        strDenegada := 'S';
        ValidaRejeicao(IntToStr(ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.cStat) + ' - ' + ACBrNFe1.NotasFiscais.Items[0]
          .NFe.procNFe.xMotivo);
      end;

      Memo1.Lines.Text := (E.Message);
      Foi := False;
      // ==============================
      for i := 0 to Memo1.Lines.Count - 1 do
      begin
        colunai := pos('[', Memo1.Lines[i]);
        colunaf := pos(']', Memo1.Lines[i]);
        if ((colunai > 0) and (colunaf > colunai)) then
        begin
          chaveDup := '';
          Linha := Memo1.Lines[i];
          for j := colunai + 1 to colunaf - 1 do
            chaveDup := chaveDup + Linha[j];
          achou := true;
        end;
      end;

      if achou then
        messagebox(Handle, pchar(E.Message), 'Envio NFe', MB_OK);

      if achou then
      begin
        ACBrNFe1.NotasFiscais.Clear;
        ACBrNFe1.WebServices.Consulta.NFeChave := chaveDup;
        ACBrNFe1.WebServices.Consulta.Executar;
        aux := UTF8Encode(ACBrNFe1.WebServices.Consulta.RetWS);
        colunai := pos('<protNFe', aux);
        colunaf := pos('</protNFe>', aux);
        protocoloDupl := '';
        for j := colunai to colunaf + 9 do
          protocoloDupl := protocoloDupl + aux[j];

        if RecuperarXMLFATPED(chaveDup, protocoloDupl, true) then
        begin
          ACBrNFe1.NotasFiscais.LoadFromFile
            (ACBrNFe1.Configuracoes.Arquivos.GetPathNFe(now) + '\' + chaveDup +
            '-nfe.xml');
          Foi := true;
        end
        else
          Foi := False;
      end;
    end;
  end;

  // =============================================

  if not Foi then
  begin
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
      IntToStr(VNumLote) + '.txt');
    ReescreveChaveEnviada('', '');

    Application.Terminate;
    SysUtils.Abort;
    Application.Terminate;
  end
  else
  begin
    if (ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].nProt <> '') or
      (chaveDup <> '') then
    begin
      Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Retorno.RetWS);
      /// LoadXML(Memo1, WBResposta);
      CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' +
        ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
        '-nfe.xml'), pchar(VCGeraisCaminhoArquivoRetorno + '\' +
        IntToStr(VNumLote) + ' - NF-e- ' +
        ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe +
        '.xml'), true);

      RecChave := RecuperaChaveEnviando;
      if RecChave <> ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items[0].chDFe
      then
        ReescreveChaveEnviada(ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.
          Items[0].chDFe, ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.Items
          [0].nProt);
    end
    else
    begin
      ReescreveChaveEnviada('', '');
    end;
  end;

end;

procedure TForm1.ConfirmaEnvioDPEC(strEnviado: string);
var
  arq: TIniFile;
  AuxStr: string;
begin
  AuxStr := ExtractFilePath(Application.ExeName) + 'Config.ini';
  arq := TIniFile.Create(AuxStr);
  try
    arq.WriteString('DPEC', 'CONFIRMA_ENV', strEnviado);

  finally
    FreeAndnil(arq);
  end;

end;

procedure TForm1.ContigenciaDPEC;
var
  Arquivo: TextFile;
  Linha: string;
  i, j, colunai, colunaf: integer;
  achou: Boolean;
  chaveDup, protocoloDupl, aux: string;
  ArrumaDuplicada: TextFile;
  Cdi: integer;
  intVezes: integer;
begin
  { achou := False;
    chaveDup := '';

    try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\DPEC' + IntToStr(ParNF) + '.txt') = False then
    begin
    Memo1.Lines.Clear;
    Memo1.Lines.Text := 'Arquivo de Envio de dados NÃO encontrado';
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-ERRO' + IntToStr(ParNF) + '.txt');
    Application.Terminate;
    end;
    except
    on E: Exception do
    begin
    Memo1.Lines.Clear;
    Memo1.Lines.Text := 'Validando Caminho de Leitura: ' + E.Message;
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-ERRO' + IntToStr(ParNF) + '.txt');
    Application.Terminate;
    end;
    end;

    // Busca informações do Arquivo de Leitura
    try
    EnviaNFe2('', '', False, 4);
    Except
    on E: Exception do
    begin
    ShowMessage(E.Message);
    end;
    end;

    try
    Foi := true;
    try
    Memo2.Lines.Clear;
    ACBrNFe1.Enviar(VNumLote, False);

    if ACBrNFe1.WebServices.EnviarDPEC.Executar then
    begin
    // protocolo de envio ao DPEC e impressão do DANFE
    ACBrNFe1.DANFE.ProtocoloNFe := ACBrNFe1.WebServices.EnviarDPEC.nRegDPEC + ' ' + DateTimeToStr
    (ACBrNFe1.WebServices.EnviarDPEC.dhRegDPEC);
    GravaProtcoloDpec(ACBrNFe1.DANFE.ProtocoloNFe);
    // ACBrNFe1.NotasFiscais.Imprimir;
    end;

    finally
    Memo2.Lines.Add('DPEC');
    Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' + IntToStr(VNumLote) + '.txt');
    end;

    ACBrNFe1.NotasFiscais.ImprimirPDF;

    CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0]
    .chNFe + '-nfe.xml'),
    pchar(VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(VNumLote)
    + ' - NF-e- ' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe + '.xml'), true);

    except
    on E: Exception do
    begin
    Foi := False;
    if pos('Duplicidade', E.Message) > 0 then
    begin
    aux := E.Message;
    chaveDup := Copy(aux, pos('[chNFe:', aux) + 7, 44);
    achou := true;
    end
    else
    begin
    Memo1.Lines.Clear;
    Memo1.Lines.Text := 'Envio para DPEC: ' + E.Message;
    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' + IntToStr(ParNF) + '.txt');
    Application.Terminate;
    end;
    end;
    end;
    ReescreveChaveEnviada(ifthen(achou, chaveDup, ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe),
    ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].nProt);
    // =============================================

    if not Foi then
    begin

    Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' + IntToStr(VNumLote) + '.txt');
    // ReescreveChaveEnviada('', '');

    Application.Terminate;
    SysUtils.Abort;
    Application.Terminate;

    end
    else
    begin

    if (ACBrNFe1.WebServices.ConsultaDPEC.nRegDPEC <> '') or (chaveDup <> '') then
    begin

    Memo1.Lines.Text := UTF8Encode(ACBrNFe1.WebServices.Retorno.RetWS);

    /// LoadXML(Memo1, WBResposta);

    intVezes := 0;

    CopyFile(pchar(VCGeraisCaminhoArquivoRetorno + '\' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0]
    .chNFe + '-nfe.xml'),
    pchar(VCGeraisCaminhoArquivoRetorno + '\' + IntToStr(VNumLote)
    + ' - NF-e- ' + ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe + '.xml'), true);

    RecChave := RecuperaChaveEnviando;

    if RecChave <> ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe then
    ReescreveChaveEnviada(ifthen(achou, chaveDup, ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].chNFe),
    ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtNFe.Items[0].nProt);
    end
    else
    begin
    // ReescreveChaveEnviada('', '');
    end;
    end; }
end;

procedure TForm1.DPEC_SEFAZ;
var
  Arquivo: TextFile;
  Linha: string;
begin
  try
    if fileexists(VCGeraisCaminhoArquivoLeitura + '\DEPCSEFAZ' + IntToStr(ParNF)
      + '.txt') = False then
    begin
      Memo1.Lines.Clear;

      Memo1.Lines.Text := 'Arquivo de Localização de DPEC NÃO encontrado';

      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
        IntToStr(ParNF) + '.txt');
      messagebox(0, pchar('Arquivo não localizado: ' +
        VCGeraisCaminhoArquivoLeitura + '\DEPCSEFAZ' + IntToStr(ParNF) + '.txt'
        + '. Favor verifique se a pasta existe e tente novamente.'),
        'Envio DPEC->SeFaz.', MB_OK + MB_ICONEXCLAMATION);

      Application.Terminate;

    end;
  except
    on E: Exception do
    begin
      messagebox(0, pchar('Erro ao verificar os arquivos: ' + E.Message),
        'Reimpressão do Danfe', MB_OK + MB_ICONEXCLAMATION);

      Memo1.Lines.Clear;
      Memo1.Lines.Text := 'Envio DPEC->SeFaz : ' + E.Message;
      Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
        IntToStr(ParNF) + '.txt');
      Application.Terminate;
    end;
  end;

  AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\DEPCSEFAZ' +
    IntToStr(ParNF) + '.txt');
  Reset(Arquivo);
  // x:\XML\34619 - NF-e- 35150361748646000146550010000346191508302463-nfe.xml
  Readln(Arquivo, Linha);

  // Linha := 'x:\XML\' + Copy(Linha, (pos('NF-e- ', Linha) + length('NF-e- ')), 44)  + '-nfe.xml';
  try
    try
      if fileexists(Linha) then
      begin

        try

          ACBrNFe1.NotasFiscais.Clear;
          ACBrNFe1.NotasFiscais.LoadFromFile(Linha);
          ACBrNFe1.Enviar(IntToStr(ParNF), False);

          ReescreveChaveEnviada(ACBrNFe1.WebServices.Retorno.NFeRetorno.ProtDFe.
            Items[0].chDFe, '');

          ConfirmaEnvioDPEC('S');

        except
          on E: Exception do
          begin
            if pos('Duplicidade', E.Message) > 0 then
              ConfirmaEnvioDPEC('S')
            else
              ConfirmaEnvioDPEC('N');
          end;
        end;

        // ACBrNFe1.NotasFiscais.Imprimir;
      end
      else
      begin
        messagebox(0, pchar('XML não encontrado em: ' + Linha +
          '. Favor verifique.'), 'Reimpressão da Danfe',
          MB_OK + MB_ICONWARNING);
      end;
    except
      on E: Exception do
      begin
        messagebox(0, pchar('Erro na Geração da Danfe : ' + E.Message),
          'Reimpressão do Danfe', MB_OK + MB_ICONEXCLAMATION);
      end;

    end;
  finally
    Memo2.Lines.SaveToFile(VCGeraisCaminhoArquivosXML + '\LogGeral-' +
      IntToStr(ParNF) + '.txt');
  end;

  CloseFile(Arquivo);

end;

function TForm1.EnviaNFe2(RChave, RProtocolo: String; SN: Boolean;
  TipoEnvio: integer = 3): Boolean;
var
  Arquivo: TextFile;
  Linha: string;
  i, j, colunai, colunaf: integer;
  Cdi: integer;
  VSNprod, VSNAliq, vliq, voriginal, Pfcp, Vfcp: Real;
  codigoBarra, unidadeTributacao, numeroparcela, FCI, nfat, loja: string;
begin
  Result := False;
  Vfcp := 0;

  // Carrega Arquivo para leitura
  if ACAO = 'COMPLEMENTO' then
  begin
    AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\CPNOTA' +
      IntToStr(ParNF) + '.txt');
  end
  else if ACAO = 'COMPLEMENTOSN' then
  begin
    AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\CPNOTA' +
      IntToStr(ParNF) + '.txt');
  end
  else if ACAO = 'ENVIA' then
  begin
    AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\EVNOTA' +
      IntToStr(ParNF) + '.txt');
  end
  else if ACAO = 'ENVIASN' then
  begin
    AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\EVNOTA' +
      IntToStr(ParNF) + '.txt');
  end
  else if ACAO = 'DPEC' then
  begin
    AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\DPEC' +
      IntToStr(ParNF) + '.txt');
  end
  else if ACAO = 'XML' then
  begin
    AssignFile(Arquivo, VCGeraisCaminhoArquivoLeitura + '\EVNOTA' +
      IntToStr(ParNF) + '.txt');
  end;

  // Vai para Primeira Linha
  Reset(Arquivo);

  // Lê a Linha
  Readln(Arquivo, Linha);

  with ACBrNFe1.NotasFiscais.Add.NFe do
  begin
    // Loop para ler todas as linhas do arquivo
    while not Eof(Arquivo) do
    begin
      // =============================EM0201=============================//
      if (Copy(Linha, 0, 6)) = 'EM0201' then
      begin
        // Carrega Chave
        // infNFe.ID:= 'NFe'+Copy(Linha,9,44);
        if RChave <> '' then
          infNFe.ID := 'NFe' + RChave;

        Chave := Copy(Linha, 9, 44);
        vchave := Copy(Linha, 9, 44);

      end
      else
        // fim do EM0201
        // =============================EM1201=============================//
        if (Copy(Linha, 0, 6)) = 'EM1201' then // Dados de contigencia
        begin
          Ide.dhCont := now;
          Ide.xJust := Copy(Linha, 26, 255);
        end
        // fim do EM0201
        // =============================EM0202=============================//
        else if (Copy(Linha, 0, 6)) = 'EM0202' then
        begin
          // forma de Pagamento
          Ide.cUF := strtoint(Copy(Linha, 7, 2));
          // Código da UF do emitente do documento fiscal
          Ide.natOp := trim(Copy(Linha, 18, 60));
          // Descrição da natureza de operação
          // Indicador da forma de pagamento 0-Pagamento à vista 1-Pagamento à prazo 2-Outros showmessage(copy(Linha,78,1));
          if Copy(Linha, 78, 1) = '0' then
            Ide.indPag := ipVista
          else if Copy(Linha, 78, 1) = '1' then
            Ide.indPag := ipPrazo
          else
            Ide.indPag := ipOutras;
          Ide.modelo := strtoint(Copy(Linha, 79, 2));
          // Código do Modelo do documento fiscal
          Ide.serie := strtoint(Copy(Linha, 81, 1));
          // Série do documento fiscal
          Ide.nNF := strtoint(Copy(Linha, 82, 9));
          // Número do documento fiscal
          VNumLote := strtoint(Copy(Linha, 82, 9)); // numero do lote de envio
          try
            // Ide.dEmi := strtodate(Copy(Linha, 99, 2) + '/' + Copy(Linha, 96, 2) + '/' + Copy(Linha, 91, 4));
            Ide.dEmi := now;
            // Data de emissão do documento fiscal
          except
            Memo1.Lines.Text := 'Problemas com a Data de Emissão da nota';
            Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
              IntToStr(VNumLote) + '.txt');
            Application.Terminate;
          end;
          // Data de saida ou entrada da Mercadoria/Produto
          if (Copy(Linha, 109, 2) + '/' + Copy(Linha, 106, 2) + '/' +
            Copy(Linha, 101, 4)) <> '00/00/0000' then
            Ide.dSaiEnt :=
              strtodate((Copy(Linha, 109, 2) + '/' + Copy(Linha, 106, 2) + '/' +
              Copy(Linha, 101, 4)));
          // Tipo do documento fiscal
          if Copy(Linha, 111, 1) = '1' then
            Ide.tpNF := tnSaida
          else
            Ide.tpNF := tnEntrada;
          Ide.cMunFG := strtoint(Copy(Linha, 112, 7));
          // Código do Municipio de Ocorrência do Fato Gerador
          // Formato de Impressao do DANFE
          if (Copy(Linha, 119, 1)) = '1' then
            Ide.tpImp := tiRetrato
          else
            Ide.tpImp := tiPaisagem;

          if TipoEnvio = 3 then
            Ide.tpEmis := teNormal
          else if TipoEnvio = 4 then
            Ide.tpEmis := teDPEC;

          // Forma de emissão da NF-e
          Ide.cDV := strtoint(Copy(Linha, 121, 1));
          // Digito verificador da Chave de Acesso da NF-e

          // Identificação do Ambiente
          if (VWebSAmbiente) = 1 then
            Ide.tpAmb := taHomologacao
          else
            Ide.tpAmb := taProducao;

          // Finalidade de emissão da NF-e
          { if (Copy(Linha, 123, 1)) = '1' then
            ide.finNFe := fnNormal; }

          // Finalidade de emissão da NF-e
          case (strtoint(Copy(Linha, 123, 1))) of
            1:
              Ide.finNFe := fnNormal;
            2:
              Ide.finNFe := fnComplementar;
            3:
              Ide.finNFe := fnAjuste;
            4:
              Ide.finNFe := fnDevolucao;

          end;
          // Forma de Pagamento Ajuste e Devolução
          // if (Ide.finNFe = fnAjuste) or (Ide.finNFe = fnDevolucao) or
          // (Ide.finNFe = fnComplementar) then
          if (Ide.finNFe <> fnNormal) then
          begin
            with pag.Add do
            begin
              tPag := fpSemPagamento;
              vPag := 0;
            end;
          end;

          { if (Copy(Linha, 123, 1)) = '1' then
            begin
            Ide.finNFe := fnNormal;
            end
            else
            if (Copy(Linha, 123, 1)) = '2' then
            begin
            Ide.finNFe := fnComplementar;

            end; }

          Ide.procEmi := peAplicativoContribuinte;

          // Processo de emissão da NF-e
          Ide.verProc := Copy(Linha, 125, 20);

          case (strtoint(Copy(Linha, 145, 2))) of
            1:
              Ide.idDest := doInterna;
            2:
              Ide.idDest := doInterestadual;
            3:
              Ide.idDest := doExterior;
          end;

          case (strtoint(Copy(Linha, 147, 2))) of
            0:
              Ide.indFinal := cfNao;
            1:
              Ide.indFinal := cfConsumidorFinal;

          end;

          case (strtoint(Copy(Linha, 149, 2))) of
            0:
              Ide.indPres := pcNao;
            1:
              Ide.indPres := pcPresencial;
            2:
              Ide.indPres := pcInternet;
            3:
              Ide.indPres := pcTeleatendimento;
            4:
              Ide.indPres := pcEntregaDomicilio;
            9:
              Ide.indPres := pcOutros;

          end;

        end
        else if (Copy(Linha, 0, 6)) = 'EM1202' then
        begin

          // Versão do processo de emissão da NF-e
          with Ide.NFref.Add do
          begin
            if Copy(Linha, 7, 3) <> '000' then
              refNFe := Copy(Linha, 7, 44);

            { if Copy(Linha, 51, 2)  <> '00' then
              RefNF.cUF := strtoint(Copy(Linha, 51, 2));

              if trim(Copy(Linha, 53, 2)) <> '' then
              RefNF.modelo := strtoint(Copy(Linha, 53, 2));

              RefNF.AAMM := Copy(Linha, 55, 4);

              if trim(Copy(Linha, 59, 3)) <> '' then
              RefNF.serie := strtoint(Copy(Linha, 59, 3)); RefNF.CNPJ := Copy(Linha, 62, 14);

              if trim(Copy(Linha, 76, 9)) <> '' then
              RefNF.nNF := strtoint(Copy(Linha, 76, 9)); }

            if trim(Copy(Linha, 87, 3)) <> '' then
            begin
              RefECF.modelo := ECFModRef2D;
              RefECF.nECF := Copy(Linha, 87, 3);
            end;

            if trim(Copy(Linha, 90, 6)) <> '' then
              RefECF.nCOO := Copy(Linha, 90, 6);

            // showmessage(copy(Linha, 177, 44));
          end;
        end
        else
          // =============================EM0203=============================//
          if (Copy(Linha, 0, 6)) = 'EM0203' then
          begin

            // Razão social obrigatória para ambiente de homologação
            if VWebSAmbiente = 1 then
              strxNome :=
                'NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL'
            else
              strxNome := Copy(Linha, 35, 60);

            // CNPJ do emitente  ou  CPF do emitente
            if ValidaCNPJ(Copy(Linha, 7, 14), False) = true then
              emit.CNPJCPF := trim((Copy(Linha, 7, 14)))
            else
              emit.CNPJCPF := trim((Copy(Linha, 21, 14)));
            emit.xNome := trim(strxNome);
            // copy(Linha, 35, 60); // Razão social ou Nome do emitente
            emit.xFant := trim(Copy(Linha, 95, 60)); // Nome fantasia
            emit.EnderEmit.xLgr := trim(Copy(Linha, 155, 60)); // Logradouro
            emit.EnderEmit.nro := trim(Copy(Linha, 215, 60)); // Número
            emit.EnderEmit.xCpl := trim(Copy(Linha, 275, 60)); // Complemento
            emit.EnderEmit.xBairro := trim(Copy(Linha, 335, 60)); // Bairro
            emit.EnderEmit.cMun := strtoint(Copy(Linha, 395, 7));
            // Código do municipio
            emit.EnderEmit.xMun := trim(Copy(Linha, 402, 60));
            // Nome do municipio
            emit.EnderEmit.uf := (Copy(Linha, 462, 2)); // Sigla da UF
            ACBrNFe1.Configuracoes.WebServices.uf := (Copy(Linha, 462, 2));
            emit.EnderEmit.CEP := strtoint(Copy(Linha, 464, 8));
            // Código do CEP
            emit.EnderEmit.cPais := strtoint(Copy(Linha, 472, 4));
            // Código do País
            emit.EnderEmit.xPais := trim(Copy(Linha, 476, 60)); // Brasil
            emit.EnderEmit.fone := trim(Copy(Linha, 536, 10)); // Telefone
            emit.ie := trim(Copy(Linha, 546, 14)); // IE
            emit.IEST := trim(Copy(Linha, 560, 18)); // IEST

            if SN then
              emit.CRT := crtSimplesNacional;

            // =========================Responsável Técnico=======================//
            if (emit.EnderEmit.uf = 'SC') then
            begin
              infRespTec.CNPJ := '05557708000161';
              infRespTec.xContato := 'CARLOS ALBERTO FERREIRA DOS SANTOS';
              infRespTec.email := 'carlos@emerion.com.br';
              infRespTec.fone := '01120217221';
            end;
            // ====================================================================//
          end
          else // fim do EM0203
            // =============================EM0204=============================//
            if (Copy(Linha, 0, 6)) = 'EM0204' then
            begin
              if Copy(Linha, 402, 2) <> 'EX' then
              begin
                // CNPJ do emitente  ou  CPF do emitente
                if ((ValidaCNPJ(Copy(Linha, 7, 14), False) = true) and
                  (Copy(Linha, 7, 14) <> '              ')) then
                begin
                  dest.CNPJCPF := (Copy(Linha, 7, 14));
                end
                else
                begin
                  dest.CNPJCPF := (Copy(Linha, 21, 11));
                end;
              end
              else
                dest.idEstrangeiro := Copy(Linha, 512, 20);

              dest.xNome := trim(Copy(Linha, 35, 60));
              // Razão social ou nome do destinatario
              dest.EnderDest.xLgr := trim(Copy(Linha, 95, 60)); // Logradouro
              dest.EnderDest.nro := trim(Copy(Linha, 155, 60)); // Número
              dest.EnderDest.xCpl := trim(Copy(Linha, 215, 60)); // Complemento
              dest.EnderDest.xBairro := trim(Copy(Linha, 275, 60)); // Bairro
              dest.EnderDest.cMun := strtoint(Copy(Linha, 335, 7));
              // Código do Municipio
              dest.EnderDest.xMun := trim(Copy(Linha, 342, 60));
              // Nome do Municipio
              dest.EnderDest.uf := (Copy(Linha, 402, 2)); // Sigla da UF
              dest.EnderDest.CEP := strtoint(Copy(Linha, 404, 8));
              // Código do Cep
              dest.EnderDest.cPais := strtoint(Copy(Linha, 412, 4));
              // Código do País
              dest.EnderDest.xPais := trim(Copy(Linha, 416, 60)); // Brasil
              dest.EnderDest.fone := trim(Copy(Linha, 476, 10)); // Telefone
              dest.ie := trim(Copy(Linha, 486, 14)); // IE
              dest.ISUF := trim(Copy(Linha, 500, 12)); // Inscrição SUFRAMA}

              dest.email := trim(Copy(Linha, 548, 60));
              // Email para Recepção de XML
              ACBrNFe1.NotasFiscais[0].NFe.dest.email :=
                trim(Copy(Linha, 548, 60));

              dest.idEstrangeiro := trim(Copy(Linha, 512, 20));
              // Inscrição de Estrangeiro}

              case strtoint(Copy(Linha, 532, 1)) of
                1:
                  dest.indIEDest := inContribuinte;
                2:
                  dest.indIEDest := inIsento;
                9:
                  dest.indIEDest := inNaoContribuinte;
              end; // Indica Contribuinte

              dest.IM := trim(Copy(Linha, 533, 15)); // Inscrição municiapl
            end; // fim do EM0204

      // ==============================EM0205=============================//
      // =Endereço de Entrega

      if (Copy(Linha, 0, 6)) = 'EM0205' then
      begin

        // if ((ValidaCNPJ(copy(Linha, 7, 14), False) = true) and (copy(Linha, 7, 14) <> '              ')) then
        // begin
        Entrega.CNPJCPF := (Copy(Linha, 7, 14));
        // end;
        Entrega.xLgr := trim(Copy(Linha, 21, 60));
        Entrega.nro := trim(Copy(Linha, 81, 60));
        Entrega.xCpl := trim(Copy(Linha, 141, 60));
        Entrega.xBairro := trim(Copy(Linha, 201, 60));
        Entrega.cMun := strtoint(Copy(Linha, 261, 7));
        Entrega.xMun := trim(Copy(Linha, 268, 60));
        Entrega.uf := trim(Copy(Linha, 328, 2));

      end; // fim do EM0205

      // =============================EM0206=============================//

      if (Copy(Linha, 1, 6)) = 'EM0206' then
      begin

        with Det.Add do
        begin
          Prod.nItem := strtoint(Copy(Linha, 9, 3));
          // Nro. do item
          Prod.cProd := trim(Copy(Linha, 12, 60));
          // Código do Produto ou serviço
{$REGION 'Código Barra'}
          codigoBarra := trim(Copy(Linha, 72, 14)); // cEAN

          if trim(codigoBarra) = '' then
            codigoBarra := 'SEM GTIN';

          Prod.cEAN := trim(codigoBarra); // cEAN
{$ENDREGION}
          Prod.xProd := trim(Copy(Linha, 86, 120));
          // Descrição do produto ou serviço
          Prod.NCM := trim(Copy(Linha, 206, 8)); // Código NCM
          Prod.EXTIPI := trim(Copy(Linha, 214, 3)); // EX_TIPI
          Prod.CFOP := trim(Copy(Linha, 219, 4));
          // Prod.CEST := trim(Copy(Linha, 219, 4));
          // Código fiscal da operação

          if trim(Copy(Linha, 223, 6)) <> '' then
            Prod.uCom := trim(Copy(Linha, 223, 6)) // Unidade comercial
          else
            Prod.uCom := '0';

          Prod.qCom := StrToFloat(StringReplace((Copy(Linha, 229, 15)), '.',
            ',', [rfReplaceAll]));
          // Quantidade comercial
          Prod.vUnCom := StrToFloat(StringReplace((Copy(Linha, 244, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor unitário de comercialização
          Prod.vProd := StrToFloat(StringReplace((Copy(Linha, 259, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor Total Bruto dos Produtos ou Serviços
          // if trim((Copy(Linha, 274, 14))) = '' then
          // Prod.cEANTrib :=  'SEM GTIN' // cEANTrib
          // else
          Prod.cEANTrib := codigoBarra; // (Copy(Linha, 274, 14)); // cEANTrib

          unidadeTributacao := trim(Copy(Linha, 288, 6));
          if unidadeTributacao <> '' then
          begin
            if unidadeTributacao = 'un' then
              Prod.uTrib := 'und'
            else
              Prod.uTrib := unidadeTributacao;
          end
          else // Unidade Tributavel
            Prod.uTrib := '0';

          // Prod.uTrib := 'und';
          Prod.qTrib := StrToFloat(StringReplace((Copy(Linha, 294, 15)), '.',
            ',', [rfReplaceAll]));
          // UQuantidade Tributavel
          Prod.vUnTrib := StrToFloat(StringReplace((Copy(Linha, 309, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor Unitário de tributação
          Prod.vFrete := StrToFloat(StringReplace((Copy(Linha, 324, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor Total do Frete
          Prod.vSeg := StrToFloat(StringReplace((Copy(Linha, 339, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor Total do Seguro
          Prod.vOutro := StrToFloat(StringReplace((Copy(Linha, 354, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor do outras despesas acessorias
          Prod.vDesc := StrToFloat(StringReplace((Copy(Linha, 369, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor do outras Desconto
          FCI := trim(Copy(Linha, 517, 50));
          if FCI <> '' then
            Prod.nFCI := FCI;
          // FCI - Número de controle da FCI - Ficha de Conteúdo de Importação

          /// //// Imposto de importação

          if length(Linha) > 385 then
          begin
            if trim(Copy(Linha, 384, 15)) <> '' then
              Imposto.II.vBC := StrToFloat(StringReplace((Copy(Linha, 384, 15)),
                '.', ',', [rfReplaceAll]));
            // Valor da Base de Calculo para II
            if trim(Copy(Linha, 399, 15)) <> '' then
              Imposto.II.vII := StrToFloat(StringReplace((Copy(Linha, 399, 15)),
                '.', ',', [rfReplaceAll]));
            // Valor do Imposto de importação

            if trim(Copy(Linha, 414, 15)) <> '' then
              Imposto.II.vDespAdu :=
                StrToFloat(StringReplace((Copy(Linha, 414, 15)), '.', ',',
                [rfReplaceAll]));
            // Valor de Despesas Aduaneiras

            if trim(Copy(Linha, 429, 15)) <> '' then
              Imposto.II.vIOF :=
                StrToFloat(StringReplace((Copy(Linha, 429, 15)), '.', ',',
                [rfReplaceAll]));
            // Valor do IOF

            // Informar o valor aproximado em cada item
            if trim(Copy(Linha, 444, 15)) <> '' then
              Imposto.vTotTrib :=
                StrToFloat(StringReplace((Copy(Linha, 444, 15)), '.', ',',
                [rfReplaceAll]));

          end;

          // Pedido de Compra
          if length(Linha) > 445 then
          begin
            if trim(Copy(Linha, 459, 15)) <> '' then
              Prod.xPed := trim(Copy(Linha, 459, 15));

            if trim(Copy(Linha, 474, 6)) <> '' then
              Prod.nItemPed := trim(Copy(Linha, 474, 6));
          end;

          // Pedido de Compra
          if length(Linha) > 480 then
          begin
            if trim(Copy(Linha, 480, 7)) <> '' then
              Prod.CEST := trim(Copy(Linha, 480, 7));
          end;

          // ANP Combustiveis
          if length(Linha) > 487 then
          begin
            if trim(Copy(Linha, 487, 9)) <> '' then
              Prod.comb.cProdANP := strtoint(trim(Copy(Linha, 487, 9)));

            if trim(Copy(Linha, 567, 75)) <> '' then
              Prod.comb.descANP := trim(Copy(Linha, 567, 75));

            if trim(Entrega.uf) <> '' then
              Prod.comb.UFCons := Entrega.uf
            else
              Prod.comb.UFCons := dest.EnderDest.uf;
          end;

          if length(Linha) > 496 then
          begin
            if trim(Copy(Linha, 496, 21)) <> '' then
              Prod.comb.CODIF := trim(Copy(Linha, 496, 21));
          end;

          // Ler aProxima linha pra chamar o 207
          Readln(Arquivo, Linha); // Ler a proxima linha
          // =============================EM1206=============================//
          if (Copy(Linha, 0, 6)) = 'EM1206' then
          begin
            infAdProd := (Copy(Linha, 7, 506));
          end; // fim do EM1206
          // Ler aProxima linha pra chamar o 207
          Readln(Arquivo, Linha); // Ler a proxima linha
          // =============================EM1207  DI=============================//

          if (Copy(Linha, 0, 6)) = 'EM1207' then
          begin
            while (Copy(Linha, 0, 6)) <> 'EM0207' do
            begin
              if (Copy(Linha, 0, 6)) = 'EM1207' then
              begin
                if Prod.CFOP[1] in ['3'] then
                begin
                  with Prod.DI.Add do
                  begin
                    nDi := Copy(Linha, 7, 10);
                    dDi := strtodate(Copy(Linha, 17, 10));
                    xLocDesemb := Copy(Linha, 27, 60);
                    UFDesemb := Copy(Linha, 87, 2);
                    dDesemb := strtodate(Copy(Linha, 89, 10));
                    cExportador := Copy(Linha, 99, 60);
                    Readln(Arquivo, Linha);
                    // Ler a proxima linha
                    while (Copy(Linha, 0, 6)) = 'EM1208' do
                    begin
                      Cdi := 1;
                      if Prod.CFOP[1] in ['3'] then
                        with adi.Add do
                        begin

                          begin
                            nSeqAdi := strtoint(trim((Copy(Linha, 7, 3))));
                            nAdicao := strtoint(trim((Copy(Linha, 10, 3))));
                            cFabricante := trim(Copy(Linha, 13, 60));
                            vDescDI :=
                              StrToFloat(StringReplace((Copy(Linha, 73, 15)),
                              '.', ',', [rfReplaceAll]));
                          end;
                        end;

                      Readln(Arquivo, Linha); // Ler a proxima linha

                    end;
                  end;
                end
                else
                begin
                  Readln(Arquivo, Linha); // Ler a proxima linha
                end;
              end
              else
              begin
                Readln(Arquivo, Linha); // Ler a proxima linha
              end;
            end;
          end;

          if (Copy(Linha, 0, 6)) = 'EM3207' then
          begin
            while (Copy(Linha, 0, 6)) = 'EM3207' do
            begin
              with Prod.rastro.Add do
              // quando usar Medicamentos usa Prod.med, nosso caso é só rastreamento então Pro.rastro
              begin
                nLote := trim(Copy(Linha, 7, 20));
                qLote := StrToFloat(StringReplace(trim(Copy(Linha, 27, 15)),
                  '.', ',', [rfReplaceAll]));
                dFab := strtodate(Copy(Linha, 42, 10));
                dVal := strtodate(Copy(Linha, 52, 10));
                // vPMC := StrToFloat(StringReplace(trim(Copy(Linha, 62, 15)), '.',
                // ',', [rfReplaceAll]));
              end;

              Readln(Arquivo, Linha); // Ler a proxima linha
            end;
          end;

          if (Copy(Linha, 0, 6)) = 'EM4207' then
          begin
            while (Copy(Linha, 0, 6)) = 'EM4207' do
            begin
              with Prod.med.Add do
              begin
                nLote := trim(Copy(Linha, 7, 20));
                qLote := StrToFloat(StringReplace(trim(Copy(Linha, 27, 15)),
                  '.', ',', [rfReplaceAll]));
                dFab := strtodate(Copy(Linha, 42, 10));
                dVal := strtodate(Copy(Linha, 52, 10));
                vPMC := StrToFloat(StringReplace(trim(Copy(Linha, 62, 15)), '.',
                  ',', [rfReplaceAll]));
              end;

              Readln(Arquivo, Linha); // Ler a proxima linha
            end;
          end;

          if (Copy(Linha, 0, 6)) = 'EM0207' then
          begin

            if not SN then
            begin

              try
                strtoint(Copy(Linha, 12, 1));
              except
                messagebox(Handle,
                  'Origem do produto não encontrada. Verifique e tente novamente.',
                  'Envio NFE', MB_OK + MB_ICONINFORMATION);
                Abort;
              end;

              case strtoint(Copy(Linha, 12, 1)) of
                0:
                  Imposto.ICMS.orig := oeNacional; // Origem da mercadoria
                1:
                  Imposto.ICMS.orig := oeEstrangeiraImportacaoDireta;
                2:
                  Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasil;
                3:
                  Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior40;
                // Origem da mercadoria
                4:
                  Imposto.ICMS.orig := oeNacionalProcessosBasicos;
                5:
                  Imposto.ICMS.orig :=
                    oeNacionalConteudoImportacaoInferiorIgual40;
                6:
                  Imposto.ICMS.orig := oeEstrangeiraImportacaoDiretaSemSimilar;
                7:
                  Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasilSemSimilar;
                8:
                  Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior70;

              end;

              try
                strtoint(Copy(Linha, 13, 2));
              except
                messagebox(Handle,
                  'Situação tributária do ICMS não encontrada. Verifique e tente novamente.',
                  'Envio NFE', MB_OK + MB_ICONINFORMATION);
                Abort;
              end;

              case strtoint(Copy(Linha, 13, 2)) of
                00:
                  Imposto.ICMS.CST := cst00;
                10:
                  Imposto.ICMS.CST := cst10;
                20:
                  Imposto.ICMS.CST := cst20;
                30:
                  Imposto.ICMS.CST := cst30;
                40:
                  Imposto.ICMS.CST := cst40;
                41:
                  Imposto.ICMS.CST := cst41;
                45:
                  Imposto.ICMS.CST := cst45;
                50:
                  Imposto.ICMS.CST := cst50;
                51:
                  Imposto.ICMS.CST := cst51;
                60:
                  Imposto.ICMS.CST := cst60;
                70:
                  Imposto.ICMS.CST := cst70;
                80:
                  Imposto.ICMS.CST := cst80;
                81:
                  Imposto.ICMS.CST := cst81;
                90:
                  Imposto.ICMS.CST := cst90;
              end;

              Imposto.ICMS.modBCST := dbisPrecoTabelado;
              // Modalidade de determinação da BC do ICMS ST
              Imposto.ICMS.pRedBC :=
                StrToFloat(StringReplace((Copy(Linha, 17, 15)), '.', ',',
                [rfReplaceAll]));
              // Percential de redução de BC do ICMS

              Imposto.ICMS.vBC :=
                StrToFloat(StringReplace((Copy(Linha, 32, 15)), '.', ',',
                [rfReplaceAll]));

              // Valor da BC do ICMS
              Imposto.ICMS.pICMS :=
                StrToFloat(StringReplace((Copy(Linha, 46, 5)), '.', ',',
                [rfReplaceAll]));
              // Aliquota do imposto
              Imposto.ICMS.vICMS :=
                StrToFloat(StringReplace((Copy(Linha, 53, 15)), '.', ',',
                [rfReplaceAll]));
              // Valor do ICMS
              Imposto.ICMS.vBCST :=
                StrToFloat(StringReplace((Copy(Linha, 67, 15)), '.', ',',
                [rfReplaceAll]));

              Imposto.ICMS.vBCST :=
                StrToFloat(StringReplace((Copy(Linha, 67, 15)), '.', ',',
                [rfReplaceAll]));

              // Valor da BC do ICMS ST ----------mexi aqui
              Imposto.ICMS.pICMSST :=
                StrToFloat(StringReplace((Copy(Linha, 81, 5)), '.', ',',
                [rfReplaceAll]));
              // Aliquota do imposto do ICMS ST 5
              Imposto.ICMS.pMVAST :=
                StrToFloat(StringReplace((Copy(Linha, 87, 5)), '.', ',',
                [rfReplaceAll]));

              // Percentual da Margem de valor Adicionado do ICMS ST 5
              Imposto.ICMS.vICMSST :=
                StrToFloat(StringReplace((Copy(Linha, 92, 15)), '.', ',',
                [rfReplaceAll]));
              // Valor do ICMS ST 15
              // Imposto.ICMS.modBC := dbiValorOperacao;
              // Valor do ICMS Desonerado
              Imposto.ICMS.vICMSDeson :=
                StrToFloat(StringReplace((Copy(Linha, 107, 15)), '.', ',',
                [rfReplaceAll]));

              if ((Imposto.ICMS.CST = cst60) and
                (Ide.indFinal <> cfConsumidorFinal)) then
              begin
                if (StrToFloat(StringReplace((Copy(Linha, 198, 15)), '.', ',',
                  [rfReplaceAll])) > 0) then
                begin
                  Imposto.ICMS.vBCSTRet :=
                    StrToFloat(StringReplace((Copy(Linha, 153, 15)), '.', ',',
                    [rfReplaceAll]));
                  Imposto.ICMS.vICMSSubstituto :=
                    StrToFloat(StringReplace((Copy(Linha, 168, 15)), '.', ',',
                    [rfReplaceAll]));
                  Imposto.ICMS.vICMSSTRet :=
                    StrToFloat(StringReplace((Copy(Linha, 183, 15)), '.', ',',
                    [rfReplaceAll]));
                  Imposto.ICMS.pST :=
                    StrToFloat(StringReplace((Copy(Linha, 198, 15)), '.', ',',
                    [rfReplaceAll]));
                end;
              end;

              case strtoint(Copy(Linha, 121, 2)) of
                01:
                  Imposto.ICMS.motDesICMS := mdiTaxi;
                03:
                  Imposto.ICMS.motDesICMS := mdiProdutorAgropecuario;
                04:
                  Imposto.ICMS.motDesICMS := mdiFrotistaLocadora;
                05:
                  Imposto.ICMS.motDesICMS := mdiDiplomaticoConsular;
                06:
                  Imposto.ICMS.motDesICMS := mdiAmazoniaLivreComercio;
                07:
                  Imposto.ICMS.motDesICMS := mdiSuframa;
                08:
                  Imposto.ICMS.motDesICMS := mdiVendaOrgaosPublicos;
                09:
                  Imposto.ICMS.motDesICMS := mdiOutros;
                10:
                  Imposto.ICMS.motDesICMS := mdiDeficienteCondutor;
                11:
                  Imposto.ICMS.motDesICMS := mdiDeficienteNaoCondutor;
                12:
                  Imposto.ICMS.motDesICMS := mdiOrgaoFomento;
              end;

            end
            else
            begin
              try
                strtoint(Copy(Linha, 12, 1));
              except
                messagebox(Handle,
                  'Origem do produto não encontrada. Verifique e tente novamente.',
                  'Envio NFE', MB_OK + MB_ICONINFORMATION);
                Abort;
              end;
              case strtoint(Copy(Linha, 12, 1)) of
                0:
                  Imposto.ICMS.orig := oeNacional; // Origem da mercadoria
                1:
                  Imposto.ICMS.orig := oeEstrangeiraImportacaoDireta;
                2:
                  Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasil;
                3:
                  Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior40;
                // Origem da mercadoria
                4:
                  Imposto.ICMS.orig := oeNacionalProcessosBasicos;
                5:
                  Imposto.ICMS.orig :=
                    oeNacionalConteudoImportacaoInferiorIgual40;
                6:
                  Imposto.ICMS.orig := oeEstrangeiraImportacaoDiretaSemSimilar;
                7:
                  Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasilSemSimilar;
                8:
                  Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior70;
              end;

              try
                strtoint(Copy(Linha, 13, 3));
              except
                messagebox(Handle,
                  'Situação tributária do ICMS não encontrada. Verifique e tente novamente.',
                  'Envio NFE', MB_OK + MB_ICONINFORMATION);
                Abort;
              end;
              case strtoint(Copy(Linha, 13, 3)) of
                101:
                  Imposto.ICMS.CSOSN := csosn101;
                102:
                  Imposto.ICMS.CSOSN := csosn102;
                103:
                  Imposto.ICMS.CSOSN := csosn103;
                201:
                  Imposto.ICMS.CSOSN := csosn201;
                202:
                  Imposto.ICMS.CSOSN := csosn202;
                203:
                  Imposto.ICMS.CSOSN := csosn203;
                300:
                  Imposto.ICMS.CSOSN := csosn300;
                400:
                  Imposto.ICMS.CSOSN := csosn400;
                500:
                  Imposto.ICMS.CSOSN := csosn500;
                900:
                  Imposto.ICMS.CSOSN := csosn900;
              end;
              Imposto.ICMS.modBCST := dbisPrecoTabelado;
              // Modalidade de determinação da BC do ICMS ST
              Imposto.ICMS.pRedBC :=
                StrToFloat(StringReplace((Copy(Linha, 17, 15)), '.', ',',
                [rfReplaceAll]));
              // Percential de redução de BC do ICMS
              Imposto.ICMS.vBC :=
                StrToFloat(StringReplace((Copy(Linha, 32, 15)), '.', ',',
                [rfReplaceAll]));
              // Valor da BC do ICMS
              Imposto.ICMS.pICMS :=
                StrToFloat(StringReplace((Copy(Linha, 47, 7)), '.', ',',
                [rfReplaceAll]));
              // Aliquota do imposto

              (* --simples-- *)
              Imposto.ICMS.pCredSN :=
                StrToFloat(StringReplace((Copy(Linha, 47, 7)), '.', ',',
                [rfReplaceAll]));
              // Aliquota do imposto
              VSNAliq := StrToFloat(StringReplace((Copy(Linha, 47, 7)), '.',
                ',', [rfReplaceAll]));
              // Aliquota do imposto
              Imposto.ICMS.vICMS :=
                StrToFloat(StringReplace((Copy(Linha, 54, 13)), '.', ',',
                [rfReplaceAll]));
              // Valor do ICMS
              (* --simples-- *)
              // Imposto.ICMS.vCredICMSSN := VSNprod * (VSNAliq / 100);
              Imposto.ICMS.vCredICMSSN :=
                StrToFloat(StringReplace((Copy(Linha, 54, 13)), '.', ',',
                [rfReplaceAll]));

              if Copy(Linha, 13, 3) <> '500' then
              begin
                // Aliquota do imposto
                Imposto.ICMS.vBCST :=
                  StrToFloat(StringReplace((Copy(Linha, 67, 15)), '.', ',',
                  [rfReplaceAll]));
                // Valor da BC do ICMS ST
                Imposto.ICMS.pICMSST :=
                  StrToFloat(StringReplace((Copy(Linha, 82, 5)), '.', ',',
                  [rfReplaceAll]));
                // Aliquota do imposto do ICMS ST 5
                Imposto.ICMS.pMVAST :=
                  StrToFloat(StringReplace((Copy(Linha, 87, 5)), '.', ',',
                  [rfReplaceAll]));
                // Percentual da Margem de valor Adicionado do ICMS ST 5
                Imposto.ICMS.vICMSST :=
                  StrToFloat(StringReplace((Copy(Linha, 92, 15)), '.', ',',
                  [rfReplaceAll]));

              end
              else
              begin
                Imposto.ICMS.vBCSTRet :=
                  StrToFloat(StringReplace((Copy(Linha, 67, 15)), '.', ',',
                  [rfReplaceAll]));
                // Valor da BC do ICMS ST
                Imposto.ICMS.pICMSST :=
                  StrToFloat(StringReplace((Copy(Linha, 82, 5)), '.', ',',
                  [rfReplaceAll]));
                // Aliquota do imposto do ICMS ST 5
                Imposto.ICMS.pMVAST :=
                  StrToFloat(StringReplace((Copy(Linha, 87, 5)), '.', ',',
                  [rfReplaceAll]));
                // Percentual da Margem de valor Adicionado do ICMS ST 5
                Imposto.ICMS.vICMSSTRet :=
                  StrToFloat(StringReplace((Copy(Linha, 92, 15)), '.', ',',
                  [rfReplaceAll]));
                // Valor do ICMS Desonerado
                Imposto.ICMS.vICMSDeson :=
                  StrToFloat(StringReplace((Copy(Linha, 107, 15)), '.', ',',
                  [rfReplaceAll]));

                if length(Linha) > 124 then
                begin
                  Imposto.ICMS.pCredSN :=
                    StrToFloat(StringReplace((Copy(Linha, 124, 15)), '.', ',',
                    [rfReplaceAll]));
                end;

                if length(Linha) > 139 then
                begin
                  Imposto.ICMS.vCredICMSSN :=
                    StrToFloat(StringReplace((Copy(Linha, 139, 15)), '.', ',',
                    [rfReplaceAll]));
                end;

              end;
              case strtoint(Copy(Linha, 122, 2)) of
                01:
                  Imposto.ICMS.motDesICMS := mdiTaxi;
                03:
                  Imposto.ICMS.motDesICMS := mdiProdutorAgropecuario;
                04:
                  Imposto.ICMS.motDesICMS := mdiFrotistaLocadora;
                05:
                  Imposto.ICMS.motDesICMS := mdiDiplomaticoConsular;
                06:
                  Imposto.ICMS.motDesICMS := mdiAmazoniaLivreComercio;
                07:
                  Imposto.ICMS.motDesICMS := mdiSuframa;
                08:
                  Imposto.ICMS.motDesICMS := mdiVendaOrgaosPublicos;
                09:
                  Imposto.ICMS.motDesICMS := mdiOutros;
                10:
                  Imposto.ICMS.motDesICMS := mdiDeficienteCondutor;
                11:
                  Imposto.ICMS.motDesICMS := mdiDeficienteNaoCondutor;
                12:
                  Imposto.ICMS.motDesICMS := mdiOrgaoFomento;
              end;

            end;

          end; // fim do EM0207
          // Ler aProxima linha pra chamar o 208
          Readln(Arquivo, Linha); // Ler a proxima linha

          if (Copy(Linha, 0, 6)) = 'EM2207' then
          begin
            // if Ide.idDest = doInterna then
            // begin
            Pfcp := StrToFloat(StringReplace((Copy(Linha, 22, 15)), '.', ',',
              [rfReplaceAll]));
            // Percentual do ICMS relativo ao Fundo de Combate à Pobreza (FCP) na UF de destino

            // Percentual do Fundo de Combate à Pobreza (FCP)
            // Imposto.ICMS.pFCP :=StrToFloat(StringReplace((Copy(Linha, 22, 15)), '.', ',', [rfReplaceAll]));
            // Valor do Fundo de Combate à Pobreza (FCP)
            if Imposto.ICMS.CST = cst00 then
            begin

              // Vfcp := Imposto.ICMS.vBC * (Pfcp / 100); //Calculo
              Vfcp := StrToFloat(StringReplace((Copy(Linha, 82, 15)), '.', ',',
                [rfReplaceAll])); // Valor resgatado do arquivo

              if Vfcp <> 0 then
              begin
                if NOT((Ide.idDest = doInterestadual) and
                  (Ide.indFinal = cfConsumidorFinal)) then
                begin
                  Imposto.ICMS.vBCFCP := Imposto.ICMS.vBC; // Mesma Base ??????
                  Imposto.ICMS.Pfcp := Pfcp;
                  Imposto.ICMS.Vfcp := Vfcp;
                end
                else
                begin
                  if dest.indIEDest = inNaoContribuinte then
                  begin
                    Imposto.ICMS.vBCFCP := Imposto.ICMS.vBC;
                    // Mesma Base ??????
                    Imposto.ICMSUFDest.pFCPUFDest := Pfcp;
                    Imposto.ICMSUFDest.vFCPUFDest := Vfcp;
                  end

                end;
              end;
            end
            else if Imposto.ICMS.CST in [cst10, cst20, cst70, cst90, cst51] then
            begin
              Vfcp := (Imposto.ICMS.vBCFCP * Pfcp) / 100;

              if Vfcp <> 0 then
              begin

                if NOT((Ide.idDest = doInterestadual) and
                  (Ide.indFinal = cfConsumidorFinal)) then
                begin
                  Imposto.ICMS.vBCFCP := Imposto.ICMS.vBC; // Mesma Base ??????
                  Imposto.ICMS.Pfcp := Pfcp;
                  Imposto.ICMS.Vfcp := Vfcp;
                end
                else
                begin
                  if dest.indIEDest = inNaoContribuinte then
                  begin
                    Imposto.ICMS.vBCFCP := Imposto.ICMS.vBC;
                    // Mesma Base ??????
                    Imposto.ICMSUFDest.pFCPUFDest := Pfcp;
                    Imposto.ICMSUFDest.vFCPUFDest := Vfcp;
                  end
                end;
              end;

            end;

            // Imposto.ICMSUFDest.vFCPUFDest := Vfcp;// StrToFloat(StringReplace((Copy(Linha, 82, 15)), '.', ',',
            // ShowMessage('VALOR vFCPUFDest ' +floattostr(Imposto.ICMSUFDest.vFCPUFDest));
            // [rfReplaceAll])); // Valor do ICMS relativo ao Fundo de Combate à Pobreza (FCP) da UF de destino

            Imposto.ICMSUFDest.vBCUFDest :=
              StrToFloat(StringReplace((Copy(Linha, 7, 15)), '.', ',',
              [rfReplaceAll])); // Valor da BC do ICMS na UF de destino

            Imposto.ICMSUFDest.vBCFCPUFDest := Imposto.ICMSUFDest.vBCUFDest;

            Imposto.ICMSUFDest.pICMSUFDest :=
              StrToFloat(StringReplace((Copy(Linha, 37, 15)), '.', ',',
              [rfReplaceAll])); // Alíquota interna da UF de destino

            Imposto.ICMSUFDest.pICMSInter :=
              StrToFloat(StringReplace((Copy(Linha, 52, 15)), '.', ',',
              [rfReplaceAll])); // Alíquota interestadual das UF envolvidas

            Imposto.ICMSUFDest.pICMSInterPart :=
              StrToFloat(StringReplace((Copy(Linha, 67, 15)), '.', ',',
              [rfReplaceAll]));
            // Percentual provisório de partilha do ICMS Interestadual
            // ShowMessage('VALOR pICMSInterPart ' + floattostr(Imposto.ICMSUFDest.pICMSInterPart));

            Imposto.ICMSUFDest.vICMSUFDest :=
              StrToFloat(StringReplace((Copy(Linha, 97, 15)), '.', ',',
              [rfReplaceAll]));
            // Valor do ICMS Interestadual para a UF de destino
            // ShowMessage('Valor vICMSUFDest ' + floattostr(Imposto.ICMSUFDest.vICMSUFDest));

            Imposto.ICMSUFDest.vICMSUFRemet :=
              StrToFloat(StringReplace((Copy(Linha, 112, 15)), '.', ',',
              [rfReplaceAll]));
            // Valor do ICMS Interestadual para a UF do remetente
            // ShowMessage('Valor vICMSUFRemet ' + floattostr(Imposto.ICMSUFDest.vICMSUFRemet));

            // Ler aProxima linha pra chamar o 209
            Readln(Arquivo, Linha); // Ler a proxima linha
          end; // fim do EM2207

          // =============================EM0208=============================//
          if (Copy(Linha, 0, 6)) = 'EM0208' then
          begin

            Imposto.IPI.vBC := StrToFloat(StringReplace((Copy(Linha, 15, 15)),
              '.', ',', [rfReplaceAll]));
            // 15 Valor da BC do IPI
            Imposto.IPI.pIPI := StrToFloat(StringReplace((Copy(Linha, 30, 5)),
              '.', ',', [rfReplaceAll]));
            // 5 Aliquota do imposto
            Imposto.IPI.vIPI := StrToFloat(StringReplace((Copy(Linha, 35, 15)),
              '.', ',', [rfReplaceAll]));
            // 15 Valor do IPI

            // CASE   Situação tributária do IPI
            try
              strtoint(Copy(Linha, 50, 2))
            except
              messagebox(Handle,
                'Situação de IPI não encontrada. Verifique e tente novamente.',
                'Envio NFE', MB_OK + MB_ICONINFORMATION);
              Abort;
            end;

            case strtoint(Copy(Linha, 50, 2)) of
              00:
                Imposto.IPI.CST := ipi00;
              49:
                Imposto.IPI.CST := ipi49;
              50:
                Imposto.IPI.CST := ipi50;
              99:
                Imposto.IPI.CST := ipi99;
              01:
                Imposto.IPI.CST := ipi01;
              02:
                Imposto.IPI.CST := ipi02;
              03:
                Imposto.IPI.CST := ipi03;
              04:
                Imposto.IPI.CST := ipi04;
              05:
                Imposto.IPI.CST := ipi05;
              51:
                Imposto.IPI.CST := ipi51;
              52:
                Imposto.IPI.CST := ipi52;
              53:
                Imposto.IPI.CST := ipi53;
              54:
                Imposto.IPI.CST := ipi54;
              55:
                Imposto.IPI.CST := ipi55;
            end;
          end; // fim do EM0208

          // Ler aProxima linha pra chamar o 209
          Readln(Arquivo, Linha); // Ler a proxima linha
          // =============================EM0209=============================//
          if (Copy(Linha, 0, 6)) = 'EM0209' then
          begin

            // CASEEEEE 02 Situação Tributaria do PIS       showmessage(copy(EM0209,12,2));
            case strtoint(Copy(Linha, 12, 2)) of
              01:
                Imposto.PIS.CST := pis01;
              02:
                Imposto.PIS.CST := pis02;
              03:
                Imposto.PIS.CST := pis03;
              04:
                Imposto.PIS.CST := pis04;
              06:
                Imposto.PIS.CST := pis06;
              07:
                Imposto.PIS.CST := pis07;
              08:
                Imposto.PIS.CST := pis08;
              09:
                Imposto.PIS.CST := pis09;
              99:
                Imposto.PIS.CST := pis99;
            end;

            Imposto.PIS.vBC := StrToFloat(StringReplace((Copy(Linha, 17, 15)),
              '.', ',', [rfReplaceAll]));
            // 15 BC PIS
            Imposto.PIS.pPIS := StrToFloat(StringReplace((Copy(Linha, 32, 5)),
              '.', ',', [rfReplaceAll]));
            // 5 Percentual do PIS
            Imposto.PIS.vPIS := StrToFloat(StringReplace((Copy(Linha, 37, 15)),
              '.', ',', [rfReplaceAll]));
            // 15 Valor do PIS
            // CASE Situação Tributaria do COFINS       showmessage(copy(EM0209,52,2));

            case strtoint(Copy(Linha, 52, 2)) of
              01:
                Imposto.COFINS.CST := cof01;
              02:
                Imposto.COFINS.CST := cof02;
              03:
                Imposto.COFINS.CST := cof03;
              04:
                Imposto.COFINS.CST := cof04;
              06:
                Imposto.COFINS.CST := cof06;
              07:
                Imposto.COFINS.CST := cof07;
              08:
                Imposto.COFINS.CST := cof08;
              09:
                Imposto.COFINS.CST := cof09;
              99:
                Imposto.COFINS.CST := cof99;
            end;
            Imposto.COFINS.vBC :=
              StrToFloat(StringReplace((Copy(Linha, 57, 15)), '.', ',',
              [rfReplaceAll]));
            // 15 BC COFINS
            Imposto.COFINS.pCOFINS :=
              StrToFloat(StringReplace((Copy(Linha, 72, 5)), '.', ',',
              [rfReplaceAll]));
            // 5 Percentual do COFINS
            Imposto.COFINS.vCOFINS :=
              StrToFloat(StringReplace((Copy(Linha, 77, 15)), '.', ',',
              [rfReplaceAll]));
            // 15 Valor do COFINS
          end; // fim do EM0209
        end; // with do Add.item
      end; // fim do EM0206

      // =============================EM0210=============================//
      if (Copy(Linha, 1, 6)) = 'EM0210' then
      begin
        // Base de Calculo do ICMS
        Total.ICMSTot.vBC := StrToFloat(StringReplace((Copy(Linha, 7, 15)), '.',
          ',', [rfReplaceAll]));
        // Valor Total do ICMS
        Total.ICMSTot.vICMS := StrToFloat(StringReplace((Copy(Linha, 22, 15)),
          '.', ',', [rfReplaceAll]));
        // Base de Calculo do ICMS ST   showmessage(copy(linha,38,15));
        Total.ICMSTot.vBCST := StrToFloat(StringReplace((Copy(Linha, 37, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total do ICMS ST  showmessage(copy(linha,53,15));
        Total.ICMSTot.vST := StrToFloat(StringReplace((Copy(Linha, 52, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total dos produtos e serviços  showmessage(copy(linha,68,15));
        Total.ICMSTot.vProd := StrToFloat(StringReplace((Copy(Linha, 67, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total do Frete       showmessage(copy(linha,83,15));
        Total.ICMSTot.vFrete := StrToFloat(StringReplace((Copy(Linha, 82, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total do Seguro      showmessage(copy(linha,98,15));
        Total.ICMSTot.vSeg := StrToFloat(StringReplace((Copy(Linha, 97, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total do Desconto     showmessage(copy(linha,113,15));
        Total.ICMSTot.vDesc := StrToFloat(StringReplace((Copy(Linha, 112, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total do II           showmessage(copy(linha,128,15));
        Total.ICMSTot.vII := StrToFloat(StringReplace((Copy(Linha, 127, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total do IPI        showmessage(copy(linha,143,15));
        Total.ICMSTot.vIPI := StrToFloat(StringReplace((Copy(Linha, 142, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total do PIS         showmessage(copy(linha,158,15));
        Total.ICMSTot.vPIS := StrToFloat(StringReplace((Copy(Linha, 157, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total do COFINS      showmessage(copy(linha,173,15));
        Total.ICMSTot.vCOFINS :=
          StrToFloat(StringReplace((Copy(Linha, 172, 15)), '.', ',',
          [rfReplaceAll]));
        // Outras Despesas Acessórias showmessage(copy(linha,188,15));
        Total.ICMSTot.vOutro := StrToFloat(StringReplace((Copy(Linha, 187, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total da NFe       showmessage(copy(linha,203,15));
        Total.ICMSTot.vNF := StrToFloat(StringReplace((Copy(Linha, 202, 15)),
          '.', ',', [rfReplaceAll]));

        // Informar a soma no cabeçalho da nota
        // if (length(Linha) > 222) and (trim(Copy(Linha, 217, 15)) <> '') then
        Total.ICMSTot.vTotTrib :=
          StrToFloat(StringReplace((Copy(Linha, 217, 15)), '.', ',',
          [rfReplaceAll]));
        // Valor total Desonerado
        Total.ICMSTot.vICMSDeson :=
          StrToFloat(StringReplace((Copy(Linha, 232, 15)), '.', ',',
          [rfReplaceAll]));

        if (length(Linha) > 247) then
        begin
          Total.ICMSTot.vFCPUFDest :=
            StrToFloat(StringReplace((Copy(Linha, 247, 15)), '.', ',',
            [rfReplaceAll]));
          // ShowMessage('Total.ICMSTot.vFCPUFDest ' + floattostr(Total.ICMSTot.vFCPUFDest));

          Total.ICMSTot.vICMSUFDest :=
            StrToFloat(StringReplace((Copy(Linha, 262, 15)), '.', ',',
            [rfReplaceAll]));
          // ShowMessage('Total.ICMSTot.vICMSUFDest ' + floattostr(Total.ICMSTot.vICMSUFDest));

          Total.ICMSTot.vICMSUFRemet :=
            StrToFloat(StringReplace((Copy(Linha, 277, 15)), '.', ',',
            [rfReplaceAll]));
          // ShowMessage('Total.ICMSTot.vICMSUFRemet ' + floattostr(Total.ICMSTot.vICMSUFRemet));

        end;

      end; // fim do EM0210
      // =============================EM0211=============================//
      if (Copy(Linha, 1, 6)) = 'EM0211' then
      begin

        if (Copy(Linha, 7, 1)) = '0' then
        // 1  Modalidade do Frete                showmessage(copy(linha,7,1));
        begin
          Transp.modFrete := mfContaEmitente;
        end
        else if (Copy(Linha, 7, 1)) = '1' then
        begin
          Transp.modFrete := mfContaDestinatario;
        end
        else if (Copy(Linha, 7, 1)) = '2' then
        begin
          Transp.modFrete := mfContaTerceiros;
        end
        else
          Transp.modFrete := mfSemFrete;

        // CNPJ   showmessage(copy(linha,9,14)); /CPF showmessage(copy(linha,23,14));
        if ValidaCNPJ(Copy(Linha, 8, 14), False) = true then
          Transp.Transporta.CNPJCPF := (Copy(Linha, 8, 14))
        else if VALIDAcpf(Copy(Linha, 22, 14)) then
          Transp.Transporta.CNPJCPF := (Copy(Linha, 22, 14));

        Transp.Transporta.xNome := (Copy(Linha, 36, 60));
        // Razão social ou nome                     showmessage(copy(linha,36,60));
        Transp.Transporta.ie := (Copy(Linha, 96, 14));
        // IE                                                showmessage(copy(linha,96,14));
        Transp.Transporta.xEnder := (Copy(Linha, 110, 60));
        // Endereço completo                     showmessage(copy(linha,110,60));
        Transp.Transporta.xMun := (Copy(Linha, 170, 60));
        // Nome do Municipio                     showmessage(copy(linha,170,60));
        Transp.Transporta.uf := (Copy(Linha, 230, 2));
        // Sigla da UF                     showmessage(copy(linha,230,2));

        with Transp.Vol.Add do
        begin
          qVol := strtoint(Copy(Linha, 232, 15));
          // Quantidade de volume                           showmessage(copy(linha,232,15));
          esp := (Copy(Linha, 247, 60));
          // Especie dos volumes transportados                           showmessage(copy(linha,247,60));
          marca := (Copy(Linha, 307, 60));
          // Marca dos volumes transportados                           showmessage(copy(linha,307,60));
          pesoL := StrToFloat(StringReplace((Copy(Linha, 367, 15)), '.', ',',
            [rfReplaceAll]));
          // Peso Liquido (em Kg)                           showmessage(copy(linha,367,15));
          pesoB := StrToFloat(StringReplace((Copy(Linha, 382, 15)), '.', ',',
            [rfReplaceAll]));
          // Peso Bruto (em Kg)                           showmessage(copy(linha,382,15));
          nVol := (Copy(Linha, 397, 10));
        end;

        Transp.VeicTransp.placa := (Copy(Linha, 407, 7));
        Transp.VeicTransp.uf := (Copy(Linha, 414, 2));;
        // Transp.VeicTransp.rntc ='';

      end; // fim do EM0211

      // =============================EM1211=============================//
      if (Copy(Linha, 0, 6)) = 'EM1211' then
      begin
        // exporta.UFembarq := Copy(Linha, 7, 2);
        // exporta.xLocEmbarq := Copy(Linha, 9, 60);
        exporta.UFSaidaPais := Copy(Linha, 7, 2);
        exporta.xLocExporta := Copy(Linha, 9, 60);
        exporta.xLocDespacho := Copy(Linha, 69, 60);
      end;
      // Fim do EM1211
      // =============================EM0212=============================//
      if (Copy(Linha, 0, 6)) = 'EM0212' then
      begin
        nfat := Copy(Linha, 7, 60);
        vliq := StrToFloat(StringReplace((Copy(Linha, 97, 15)), '.', ',',
          [rfReplaceAll]));

        voriginal := StrToFloat(StringReplace((Copy(Linha, 67, 15)), '.', ',',
          [rfReplaceAll]));

        loja := Copy(Linha, 113, 4);
        if loja = 'loja' then
        begin
          with pag.Add do
          begin
            tPag := fpBoletoBancario;
            vPag := StrToFloat(StringReplace((Copy(Linha, 97, 15)), '.', ',',
              [rfReplaceAll]));

          end;
        end;

      end; // fim do EM0212

      // =============================EM0213=============================//

      if (Copy(Linha, 0, 6)) = 'EM0213' then
      begin

{$REGION 'Forma de Pagamento'}
        if (Ide.finNFe <> fnNormal) then
        begin
          with pag.Add do
          begin
            tPag := fpSemPagamento;
            vPag := 0;
          end;
        end
        else
        begin
          with pag.Add do
          begin
            tPag := fpDuplicataMercantil; // fpDuplicata
            vPag := StrToFloat(StringReplace((Copy(Linha, 77, 15)), '.', ',',
              [rfReplaceAll]));
          end;
        end;

{$ENDREGION}
        if ((strtodate(Copy(Linha, 75, 2) + '/' + Copy(Linha, 72, 2) + '/' +
          Copy(Linha, 67, 4)) > Date) or (strtoIntDef(numeroparcela, 0) > 1)) then
        begin
          with Cobr.Dup.Add do
          begin
            numeroparcela := (Copy(Linha, 7, 60));
            nDup := LeftZero((Copy(numeroparcela, pos('-', numeroparcela) + 1,
              length(numeroparcela))), 3);
            // Número da fatura
            dVenc := strtodate(Copy(Linha, 75, 2) + '/' + Copy(Linha, 72, 2) +
              '/' + Copy(Linha, 67, 4));
            // Data de vencimento
            vDup := StrToFloat(StringReplace((Copy(Linha, 77, 15)), '.', ',',
              [rfReplaceAll]));
            // Valor da duplicata

          end; // with
        end;

        Cobr.Fat.nfat := nfat;
        // Cobr.Fat.vOrig := voriginal + 0.001;
        Cobr.Fat.vOrig := voriginal;
        // Valor Original

        Cobr.Fat.vDesc := 0;

        Cobr.Fat.vliq := vliq;

      end; // fim do EM0213
      Readln(Arquivo, Linha); // Ler a proxima linha
    end; // enquanto não chegar ao fim do Arquivo
    // =============================EM0214=============================//
    if (Copy(Linha, 0, 6)) = 'EM0214' then
    begin

      InfAdic.infCpl := (Copy(Linha, 7, 2000));
    end; // fim do EM0214
  end; // Fim das configurações da NFe
  CloseFile(Arquivo);

  Result := true;
end;

function TForm1.EnviaNFe2DB(TipoEnvio: integer = 3): Boolean;
var
  Arquivo: TextFile;
  Linha: string;
  i, j, colunai, colunaf: integer;
  Cdi: integer;
  VSNprod, VSNAliq: Real;
  SQLCAB, SQLITEM, SQLDI, SQLDIETI, SQLEMP, SQLREF, SQLPARC: TSimpleDataSet;
  TipCndPgto, id_FinPaiCli: integer;
  EndTra: string;
  precoUnitario, valorIpi, valorMva, baseCalculo, precoComMVA,
    precoSemMva: Double;
begin
  Result := False;

  SQLCAB := TSimpleDataSet.Create(self);
  try

    SQLCAB.Connection := SQLConnection1;
    SQLCAB.ReadOnly := true;
    SQLCAB.DataSet.CommandText := SelectCabecalho('', '');

    SQLITEM := TSimpleDataSet.Create(self);
    try

      SQLITEM.Connection := SQLConnection1;
      SQLITEM.ReadOnly := true;
      SQLITEM.DataSet.CommandText := SelectItem('', '');

      SQLCAB.First;

      while not SQLCAB.Eof do
      begin

        with ACBrNFe1.NotasFiscais.Add.NFe do
        begin
          // Loop para ler todas as linhas do arquivo

          // =============================EM0201=============================//
          if (Copy(Linha, 0, 6)) = 'EM0201' then
          begin
            // Carrega Chave

            infNFe.ID := 'NFe' + SQLCAB.FieldByName('SEQNFE').AsString;

          end
          else

            // if (Copy(Linha, 0, 6)) = 'EM1201' then // Dados de contigencia
            if False then // DPEC
            begin
              // =============================EM1201=============================//
              Ide.dhCont := now;
              Ide.xJust := SQLCAB.FieldByName('JustDPEC').AsString;
            end;
          // fim do EM0201
          // =============================EM0202=============================//

          if true then
          begin

            SQLEMP := TSimpleDataSet.Create(self);

            try
              SQLEMP.Connection := SQLConnection1;

              SQLEMP.DataSet.CommandText := //
                ' Select GerEmp.ApeEmp,' + //
                ' GerEmp.NomEmp,' + //
                ' GerEmp.CgcEmp,' + //
                ' GerEmp.InsEmp,' + //
                ' GerEmp.CepEmp,' + //
                ' GerEmp.TenEmp,' + //
                ' GerEmp.EndEmp,' + //
                ' GerEmp.NumEmp,' + //
                ' GerEmp.RefEmp,' + //
                ' GerEmp.BaiEmp,' + //
                ' GerEmp.SigUfe,' + //
                ' GerEmp.PrtEmp,' + //
                ' GerEmp.FonEmp,' + //
                ' GerEmp.Id_FinUfe,' + //
                ' GerEmp.Id_FinCie,' + //
                ' GerEmp.Id_FinPai, ' + //
                ' GerEmp.TipEmp ' + //
                ' From GerEmp' + //
                ' Where GerEmp.CodEmp = ' + SQLCAB.FieldByName
                ('CODEMP').AsString;

              // Código da UF do emitente do documento fiscal'
              Ide.cUF := SQLEMP.FieldByName('Id_FinUfe').AsInteger;

              // Descrição da natureza de operação
              Ide.natOp := trim(SQLCAB.FieldByName('DesNat').AsString);

              TipCndPgto := 0;

              // Indicador da forma de pagamento 0-Pagamento à vista 1-Pagamento à prazo 2-Outros showmessage(copy(Linha,78,1));
              if SQLCAB.FieldByName('IntFin').AsString = 'Nao' then
                TipCndPgto := 2
              else
              begin
                if strtoint(BuscaSimples('FatPe3', 'QtdReg', '1 = 1',
                  SQLConnection1, //
                  ' Select Sum(FatPe3.PraPe3) as QtdReg' + //
                  ' From FatPe3' + //
                  ' Where FatPe3.CODEMP = ' + ((SQLCAB.FieldByName('CODEMP')
                  .AsString)) + //
                  ' and FatPe3.DTERES = ' +
                  QuotedStr(formatdatetime('mm/dd/yyyy',
                  SQLCAB.FieldByName('DTERES').AsDateTime)) + //
                  ' and FatPe3.NUMRES = ' + SQLCAB.FieldByName('NUMRES')
                  .AsString + //
                  ' and FatPe3.SEQLIB = ' + SQLCAB.FieldByName('SEQLIB')
                  .AsString + //
                  ' and FatPe3.SEQFAT = ' + SQLCAB.FieldByName('SEQFAT')
                  .AsString)) > 0 then
                  TipCndPgto := 1;

              end;

              if TipCndPgto = 0 then
                Ide.indPag := ipVista
              else if TipCndPgto = 1 then
                Ide.indPag := ipPrazo
              else
                Ide.indPag := ipOutras;

              /// ///////////////////////////////////////////////////

              // Código do Modelo do documento fiscal
              Ide.modelo := 55;
              // Série do documento fiscal
              Ide.serie := 1;
              // Número do documento fiscal
              Ide.nNF := SQLCAB.FieldByName('NroNfs').AsInteger;
              // numero do lote de envio
              VNumLote := SQLCAB.FieldByName('NroNfs').AsInteger;

              try
                // Data de emissão do documento fiscal
                Ide.dEmi := SQLCAB.FieldByName('DteFat').AsDateTime;
              except
                Memo1.Lines.Text := 'Problemas com a Data de Emissão da nota';
                Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno +
                  '\LogErro-' + IntToStr(VNumLote) + '.txt');
                Application.Terminate;
              end;

              // Data de saida ou entrada da Mercadoria/Produto
              // Ide.dSaiEnt := ;//Não informado

              // Tipo do documento fiscal
              if 1 = 1 then // Verificar Quando devolução e Outros Tipos
                Ide.tpNF := tnSaida
              else
                Ide.tpNF := tnEntrada;

              // Código do Municipio de Ocorrência do Fato Gerador
              Ide.cMunFG := SQLEMP.FieldByName('Id_EmpCie').AsInteger;

              // Formato de Impressao do DANFE
              if 1 = 1 then // Verificar para permitir escolha
                Ide.tpImp := tiRetrato
              else
                Ide.tpImp := tiPaisagem;

              // Forma de emissão da NF-e
              if TipoEnvio = 3 then
                Ide.tpEmis := teNormal
              else if TipoEnvio = 4 then
                Ide.tpEmis := teDPEC;

              // Digito verificador da Chave de Acesso da NF-e
              // Ide.cDV := strtoint(Copy(Linha, 121, 1));
              if trim(SQLCAB.FieldByName('SeqNFE').AsString) <> '' then
                Ide.cDV :=
                  strtoint(Copy(SQLCAB.FieldByName('SeqNFE').AsString, 44, 1));

              // Identificação do Ambiente
              if (VWebSAmbiente) = 1 then
                Ide.tpAmb := taHomologacao
              else
                Ide.tpAmb := taProducao;

              // Finalidade de emissão da NF-e
              { if (Copy(Linha, 123, 1)) = '1' then
                ide.finNFe := fnNormal; }

              // Finalidade de emissão da NF-e
              if uppercase(SQLCAB.FieldByName('MODPFA').AsString) = 'VENDAS'
              then
              begin
                Ide.finNFe := fnNormal;
              end
              else if uppercase(SQLCAB.FieldByName('MODPFA').AsString) = 'COMPLEMENTO'
              then
              // if (Copy(Linha, 123, 1)) = '2' then
              begin
                Ide.finNFe := fnComplementar;

                SQLREF := TSimpleDataSet.Create(self);
                SQLREF.Active := False;
                SQLREF.DataSet.CommandText :=
                  'select NFE_REF, CODUF, CNPJ, MODELO, SERIE, NRONFS from FATGER_REF WHERE ID_FATGER = '
                  + SQLCAB.FieldByName('ID').AsString;
                SQLREF.Active := true;

                if not SQLREF.IsEmpty then
                begin
                  SQLREF.First;
                  while not SQLREF.Eof do
                  begin

                    // Versão do processo de emissão da NF-e
                    with Ide.NFref.Add do
                    begin
                      RefNF.cUF := SQLREF.FieldByName('CODUF').AsInteger;
                      RefNF.AAMM :=
                        Copy(SQLREF.FieldByName('NFE_REF').AsString, 3, 4);
                      RefNF.CNPJ := SQLREF.FieldByName('CNPJ').AsString;
                      RefNF.modelo := SQLREF.FieldByName('MODELO').AsInteger;
                      RefNF.serie := SQLREF.FieldByName('SERIE').AsInteger;
                      RefNF.nNF := SQLREF.FieldByName('NRONFS').AsInteger;
                      refNFe := SQLREF.FieldByName('NFE_REF').AsString;

                    end;

                    SQLREF.Next;
                  end;
                end;
              end
              else if uppercase(SQLCAB.FieldByName('MODPFA').AsString) = 'DEVOLUCAO'
              then
              begin
                Ide.finNFe := fnDevolucao;

                SQLREF := TSimpleDataSet.Create(self);
                SQLREF.Active := False;
                SQLREF.DataSet.CommandText :=
                  'select NFE_REF, CODUF, CNPJ, MODELO, SERIE, NRONFS from FATGER_REF WHERE ID_FATGER = '
                  + SQLCAB.FieldByName('ID').AsString;
                SQLREF.Active := true;

                if not SQLREF.IsEmpty then
                begin
                  SQLREF.First;
                  while not SQLREF.Eof do
                  begin

                    // Versão do processo de emissão da NF-e
                    with Ide.NFref.Add do
                    begin
                      RefNF.cUF := SQLREF.FieldByName('CODUF').AsInteger;
                      RefNF.AAMM :=
                        Copy(SQLREF.FieldByName('NFE_REF').AsString, 3, 4);
                      RefNF.CNPJ := SQLREF.FieldByName('CNPJ').AsString;
                      RefNF.modelo := SQLREF.FieldByName('MODELO').AsInteger;
                      RefNF.serie := SQLREF.FieldByName('SERIE').AsInteger;
                      RefNF.nNF := SQLREF.FieldByName('NRONFS').AsInteger;
                      refNFe := SQLREF.FieldByName('NFE_REF').AsString;

                    end;

                    SQLREF.Next;
                  end;
                end;
              end
              else if uppercase(SQLCAB.FieldByName('MODPFA').AsString) = 'AJUSTE'
              then
              begin
                Ide.finNFe := fnAjuste;
              end;

              Ide.procEmi := peAplicativoContribuinte;

              // Processo de emissão da NF-e
              Ide.verProc := 'Emerion Faturamento NFeDB';



              // =============================EM0203=============================//

              // Razão social obrigatória para ambiente de homologação
              if VWebSAmbiente = 1 then
                strxNome :=
                  'NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL'
              else
                strxNome := SQLEMP.FieldByName('NOMEMP').AsString;

              // CNPJ do emitente  ou  CPF do emitente
              if ValidaCNPJ(SQLEMP.FieldByName('CgcEmp').AsString, False) = true
              then
                emit.CNPJCPF := trim(SQLEMP.FieldByName('CgcEmp').AsString)
              else
                emit.CNPJCPF := trim(SQLEMP.FieldByName('CgcEmp').AsString);

              // Razão social ou Nome do emitente
              emit.xNome := trim(strxNome); // copy(Linha, 35, 60);
              // Nome fantasia
              emit.xFant := SQLEMP.FieldByName('ApeEmp').AsString;
              // Logradouro
              emit.EnderEmit.xLgr :=
                (SQLEMP.FieldByName('TenEmp').AsString + ' ' +
                SQLEMP.FieldByName('EndEmp').AsString);
              // Número
              emit.EnderEmit.nro := SQLEMP.FieldByName('NumEmp').AsString;
              // Complemento
              emit.EnderEmit.xCpl := SQLEMP.FieldByName('RefEmp').AsString;
              // Bairro
              emit.EnderEmit.xBairro := SQLEMP.FieldByName('BaiEmp').AsString;
              // Código do municipio
              emit.EnderEmit.cMun := SQLEMP.FieldByName('Id_EmpCie').AsInteger;
              // Nome do municipio
              emit.EnderEmit.xMun := SQLEMP.FieldByName('CidEmp').AsString;
              // Sigla da UF
              emit.EnderEmit.uf := SQLEMP.FieldByName('UfeEmp').AsString;
              ACBrNFe1.Configuracoes.WebServices.uf :=
                SQLEMP.FieldByName('UfeEmp').AsString;
              // Código do CEP
              emit.EnderEmit.CEP := SQLEMP.FieldByName('CepEmp').AsInteger;
              // Código do País
              emit.EnderEmit.cPais := SQLEMP.FieldByName('NroPais_Emp')
                .AsInteger;
              // Brasil
              emit.EnderEmit.xPais := SQLEMP.FieldByName('NomPais_Emp')
                .AsString;
              // Telefone
              emit.EnderEmit.fone := SQLEMP.FieldByName('FonEmp').AsString;
              // IE
              emit.ie := SQLEMP.FieldByName('InsEmp').AsString;
              // IEST
              emit.IEST := SQLCAB.FieldByName('INSSUB').AsString;

              if SQLEMP.FieldByName('TipEmp').AsString = 'SimplesNacional' then
                emit.CRT := crtSimplesNacional;

            finally
              FreeAndnil(SQLEMP);
            end;
            // fim do EM0203
            // =============================EM0204=============================//

            // CNPJ do emitente  ou  CPF do emitente
            if ((ValidaCNPJ(SQLCAB.FieldByName('CGCCLI').AsString, False)
              = true) and (trim(SQLCAB.FieldByName('CGCCLI').AsString) <> ''))
            then //
            begin
              dest.CNPJCPF := SQLCAB.FieldByName('CGCCLI').AsString;
            end
            else
            begin
              dest.CNPJCPF := SQLCAB.FieldByName('CGCCLI').AsString;
            end;

            // Razão social ou nome do destinatario
            dest.xNome := SQLCAB.FieldByName('NomCli').AsString;
            // Logradouro
            dest.EnderDest.xLgr := trim(SQLCAB.FieldByName('TefCli').AsString +
              ' ' + SQLCAB.FieldByName('EnfCli').AsString);
            // Número
            dest.EnderDest.nro := trim(SQLCAB.FieldByName('NrfCli').AsString);
            // Complemento
            dest.EnderDest.xCpl := trim(SQLCAB.FieldByName('RffCli').AsString);
            // Bairro
            dest.EnderDest.xBairro :=
              trim(SQLCAB.FieldByName('BafCli').AsString);
            // Código do Municipio
            id_FinPaiCli := //
              strtoint(BuscaSimples('FinCie', 'SigNfe', 'Id_FinCie = ' + //
              SQLCAB.FieldByName('Id_FinCif').AsString, SQLConnection1));
            dest.EnderDest.cMun := id_FinPaiCli; //

            // Nome do Municipio
            dest.EnderDest.xMun := //
              (BuscaSimples('FinCie', 'NomCie', 'Id_FinCie = ' + //
              SQLCAB.FieldByName('Id_FinCif').AsString, SQLConnection1));
            dest.EnderDest.cMun := id_FinPaiCli; //

            // Sigla da UF_
            dest.EnderDest.uf := (SQLCAB.FieldByName('UfeCli').AsString);
            // Código do CepN
            dest.EnderDest.CEP := SQLCAB.FieldByName('CefCli').AsInteger;

            // Código do País
            dest.EnderDest.cPais := //
              strtoint(BuscaSimples('FinCli', 'Id_FinPai', ' CodCli = ' + //
              SQLCAB.FieldByName('CodCli').AsString, SQLConnection1));
            // Brasil
            dest.EnderDest.xPais := //
              trim(BuscaSimples('FinPai', 'NomPai',
              'Id_FinPai = ' + IntToStr(id_FinPaiCli), SQLConnection1));

            // TelefoneE
            dest.EnderDest.fone := trim(SQLCAB.FieldByName('TelCli').AsString);
            // trim(BuscaSimples('FinCli', 'TEL', ' 1 = 1 ', SQLConnection1),
            // ' Select Pt1Cli||Fo1Cli TEL from fincli where codcli = ' + SQLCAB.FieldByName('CodCli').AsString);

            // IE
            dest.ie := trim(BuscaSimples('FinCli', 'InsCli', ' CodCli = ' + //
              SQLCAB.FieldByName('CodCli').AsString, SQLConnection1));
            // Inscrição SUFRAMA}
            dest.ISUF := trim(BuscaSimples('FinCli', 'NroSuf', ' CodCli = ' +
              //
              SQLCAB.FieldByName('CodCli').AsString, SQLConnection1));

            // ==============================EM0205=============================//

            // CNPJ do emitente  ou  CPF do Entrega
            if ((ValidaCNPJ(SQLCAB.FieldByName('CGCCLI').AsString, False)
              = true) and (trim(SQLCAB.FieldByName('CGCCLI').AsString) <> ''))
            then //
            begin
              Entrega.CNPJCPF := SQLCAB.FieldByName('CGCCLI').AsString;
            end
            else
            begin
              Entrega.CNPJCPF := SQLCAB.FieldByName('CGCCLI').AsString;
            end;

            Entrega.xLgr := trim(SQLCAB.FieldByName('TefCli').AsString + ' ' +
              SQLCAB.FieldByName('TefCli').AsString);
            Entrega.nro := SQLCAB.FieldByName('NumCli').AsString;
            Entrega.xCpl := SQLCAB.FieldByName('RefCli').AsString;
            Entrega.xBairro := SQLCAB.FieldByName('BaiCli').AsString;

            Entrega.cMun := //
              strtoint(BuscaSimples('FinCie', 'SigNfe', 'Id_FinCie = ' + //
              SQLCAB.FieldByName('Id_FinCie').AsString, SQLConnection1));
            dest.EnderDest.cMun := id_FinPaiCli; // ;
            Entrega.xMun := //
              BuscaSimples('FinCie', 'NomCie', 'Id_FinCie = ' + //
              SQLCAB.FieldByName('Id_FinCie').AsString, SQLConnection1);
            dest.EnderDest.cMun := id_FinPaiCli; // ;

            Entrega.uf := SQLCAB.FieldByName('UfeCli').AsString;


            // =============================EM0206=============================//

            with Det.Add do
            begin
              // Nro. do item
              Prod.nItem := SQLCAB.FieldByName('NroPe2').AsInteger;
              // Código do Produto ou serviço
              Prod.cProd := trim(SQLCAB.FieldByName('CodClp').AsString +
                SQLCAB.FieldByName('CodGru').AsString +
                SQLCAB.FieldByName('CodSub').AsString +
                SQLCAB.FieldByName('CodPro').AsString);
              // cEAN!
              Prod.cEAN := trim(SQLCAB.FieldByName('cEAN').AsString);
              // Descrição do produto ou serviço
              Prod.xProd := trim(SQLCAB.FieldByName('DesPe2').AsString);
              // Código NCM
              Prod.NCM := trim(SQLCAB.FieldByName('ClsIpi').AsString);
              // EX_TIPI
              Prod.EXTIPI := '';
              // Código fiscal da operação
              Prod.CFOP := trim(SQLCAB.FieldByName('CodCfo').AsString);
              // Unidade comercial
              Prod.uCom := ifthen(trim(SQLCAB.FieldByName('CodUnd').AsString) <>
                '', trim(SQLCAB.FieldByName('CodUnd').AsString), 'UN');
              // Quantidade comercial
              Prod.qCom := SQLCAB.FieldByName('QtpPe2').AsFloat;
              // Valor unitário de comercialização
              Prod.vUnCom := SQLCAB.FieldByName('VluPe2').AsFloat;
              // Valor Total Bruto dos Produtos ou Serviços
              Prod.vProd := SQLCAB.FieldByName('TotPe2').AsFloat;
              // cEANTrib
              Prod.cEANTrib := trim(SQLCAB.FieldByName('cEANTrib').AsString);
              // Unidade Tributavel
              Prod.uTrib := ifthen(trim(SQLCAB.FieldByName('CodUnd').AsString)
                <> '', trim(SQLCAB.FieldByName('CodUnd').AsString), 'UND');
              // UQuantidade Tributavel
              Prod.qTrib := Prod.qCom;
              // := SQLCAB.FieldByName('QtpPe2').AsFloat;
              // Valor Unitário de tributação
              Prod.vUnTrib := SQLCAB.FieldByName('VluPe2').AsFloat;
              // Valor Total do Frete
              Prod.vFrete := SQLCAB.FieldByName('TotFrt').AsFloat;
              // Valor Total do Seguro
              Prod.vSeg := SQLCAB.FieldByName('TotSeg').AsFloat;
              // Valor do outras despesas acessorias
              Prod.vOutro := SQLCAB.FieldByName('TotDes').AsFloat;
              // Valor do outras Desconto
              Prod.vDesc := SQLCAB.FieldByName('TotDsr').AsFloat;

              /// //// Imposto de importação

              if length(Linha) > 385 then
              begin
                // Valor da Base de Calculo para II
                Imposto.II.vBC := SQLCAB.FieldByName('VLRBCII').AsFloat;
                // Valor do Imposto de importação
                Imposto.II.vII := SQLCAB.FieldByName('VLRIMPII').AsFloat;
                // Valor de Despesas Aduaneiras
                if trim(Copy(Linha, 414, 15)) <> '' then
                  Imposto.II.vDespAdu :=
                    SQLCAB.FieldByName('VLRDESPATU').AsFloat;
                // Valor do IOF
                if trim(Copy(Linha, 429, 15)) <> '' then
                  Imposto.II.vIOF := SQLCAB.FieldByName('VLRIOF').AsFloat;

              end;

              // PEDIDO DE COMPRA
              if trim(SQLCAB.FieldByName('NUMPEDCOMPRA').AsString) <> '' then
              begin
                if trim(Copy(Linha, 444, 15)) <> '' then
                  Prod.xPed := trim(SQLCAB.FieldByName('NUMPEDCOMPRA')
                    .AsString);

                if trim(Copy(Linha, 459, 6)) <> '' then
                  Prod.nItemPed := SQLITEM.FieldByName('NUMITEMCOMPRA')
                    .AsString;
              end;

              // Ler aProxima linha pra chamar o 207
              Readln(Arquivo, Linha); // Ler a proxima linha
              // =============================EM1206=============================//
              if (Copy(Linha, 0, 6)) = 'EM1206' then
              begin
                infAdProd := trim(SQLITEM.FieldByName('DESIMP').AsString);
              end;
              // fim do EM1206
              // Ler aProxima linha pra chamar o 207
              Readln(Arquivo, Linha); // Ler a proxima linha
              // =============================EM1207  DI=============================//

              if Prod.CFOP[1] = '3' then
              begin
                SQLDI := TSimpleDataSet.Create(self);
                try
                  SQLDI.Connection := SQLConnection1;
                  SQLDI.ReadOnly := true;
                  SQLDI.DataSet.CommandText := //
                    ' Select ID_DI, ID_CMPNF2, NUMDI, DATADI, LOCALDESEMB, UFDESEMB, DATADESEMB, CODEXPORT, TAB_ORIGEM, ID_ITEM_ORIGEM '
                    +
                  //
                    ' From DI ' + //
                    ' Where ID_CMPNF2 = ' + SQLITEM.FieldByName
                    ('ID_FATPE2').AsString;
                  SQLDI.Active := true;

                  SQLDIETI := TSimpleDataSet.Create(self);
                  try
                    SQLDIETI.Connection := SQLConnection1;
                    SQLDIETI.ReadOnly := true;

                    if not SQLDI.IsEmpty then
                    begin
                      SQLDI.First;
                      while not SQLDI.Eof do
                      begin
                        with Prod.DI.Add do
                        begin
                          nDi := SQLDI.FieldByName('numdi').AsString;
                          dDi := SQLDI.FieldByName('DATADI').AsDateTime;
                          xLocDesemb :=
                            SQLDI.FieldByName('LOCALDESEMB').AsString;
                          UFDesemb := SQLDI.FieldByName('UFDESEMB').AsString;
                          dDesemb := SQLDI.FieldByName('DATADESEMB').AsDateTime;
                          cExportador := SQLDI.FieldByName('CODEXPORT')
                            .AsString;

                          SQLDIETI.DataSet.CommandText := //
                            ' Select ID_DIDET, ID_DI, NSEQADIC, CODFAB, VDESCDI, QTDE, NADICAO '
                            +
                          //
                            ' From DIDET ' + //
                            ' WHERE ID_DI = ' +
                            SQLDI.FieldByName('ID_DI').AsString;
                          SQLDIETI.Active := true;
                          SQLDIETI.First;
                          // Ler a proxima linha
                          while not SQLDIETI.Eof do
                          begin
                            // Cdi := 1;

                            with adi.Add do
                            begin
                              nSeqAdi := SQLDI.FieldByName('NSEQADIC')
                                .AsInteger;
                              nAdicao := SQLDI.FieldByName('NADICAO').AsInteger;
                              cFabricante :=
                                SQLDI.FieldByName('CODFAB').AsString;
                              vDescDI := SQLDI.FieldByName('VDESCDI').AsFloat;
                            end;

                            SQLDIETI.Next;
                          end;
                        end;
                        SQLDI.Next;
                      end;
                    end;
                  Finally
                    FreeAndnil(SQLDIETI);
                  end;
                Finally
                  FreeAndnil(SQLDI);
                end;

              end;

              // 'EM0207'

              if not true then // if not SN then
              begin
                case SQLITEM.FieldByName('CODST1').AsInteger of
                  0:
                    Imposto.ICMS.orig := oeNacional; // Origem da mercadoria
                  1:
                    Imposto.ICMS.orig := oeEstrangeiraImportacaoDireta;
                  2:
                    Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasil;
                  3:
                    Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior40;
                  // Origem da mercadoria
                  4:
                    Imposto.ICMS.orig := oeNacionalProcessosBasicos;
                  5:
                    Imposto.ICMS.orig :=
                      oeNacionalConteudoImportacaoInferiorIgual40;
                  6:
                    Imposto.ICMS.orig :=
                      oeEstrangeiraImportacaoDiretaSemSimilar;
                  7:
                    Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasilSemSimilar;
                  8:
                    Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior70;
                end;

                case SQLITEM.FieldByName('CODST2').AsInteger of
                  00:
                    Imposto.ICMS.CST := cst00;
                  10:
                    Imposto.ICMS.CST := cst10;
                  20:
                    Imposto.ICMS.CST := cst20;
                  30:
                    Imposto.ICMS.CST := cst30;
                  40:
                    Imposto.ICMS.CST := cst40;
                  41:
                    Imposto.ICMS.CST := cst41;
                  45:
                    Imposto.ICMS.CST := cst45;
                  50:
                    Imposto.ICMS.CST := cst50;
                  51:
                    Imposto.ICMS.CST := cst51;
                  60:
                    Imposto.ICMS.CST := cst60;
                  70:
                    Imposto.ICMS.CST := cst70;
                  80:
                    Imposto.ICMS.CST := cst80;
                  81:
                    Imposto.ICMS.CST := cst81;
                  90:
                    Imposto.ICMS.CST := cst90;

                end;

                // Modalidade de determinação da BC do ICMS ST
                Imposto.ICMS.modBCST := dbisPrecoTabelado;
                // Percential de redução de BC do ICMS
                Imposto.ICMS.pRedBC := SQLITEM.FieldByName('REDICM').AsFloat;
                // Valor da BC do ICMS
                Imposto.ICMS.vBC := SQLITEM.FieldByName('BASICM').AsFloat;
                // Aliquota do imposto
                Imposto.ICMS.pICMS := SQLITEM.FieldByName('ICMPE2').AsFloat;
                // Valor do ICMSN
                Imposto.ICMS.vICMS := SQLITEM.FieldByName('TOTICM').AsFloat;
                // Valor da BC do ICMS ST
                Imposto.ICMS.vBCST := SQLITEM.FieldByName('BASSUB').AsFloat;
                // Aliquota do imposto do ICMS ST 5
                Imposto.ICMS.pICMSST := SQLITEM.FieldByName('ICMSUB').AsFloat;
                // Percentual da Margem de valor Adicionado do ICMS ST 5
                Imposto.ICMS.pMVAST := SQLITEM.FieldByName('MRGSUB').AsFloat;
                // Valor do ICMS ST 15
                Imposto.ICMS.vICMSST := SQLITEM.FieldByName('TOTSUB').AsFloat;

                // Imposto.ICMS.modBC := dbiValorOperacao;

              end
              else
              begin
                case SQLITEM.FieldByName('CODST1').AsInteger of
                  0:
                    Imposto.ICMS.orig := oeNacional;
                  // Origem da mercadoria
                  1:
                    Imposto.ICMS.orig := oeEstrangeiraImportacaoDireta;
                  2:
                    Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasil;
                  3:
                    Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior40;
                  // Origem da mercadoria
                  4:
                    Imposto.ICMS.orig := oeNacionalProcessosBasicos;
                  5:
                    Imposto.ICMS.orig :=
                      oeNacionalConteudoImportacaoInferiorIgual40;
                  6:
                    Imposto.ICMS.orig :=
                      oeEstrangeiraImportacaoDiretaSemSimilar;
                  7:
                    Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasilSemSimilar;
                end;

                case SQLITEM.FieldByName('CODST2').AsInteger of

                  101:
                    Imposto.ICMS.CSOSN := csosn101;
                  102:
                    Imposto.ICMS.CSOSN := csosn102;
                  103:
                    Imposto.ICMS.CSOSN := csosn103;
                  201:
                    Imposto.ICMS.CSOSN := csosn201;
                  202:
                    Imposto.ICMS.CSOSN := csosn202;
                  203:
                    Imposto.ICMS.CSOSN := csosn203;
                  300:
                    Imposto.ICMS.CSOSN := csosn300;
                  400:
                    Imposto.ICMS.CSOSN := csosn400;
                  500:
                    Imposto.ICMS.CSOSN := csosn500;
                  900:
                    Imposto.ICMS.CSOSN := csosn900;
                end;

                // Modalidade de determinação da BC do ICMS STO
                Imposto.ICMS.modBCST := dbisPrecoTabelado;
                // Percential de redução de BC do ICMSL
                Imposto.ICMS.pRedBC := SQLITEM.FieldByName('REDICM').AsFloat;
                // Valor da BC do ICMS
                Imposto.ICMS.vBC := SQLITEM.FieldByName('BASICM').AsFloat;
                // Aliquota do imposto
                // --simples--
                Imposto.ICMS.pICMS := SQLITEM.FieldByName('ICMPE2').AsFloat;
                // Aliquota do imposto
                Imposto.ICMS.pCredSN := SQLITEM.FieldByName('ICMPE2').AsFloat;
                // Aliquota do imposto
                VSNAliq := SQLITEM.FieldByName('ICMPE2').AsFloat;
                // Valor do ICMS
                // --simples--
                Imposto.ICMS.vICMS := SQLITEM.FieldByName('TOTICM').AsFloat;

                // Imposto.ICMS.vCredICMSSN := VSNprod * (VSNAliq / 100);
                Imposto.ICMS.vCredICMSSN :=
                  SQLITEM.FieldByName('TOTICM').AsFloat;

                if Copy(Linha, 13, 3) <> '500' then
                begin
                  // Aliquota do imposto
                  Imposto.ICMS.vBCST := SQLITEM.FieldByName('BASSUB').AsFloat;
                  // Valor da BC do ICMS ST
                  Imposto.ICMS.pICMSST := SQLITEM.FieldByName('ICMSUB').AsFloat;
                  // Aliquota do imposto do ICMS ST 5
                  Imposto.ICMS.pMVAST := SQLITEM.FieldByName('MRGSUB').AsFloat;
                  // Percentual da Margem de valor Adicionado do ICMS ST 5
                  Imposto.ICMS.vICMSST := SQLITEM.FieldByName('TOTSUB').AsFloat;

                end
                else
                begin
                  Imposto.ICMS.vBCSTRet :=
                    SQLITEM.FieldByName('BASSUB').AsFloat;
                  // Valor da BC do ICMS ST
                  Imposto.ICMS.pICMSST := SQLITEM.FieldByName('ICMSUB').AsFloat;
                  // Aliquota do imposto do ICMS ST 5
                  Imposto.ICMS.pMVAST := SQLITEM.FieldByName('MRGSUB').AsFloat;
                  // Percentual da Margem de valor Adicionado do ICMS ST 5
                  Imposto.ICMS.vICMSSTRet :=
                    SQLITEM.FieldByName('TOTSUB').AsFloat;

                  if (SQLITEM.FieldByName('TOTSUB').AsFloat > 0) then
                  begin
                    precoUnitario := SQLITEM.FieldByName('VPFITE').AsFloat;
                    valorIpi := SQLITEM.FieldByName('TOTIPI').AsFloat;
                    valorMva := SQLITEM.FieldByName('MRGSUB').AsFloat;
                    baseCalculo := (precoUnitario + valorIpi) *
                      (1 + valorMva / 100);
                    precoComMVA :=
                      ((baseCalculo * SQLCAB.FieldByName('QtpPe2').AsFloat) *
                      SQLITEM.FieldByName('ICMSUB').AsFloat) / 100;
                    precoSemMva :=
                      (((precoUnitario + valorIpi) *
                      SQLCAB.FieldByName('QtpPe2').AsFloat) *
                      SQLITEM.FieldByName('ICMSUB').AsFloat) / 100;

                    Imposto.ICMS.vBCSTRet :=
                      SQLITEM.FieldByName('BASSUB').AsFloat;
                    Imposto.ICMS.vICMSSubstituto := precoSemMva;
                    Imposto.ICMS.vICMSSTRet := precoComMVA - precoSemMva;
                    Imposto.ICMS.vICMS := SQLITEM.FieldByName('ICMSUB').AsFloat;
                  end;
                end;

              end;


              // =============================EM0208=============================//

              Imposto.IPI.vBC := SQLITEM.FieldByName('BASIPI').AsFloat;
              // 15 Valor da BC do IPI
              Imposto.IPI.pIPI := SQLITEM.FieldByName('IPIPE2').AsFloat;
              // 5 Aliquota do imposto
              Imposto.IPI.vIPI := SQLITEM.FieldByName('TOTIPI').AsFloat;
              // 15 Valor do IPI
              // CASE   Situação tributária do IPI
              case SQLITEM.FieldByName('CSTIPI').AsInteger of
                00:
                  Imposto.IPI.CST := ipi00;
                49:
                  Imposto.IPI.CST := ipi49;
                50:
                  Imposto.IPI.CST := ipi50;
                99:
                  Imposto.IPI.CST := ipi99;
                01:
                  Imposto.IPI.CST := ipi01;
                02:
                  Imposto.IPI.CST := ipi02;
                03:
                  Imposto.IPI.CST := ipi03;
                04:
                  Imposto.IPI.CST := ipi04;
                05:
                  Imposto.IPI.CST := ipi05;
                51:
                  Imposto.IPI.CST := ipi51;
                52:
                  Imposto.IPI.CST := ipi52;
                53:
                  Imposto.IPI.CST := ipi53;
                54:
                  Imposto.IPI.CST := ipi54;
                55:
                  Imposto.IPI.CST := ipi55;
              end;

              // =============================EM0209=============================//

              // CASEEEEE 02 Situação Tributaria do PIS       showmessage(copy(EM0209,12,2));
              case SQLITEM.FieldByName('CDSPIS').AsInteger of
                01:
                  Imposto.PIS.CST := pis01;
                02:
                  Imposto.PIS.CST := pis02;
                03:
                  Imposto.PIS.CST := pis03;
                04:
                  Imposto.PIS.CST := pis04;
                06:
                  Imposto.PIS.CST := pis06;
                07:
                  Imposto.PIS.CST := pis07;
                08:
                  Imposto.PIS.CST := pis08;
                09:
                  Imposto.PIS.CST := pis09;
                99:
                  Imposto.PIS.CST := pis99;
              end;
              // 15 BC PIS
              Imposto.PIS.vBC := SQLITEM.FieldByName('BASPIS').AsFloat;
              // 5 Percentual do PIS
              Imposto.PIS.pPIS := SQLITEM.FieldByName('ALIQPIS').AsFloat;
              // 15 Valor do PIS
              Imposto.PIS.vPIS := SQLITEM.FieldByName('TOTPIS').AsFloat;

              // CASE Situação Tributaria do COFINS       showmessage(copy(EM0209,52,2));
              case SQLITEM.FieldByName('CDSCOF').AsInteger of
                01:
                  Imposto.COFINS.CST := cof01;
                02:
                  Imposto.COFINS.CST := cof02;
                03:
                  Imposto.COFINS.CST := cof03;
                04:
                  Imposto.COFINS.CST := cof04;
                06:
                  Imposto.COFINS.CST := cof06;
                07:
                  Imposto.COFINS.CST := cof07;
                08:
                  Imposto.COFINS.CST := cof08;
                09:
                  Imposto.COFINS.CST := cof09;
                99:
                  Imposto.COFINS.CST := cof99;
              end;

              // 15 BC COFINS
              Imposto.COFINS.vBC := SQLITEM.FieldByName('BASCOF').AsFloat;
              // 5 Percentual do COFINS
              Imposto.COFINS.pCOFINS := SQLITEM.FieldByName('ALIQCOF').AsFloat;
              // 15 Valor do COFINS
              Imposto.COFINS.vCOFINS := SQLITEM.FieldByName('TOTCOF').AsFloat;

              SQLITEM.Next;
            end;
          end; // with do Add.item


          // =============================EM0210=============================//

          // Base de Calculo do ICMS
          Total.ICMSTot.vBC := SQLCAB.FieldByName('BASICM').AsFloat;
          // Valor Total do ICMS
          Total.ICMSTot.vICMS := SQLCAB.FieldByName('TOTICM').AsFloat;
          // Base de Calculo do ICMS ST   showmessage(copy(linha,38,15));
          Total.ICMSTot.vBCST := SQLCAB.FieldByName('BASSUB').AsFloat;
          // Valor Total do ICMS ST  showmessage(copy(linha,53,15));
          Total.ICMSTot.vST := SQLCAB.FieldByName('TOTSUB').AsFloat;
          // Valor Total dos produtos e serviços  showmessage(copy(linha,68,15));
          Total.ICMSTot.vProd := SQLCAB.FieldByName('TOTFAT').AsFloat;
          // Valor Total do Frete       showmessage(copy(linha,83,15));
          Total.ICMSTot.vFrete := SQLCAB.FieldByName('TOTFRT').AsFloat;
          // Valor Total do Seguro      showmessage(copy(linha,98,15));
          Total.ICMSTot.vSeg := SQLCAB.FieldByName('TOTSEG').AsFloat;
          // Valor Total do Desconto     showmessage(copy(linha,113,15));
          Total.ICMSTot.vDesc := SQLCAB.FieldByName('TOTDSR').AsFloat;
          // Valor Total do II           showmessage(copy(linha,128,15));
          Total.ICMSTot.vII := SQLCAB.FieldByName('TotImpII').AsFloat;
          // Valor Total do IPI        showmessage(copy(linha,143,15));
          Total.ICMSTot.vIPI := SQLCAB.FieldByName('TOTIPI').AsFloat;
          // Valor Total do PIS         showmessage(copy(linha,158,15));
          Total.ICMSTot.vPIS := SQLCAB.FieldByName('TOTPIS').AsFloat;
          // Valor Total do COFINS      showmessage(copy(linha,173,15));
          Total.ICMSTot.vCOFINS := SQLCAB.FieldByName('TOTCOF').AsFloat;
          // Outras Despesas Acessórias showmessage(copy(linha,188,15));
          Total.ICMSTot.vOutro := SQLCAB.FieldByName('TOTDES').AsFloat;
          // Valor Total da NFe       showmessage(copy(linha,203,15));
          Total.ICMSTot.vNF := SQLCAB.FieldByName('TOTCOF').AsFloat;
          // Valor Total da NFe       showmessage(copy(linha,218,15));
          Total.ICMSTot.vNF := SQLCAB.FieldByName('TOTCOF').AsFloat;



          // =============================EM0211=============================//

          case SQLCAB.FieldByName('ID_FRETE').AsInteger of
            0:
              Transp.modFrete := mfContaEmitente;
            1:
              Transp.modFrete := mfContaDestinatario;
            2:
              Transp.modFrete := mfContaTerceiros;
          else
            Transp.modFrete := mfSemFrete;

          end;

          // CNPJ   showmessage(copy(linha,9,14)); /CPF showmessage(copy(linha,23,14));
          if true then

            if length(trim(SQLCAB.FieldByName('CGCTRA').AsString)) < 14 then
            begin
              if ValidaCNPJ(trim(SQLCAB.FieldByName('CGCTRA').AsString)) then
                Transp.Transporta.CNPJCPF :=
                  trim(SQLCAB.FieldByName('CGCTRA').AsString);
            end
            else
            begin
              if VALIDAcpf(trim(SQLCAB.FieldByName('CGCTRA').AsString)) then
                Transp.Transporta.CNPJCPF :=
                  trim(SQLCAB.FieldByName('CGCTRA').AsString);
            end;

          if trim(SQLCAB.FieldByName('TenTra').AsString) <> '' then
            EndTra := trim(SQLCAB.FieldByName('TenTra').AsString) + '. ' +
              SQLCAB.FieldByName('EndTra').AsString + ', ' +
              trim(SQLCAB.FieldByName('numtra').AsString)
          else
            EndTra := SQLCAB.FieldByName('EndTra').AsString + ', ' +
              trim(SQLCAB.FieldByName('numtra').AsString);

          Transp.Transporta.xNome :=
            (Copy(trim(SQLCAB.FieldByName('NOMTRA').AsString), 1, 60));
          // Razão social ou nome                     showmessage(copy(linha,36,60));
          Transp.Transporta.ie :=
            (Copy(trim(SQLCAB.FieldByName('InsTra').AsString), 1, 14));
          // IE                                                showmessage(copy(linha,96,14));
          Transp.Transporta.xEnder :=
            (Copy(trim(SQLCAB.FieldByName('RefTra').AsString), 1, 60));
          // Endereço completo                     showmessage(copy(linha,110,60));
          Transp.Transporta.xMun :=
            (Copy(trim(SQLCAB.FieldByName('NumTra').AsString), 1, 60));
          // Nome do Municipio                     showmessage(copy(linha,170,60));
          Transp.Transporta.uf :=
            (Copy(trim(SQLCAB.FieldByName('UfeTra').AsString), 1, 2));
          // Sigla da UF                     showmessage(copy(linha,230,2));

          with Transp.Vol.Add do
          begin
            qVol := SQLCAB.FieldByName('AltVol').AsInteger;
            // Quantidade de volume                           showmessage(copy(linha,232,15));
            esp := (Copy(trim(SQLCAB.FieldByName('EspFat').AsString), 1, 60));
            // Especie dos volumes transportados                           showmessage(copy(linha,247,60));
            marca := (Copy(trim(SQLCAB.FieldByName('MarFat').AsString), 1, 60));
            // Marca dos volumes transportados                           showmessage(copy(linha,307,60));
            pesoL := SQLCAB.FieldByName('InfLiq').AsFloat;
            // Peso Liquido (em Kg)                           showmessage(copy(linha,367,15));
            pesoB := SQLCAB.FieldByName('InfBrt').AsFloat;
            // Peso Bruto (em Kg)                           showmessage(copy(linha,382,15));
            nVol := (Copy(trim(SQLCAB.FieldByName('NROFAT').AsString), 1, 10));
          end;
          // end; // fim do EM0211

          // =============================EM1211=============================//
          if SQLCAB.FieldByName('UFECLI').AsString = 'EX' then
          begin
            exporta.UFembarq :=
              Copy(trim(SQLCAB.FieldByName('UFEMB').AsString), 1, 2);
            exporta.xLocEmbarq :=
              Copy(trim(SQLCAB.FieldByName('LOCEMB').AsString), 1, 60);
          end;
          // Fim do EM1211
          // =============================EM0212=============================//
          if (Copy(Linha, 0, 6)) = 'EM0212' then
          begin

            Cobr.Fat.nfat := trim(SQLCAB.FieldByName('NroNfs').AsString);
            Cobr.Fat.vOrig := SQLCAB.FieldByName('TOTGER').AsFloat;
            // Valor Original showmessage(copy(linha,67,15));
            Cobr.Fat.vDesc := SQLCAB.FieldByName('TOTDESCINC').AsFloat;
            // Valor do desconto showmessage(copy(linha,82,15));
            Cobr.Fat.vliq := SQLCAB.FieldByName('TOTGER').AsFloat;
            // Valor Original showmessage(copy(linha,97,15));
          end; // fim do EM0212
          // =============================EM0213=============================//
          if (Copy(Linha, 0, 6)) = 'EM0213' then
          begin
            SQLPARC := TSimpleDataSet.Create(self);
            try

              SQLPARC.Connection := SQLConnection1;
              SQLPARC.ReadOnly := true;
              SQLPARC.DataSet.CommandText := SelectTitulos('', '');
              SQLPARC.Active := true;
              SQLPARC.First;

              while not SQLPARC.Eof do
              begin
                with Cobr.Dup.Add do
                begin
                  nDup := SQLPARC.FieldByName('NroPe3').AsString;
                  // Número da fatura      showmessage(copy(linha,7,60));
                  dVenc := SQLPARC.FieldByName('DtvPe3').AsDateTime;
                  // Data de vencimento    showmessage(datetostr(strtodate(copy(Linha,75,2)+'/'+copy(Linha,72,2)+'/'+copy(Linha,67,4))));
                  vDup := SQLPARC.FieldByName('VlpPe3').AsFloat;
                  // Valor da duplicata    showmessage(copy(linha,77,15));
                end; // with

                SQLPARC.Next;
              end;
            finally
              SQLPARC.Active := False;
              FreeAndnil(SQLPARC);
            end;
            // end; // fim do EM0213
            Readln(Arquivo, Linha); // Ler a proxima linha
          end; // enquanto não chegar ao fim do Arquivo
          // =============================EM0214=============================//
          if (Copy(Linha, 0, 6)) = 'EM0214' then
          begin
            InfAdic.infCpl := trim(SQLCAB.FieldByName('OB1FAT').AsString) + ' '
              + trim(SQLCAB.FieldByName('OB2FAT').AsString) + ' ' +
              trim(SQLCAB.FieldByName('OB3FAT').AsString) + ' ' +
              trim(SQLCAB.FieldByName('OB4FAT').AsString) + ' ' +
              trim(SQLCAB.FieldByName('OB5FAT').AsString) + ' ' +
              trim(SQLCAB.FieldByName('OB6FAT').AsString) + ' ' +
              trim(SQLCAB.FieldByName('OB7FAT').AsString) + ' ' +
              trim(SQLCAB.FieldByName('OB8FAT').AsString);

          end; // fim do EM0214

        end; // Fim das configurações da NFe
        SQLCAB.Next;
      end; // while

      Result := true;

    finally
      SQLCAB.Active := False;
      FreeAndnil(SQLITEM);
    end;
  finally
    SQLCAB.Active := False;
    FreeAndnil(SQLCAB);
  end;
end;

function TForm1.LeftZero(Const Text: string; Const Tam: word;
  Const RetQdoVazio: String = ' '): string;
begin
  Result := trim(Text);
  if Result <> '' then
  begin
    // Remove zeros desnecessario a esquerda
    if length(Result) > Tam then
    begin
      while (length(Result) > Tam) and (Result[1] = '0') do
        Delete(Result, 1, 1);
    end;
    // Preenche com zeros a esquerda
    while length(Result) < Tam do
      Result := '0' + Result;
  end;
  Result := Copy(Result, 1, Tam);
end;

end.
