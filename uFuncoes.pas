unit uFuncoes;

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
  StdCtrls,
  Dialogs,
  Mapi,
  Inifiles,
  FileCtrl,
  SqlExpr,
  DBTables,
  DBXFirebird,
  SimpleDS,
  DB,
  ACBrDFe.Conversao;

type
  TEventForm = class
  public
    procedure OnShowSleepBtn(Sender: TObject);
    procedure BtnClick(Sender: TObject);
  end;

function EmailGerenciadorPadrao(const Assunto, Texto, Anexo, Nome_Remetente,
  Email_Remetente, Nome_Destinatario, Email_Destinatario: string): Integer;
function GetImgLogConfigINI: String;
procedure ValidaRejeicao(Msg: String);
procedure consultaCliente(cgc, ie, status: String);

function Replicate(strRep: string; N: Integer): string;
function SoNumero(strCLPIPI: string): string;
function TruncaAlinhaTexto(strTexto: string; intTamanho: Integer;
  cAlinha: string = 'E'; strPreenche: String = ' '): string;
function GetSelecionaDiretorio(Title: string): string;

procedure ConexaoAlias(SQLConex: TSQLConnection); // Insere base BDE

procedure ConexaoSQLDataBase(con: TSQLConnection); // Usar Este
function GetServerNameBDE: String;

procedure SalvaStringTxt(strtxt: string; strNomeArq: string);

function BuscaSimples(Tabela, Retorna, _and: string; Conn: TSQLConnection;
  Sel: string = ''): string;

function VALIDAcpf(cpf: string): Boolean;
function StrToTpEmis(const codigo: Integer): TACBrTipoEmissao;

var
  Frm: TForm;
  lb: TLabel;
  lb2: TLabel;
  lb3: TLabel;
  lb4: TLabel;
  lb5: TLabel;
  lb6: TLabel;
  btn: TButton;
  ckb: TCheckBox;
  EventForm: TEventForm;

implementation

function EmailGerenciadorPadrao(const Assunto, Texto, Anexo, Nome_Remetente,
  Email_Remetente, Nome_Destinatario, Email_Destinatario: string): Integer;
var
  message: TMapiMessage;
  lpSender, lpRecepient: TMapiRecipDesc;
  FileAttach: TMapiFileDesc;
  SM: TFNMapiSendMail;
  MAPIModule: HModule;
begin
  // Função envia email utilizando gerenciador padrão de e-mail.

  FillChar(message, SizeOf(message), 0);
  with message do
  begin
    if (Assunto <> '') then
    begin
      lpszSubject := PAnsiChar(Assunto);
    end;
    if (Texto <> '') then
    begin
      lpszNoteText := PAnsiChar(Texto)
    end;
    if (Email_Remetente <> '') then
    begin
      lpSender.ulRecipClass := MAPI_ORIG;
      if (Nome_Remetente = '') then
      begin
        lpSender.lpszName := PAnsiChar(Email_Remetente)
      end
      else
      begin
        lpSender.lpszName := PAnsiChar(Nome_Remetente)
      end;
      lpSender.lpszAddress := PAnsiChar('SMTP:' + Email_Remetente);
      lpSender.ulReserved := 0;
      lpSender.ulEIDSize := 0;
      lpSender.lpEntryID := nil;
      lpOriginator := @lpSender;
    end;
    if (Email_Destinatario <> '') then
    begin
      lpRecepient.ulRecipClass := MAPI_TO;
      if (Nome_Destinatario = '') then
      begin
        lpRecepient.lpszName := PAnsiChar(Email_Destinatario)
      end
      else
      begin
        lpRecepient.lpszName := PAnsiChar(Nome_Destinatario)
      end;
      lpRecepient.lpszAddress := PAnsiChar('SMTP:' + Email_Destinatario);
      lpRecepient.ulReserved := 0;
      lpRecepient.ulEIDSize := 0;
      lpRecepient.lpEntryID := nil;
      nRecipCount := 1;
      lpRecips := @lpRecepient;
    end
    else
    begin
      lpRecips := nil
    end;
    if (Anexo = '') then
    begin
      nFileCount := 0;
      lpFiles := nil;
    end
    else
    begin
      FillChar(FileAttach, SizeOf(FileAttach), 0);
      FileAttach.nPosition := Cardinal($FFFFFFFF);
      FileAttach.lpszPathName := PAnsiChar(Anexo);
      nFileCount := 1;
      lpFiles := @FileAttach;
    end;
  end;
  MAPIModule := LoadLibrary(PWideChar(MAPIDLL));
  if MAPIModule = 0 then
  begin
    Result := -1
  end
  else
  begin
    try
      @SM := GetProcAddress(MAPIModule, 'MAPISendMail');
      if @SM <> nil then
      begin
        Result := SM(0, Application.Handle, message, MAPI_DIALOG or
          MAPI_LOGON_UI, 0);
      end
      else
      begin
        Result := 1
      end;

    finally
      FreeLibrary(MAPIModule);
    end;
  end;
  if Result <> 0 then
  begin
    MessageDlg('Error sending mail (' + IntToStr(Result) + ').', mtError,
      [mbOk], 0)
  end;
end;

function GetImgLogConfigINI: String;
var
  Ini: TInifile;
begin

  // Busca INI emn Config.ini Logo pra impressão
  Result := '';

  Ini := TInifile.Create(ExtractFilePath(Application.ExeName) + 'Config.ini');

  try
    Result := Ini.ReadString('IMG', 'LOGO', '');
  finally
    FreeAndNil(Ini);
  end;

end;

procedure ValidaRejeicao(Msg: String);
begin

  Frm := TForm.Create(Application);
  try
    with Frm do
    begin
      Left := 0;
      Top := 0;
      BorderIcons := [];
      BorderStyle := bsDialog;
      Caption := 'Rejei'#231#227'o NFe';
      ClientHeight := 197;
      ClientWidth := 466;
      Color := clBtnFace;
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -11;
      Font.Name := 'Tahoma';
      Font.Style := [];
      OldCreateOrder := False;
      Position := poScreenCenter;
      PixelsPerInch := 96;
      // font.TextHeight := 13;

    end;

    EventForm := TEventForm.Create;
    Try

      lb := TLabel.Create(Frm);
      with lb do
      begin
        Left := 8;
        Top := 7;
        Width := 450;
        Height := 136;
        AutoSize := False;
        Caption := Msg;
        Font.Charset := DEFAULT_CHARSET;
        Font.Color := clWindowText;
        Font.Height := -16;
        Font.Name := 'Tahoma';
        Font.Style := [fsBold];
        Name := 'lbMsg';
        Parent := Frm;
        ParentFont := False;
        WordWrap := true;
      end;

      btn := TButton.Create(Frm);
      with btn do
      begin
        Left := 345;
        Top := 152;
        Width := 97;
        Height := 35;
        Caption := 'OK';
        Enabled := False;
        Font.Charset := DEFAULT_CHARSET;
        Font.Color := clWindowText;
        Font.Height := -19;
        Font.Name := 'Tahoma';
        Font.Style := [fsBold];
        Name := 'btnOk';
        // ModalResult := 1;
        Parent := Frm;
        ParentFont := False;
        TabOrder := 1;
        OnClick := EventForm.BtnClick;
      end;

      ckb := TCheckBox.Create(Frm);

      with ckb do
      begin
        ckb.SetBounds(20, 162, 200, 24);
        Caption := 'Li e desejo continuar.';
        Checked := False;
        Enabled := False;
        Font.Charset := DEFAULT_CHARSET;
        Font.Color := clWindowText;
        Font.Height := -13;
        Font.Name := 'Tahoma';
        Font.Style := [fsBold];
        Name := 'ckbli';
        Parent := Frm;
        ParentFont := False;
        TabOrder := 0;

      end;

      Frm.OnActivate := EventForm.OnShowSleepBtn;
      Application.ProcessMessages;
      Frm.ShowModal;

    Finally
      FreeAndNil(EventForm);;
    End;

  Finally
    FreeAndNil(Frm);
  end;

end;

procedure consultaCliente(cgc, ie, status: String);
begin
  try
    Frm := TForm.Create(Application);

    with Frm do
    begin
      Left := 0;
      Top := 0;
      BorderIcons := [biSystemMenu];
      BorderStyle := bsSingle;
      Caption := 'Status Cliente';
      ClientHeight := 164;
      ClientWidth := 315;
      Color := clBtnFace;
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -11;
      Font.Name := 'Tahoma';
      OldCreateOrder := False;
      Position := poScreenCenter;
      PixelsPerInch := 96;
    end;

    lb := TLabel.Create(Frm);
    with lb do
    begin
      Left := 36;
      Top := 24;
      Width := 34;
      Height := 19;
      Caption := 'CGC';
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -16;
      Font.Name := 'Tahoma';
      Font.Style := [fsBold];
      Parent := Frm;
      ParentFont := False;
    end;

    lb2 := TLabel.Create(Frm);
    with lb2 do
    begin
      Left := 53;
      Top := 56;
      Width := 17;
      Height := 19;
      Caption := 'IE';
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -16;
      Font.Name := 'Tahoma';
      Font.Style := [fsBold];
      Parent := Frm;
      ParentFont := False;
    end;

    lb3 := TLabel.Create(Frm);
    with lb3 do
    begin
      Left := 18;
      Top := 88;
      Width := 52;
      Height := 19;
      Caption := 'Status';
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -16;
      Font.Name := 'Tahoma';
      Font.Style := [fsBold];
      Parent := Frm;
      ParentFont := False;
    end;

    lb4 := TLabel.Create(Frm);
    with lb4 do
    begin
      Left := 92;
      Top := 24;
      Width := 40;
      Height := 19;
      Caption := cgc;
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -16;
      Font.Name := 'Tahoma';
      Font.Style := [];
      Parent := Frm;
      ParentFont := False;
    end;

    lb5 := TLabel.Create(Frm);
    with lb5 do
    begin
      Left := 92;
      Top := 56;
      Width := 29;
      Height := 19;
      Caption := ie;
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -16;
      Font.Name := 'Tahoma';
      Font.Style := [];
      Parent := Frm;
      ParentFont := False;
    end;

    lb6 := TLabel.Create(Frm);
    with lb6 do
    begin
      Left := 92;
      Top := 88;
      Width := 58;
      Height := 19;
      Caption := status;
      Font.Charset := DEFAULT_CHARSET;
      Font.Color := clWindowText;
      Font.Height := -16;
      Font.Name := 'Tahoma';
      Font.Style := [];
      Parent := Frm;
      ParentFont := False;
    end;

    btn := TButton.Create(Frm);
    with btn do
    begin
      Left := 120;
      Top := 128;
      Width := 75;
      Height := 25;
      Caption := '&Ok';
      ModalResult := 1;
      Parent := Frm;
      TabOrder := 0;
    end;

    Frm.ShowModal;
  finally
    FreeAndNil(Frm);
  end;

end;

{ TEventForm }

procedure TEventForm.OnShowSleepBtn(Sender: TObject);
begin
  Application.ProcessMessages;
  TLabel(TForm(Sender).FindComponent('lbMsg')).Repaint;

  Sleep(5000);
  TCheckBox(TForm(Sender).FindComponent('ckbli')).Enabled := true;
  TButton(TForm(Sender).FindComponent('btnOk')).Enabled := true;

  Application.ProcessMessages;
end;

procedure TEventForm.BtnClick(Sender: TObject);
begin
  if ckb.Checked = true then
  begin
    Frm.ModalResult := 1;
  end
  else
    ShowMessage('Obrigatório confirmar leitura da mensagem.');

end;

function Replicate(strRep: string; N: Integer): string;
var
  icont: Integer;
begin
  Result := '';
  // Retorna str repetido tanto quando N manda
  for icont := 1 to N do
  begin
    Result := Result + strRep;
  end;

end;

function SoNumero(strCLPIPI: string): string;
var
  intChar: Integer;
begin
  intChar := 0;
  Result := '';
  for intChar := 1 to Length(strCLPIPI) do
  begin
    if strCLPIPI[intChar] in ['0' .. '9'] then
    begin
      Result := Result + strCLPIPI[intChar];
    end;
  end;

end;

function TruncaAlinhaTexto(strTexto: string; intTamanho: Integer;
  cAlinha: string = 'E'; strPreenche: String = ' '): string;
var
  strAux: string;
begin
  strAux := trim(strTexto);

  // Truncamento
  if Length(strAux) > intTamanho then
  begin
    strAux := copy(strAux, 1, intTamanho);
  end;

  // Alinhamento
  if cAlinha = 'E' then
    Result := strAux + Replicate(strPreenche, intTamanho - Length(strAux))
  else
    Result := Replicate(strPreenche, intTamanho - Length(strAux)) + strAux;

end;

function GetSelecionaDiretorio(Title: string): string;
var
  Pasta: String;
begin
  SelectDirectory(Title, '', Pasta);
  if (trim(Pasta) <> '') then
    if (Pasta[Length(Pasta)] <> '\') then
      Pasta := Pasta + '\';
  Result := Pasta;
end;

function StrToTpEmis(const codigo: Integer): TACBrTipoEmissao;
begin
  case codigo of
    0:
      Result := teNormal;
    1:
      Result := teContingencia;
    2:
      Result := teSCAN;
    3:
      Result := teDPEC;
    4:
      Result := teFSDA;
    5:
      Result := teSVCAN;
    6:
      Result := teSVCRS;
    7:
      Result := teSVCSP;
    8:
      Result := teOffLine;
  else
    raise Exception.CreateFmt
      ('Código inválido (%d). Esperado valor entre 1 e 9.', [codigo]);
  end;
end;

function VALIDAcpf(cpf: string): Boolean;
var
  i: Integer;
  Want: char;
  Wvalid: Boolean;
  Wdigit1, Wdigit2: Integer;
begin
  Result := False;
  if trim(cpf) <> '' then
  begin

    Wdigit1 := 0;
    Wdigit2 := 0;
    Want := cpf[1];
    // variavel para testar se o cpf eh repetido como 111.111.111-11
    Delete(cpf, ansipos('.', cpf), 1); // retira as mascaras se houver
    Delete(cpf, ansipos('.', cpf), 1);
    Delete(cpf, ansipos('-', cpf), 1);

    Wvalid := False;
    // testar se o cpf eh repetido como 111.111.111-11
    for i := 1 to Length(cpf) do
    begin
      if cpf[i] <> Want then
      begin
        Wvalid := true;
        // se o cpf possui um digito diferente ele passou no primeiro teste
        break
      end;
    end;
    // se o cpf eh composto por numeros repetido retorna falso
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
      soma=1o*2+2o*3+3o*4.. ate 9o*10
      digito1 = 11 - soma mod 11
      se digito > 10 digito1 =0
    }

    // verifica se o 1o digito confere
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
      soma=1o*2+2o*3+3o*4.. ate 10o*11
      digito1 = 11 - soma mod 11
      se digito > 10 digito1 =0
    }

    // confere o 2o digito verificador
    if IntToStr(Wdigit2) <> cpf[11] then
    begin
      Result := False;
      exit;
    end;

    // se chegar ate aqui o cpf e valido
    Result := true;
  end;
end;

procedure ConexaoAlias(SQLConex: TSQLConnection);
var
  ApeLin: string;
  ApeAce: string;
  DirAce: string;
  DirUsu: string;
  LinAce: string;
  SeqLin: Integer;
  GDirAce: string;
  ArqTxt: TStringList;
begin
  // Esta procedure foi copiada e feita poucas motificações para atender a necessidade da fazia afins

  ApeLin := UpperCase(ParamStr(1));

  if (trim(ApeLin) <> '') and (trim(ApeLin) <> 'DEFAULT@') then
  begin

    if fileExists(ExtractFilePath(Application.ExeName) + 'login.txt') then
    begin

      ArqTxt := TStringList.Create;
      try
        ArqTxt.LoadFromFile(ExtractFilePath(Application.ExeName) + 'login.txt');

        SeqLin := 0;

        while SeqLin <= (ArqTxt.Count - 1) do
        begin

          LinAce := UpperCase(ArqTxt[SeqLin]);

          if pos(ApeLin, LinAce) > 0 then
          begin

            DirAce := copy(LinAce, pos('##', LinAce) + 2, 100);

            DirAce := copy(DirAce, 1, pos('@@', DirAce) - 1);

            ApeAce := copy(LinAce, pos('@@', LinAce) + 2, 100);

            if pos('@@', ApeAce) > 0 then
              ApeAce := copy(ApeAce, 1, pos('@@', ApeAce) - 1);

          end;

          inc(SeqLin);

        end;

        GDirAce := LowerCase(trim(DirAce));

        SeqLin := 0;

        while SeqLin <= (ArqTxt.Count - 1) do
        begin

          LinAce := UpperCase(ArqTxt[SeqLin]);

          if pos(ApeLin, LinAce) > 0 then
          begin

            DirUsu := copy(LinAce, pos('***', LinAce) + 3, 100);

            DirUsu := copy(DirUsu, 1, pos('***', DirUsu) - 1);

          end;

          inc(SeqLin);

        end;

        SQLConex.Params.Values['Database'] := GDirAce;

      finally
        FreeAndNil(ArqTxt);
      end;

    end;
  end;
end;

procedure ConexaoSQLDataBase(con: TSQLConnection);
var
  PathBD: string;
begin

  PathBD := GetServerNameBDE;

  con.Connected := False;

  con.Params.Values['Database'] := PathBD;

  ConexaoAlias(con);

  con.Connected := true;

end;

function GetServerNameBDE: String;
var
  strList: TStringList;
begin
  strList := TStringList.Create;
  try
    if Session.isAlias('ISade') then
    begin
      Session.GetAliasParams('ISade', strList);
      // ComboBox1.Items:= CurrentAliases;
      Result := strList.Values['SERVER NAME'];
    end
    else
    begin
      ShowMessage('Não encontrado BDE Isade.');
    end;
  finally
    FreeAndNil(strList);
  end;

end;

procedure SalvaStringTxt(strtxt: string; strNomeArq: string);
var
  arq: TextFile;
begin
  if trim(strNomeArq) <> '' then
  begin
    // Se arquivo existe será apagado
    if fileExists(strNomeArq) then
      deletefile(strNomeArq);

    AssignFile(arq, strNomeArq);

    try
      Rewrite(arq);
      WriteLn(arq, strtxt);
    finally
      CloseFile(arq);
    end;
  end;

end;

function BuscaSimples(Tabela, Retorna, _and: string; Conn: TSQLConnection;
  Sel: string = ''): string;
var
  SQLBUS: TSimpleDataSet;
begin
  Result := '';
  SQLBUS := TSimpleDataSet.Create(Nil);
  try
    SQLBUS.Connection := Conn;
    if Sel = '' then
    begin
      if (trim(Tabela) <> '') and (trim(_and) <> '') and (trim(Retorna) <> '')
      then
      begin
        SQLBUS.Active := False;
        SQLBUS.DataSet.CommandText := ' select ' + Retorna + ' From ' + Tabela +
          ' Where 1 = 1 and ' + _and;
        SQLBUS.Active := true;
        if not SQLBUS.IsEmpty then
        begin
          if SQLBUS.FieldByName(Retorna) is TDateTimeField then
            Result := FormatDateTime('dd/mm/yyyy', SQLBUS.FieldByName(Retorna)
              .AsDateTime)
          else if SQLBUS.FieldByName(Retorna) is TFloatField then
            Result := FormatFloat('0.00', SQLBUS.FieldByName(Retorna).AsFloat)
          else
            Result := SQLBUS.FieldByName(Retorna).AsString;
          SQLBUS.Active := False;
        end;
      end;
    end
    else
    begin
      SQLBUS.Active := False;
      SQLBUS.DataSet.CommandText := Sel;
      SQLBUS.Active := true;

      if not SQLBUS.IsEmpty then
      begin
        if SQLBUS.Fields[0] is TDateTimeField then
          Result := FormatDateTime('dd/mm/yyyy', SQLBUS.FieldByName(Retorna)
            .AsDateTime)
        else if SQLBUS.Fields[0] is TFloatField then
          Result := FormatFloat('0.00', SQLBUS.FieldByName(Retorna).AsFloat)
        else
          Result := SQLBUS.Fields[0].AsString;
        SQLBUS.Active := False;
      end;
    end;
  finally
    FreeAndNil(SQLBUS);
  end;
end;

end.
