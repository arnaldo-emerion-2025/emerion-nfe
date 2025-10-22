unit uMsg;

interface

uses Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, StdCtrls, Dialogs;

type
  TEventForm = class
  public
    procedure OnShowSleepBtn(Sender : TObject);
    procedure BtnClick(Sender : TObject);
  end;

procedure ValidaRejeicao(Msg : String; CaptionForm : string = 'Rejei'#231#227'o NFe'); 
procedure consultaCliente(cgc, ie, status : String);

var
  Frm : TForm;
  lb : TLabel;
  lb2 : TLabel;
  lb3 : TLabel;
  lb4 : TLabel;
  lb5 : TLabel;
  lb6 : TLabel;
  btn : TButton;
  ckb : TCheckBox;
  EventForm : TEventForm;

implementation

procedure ValidaRejeicao(Msg : String; CaptionForm : string = 'Rejei'#231#227'o NFe');
begin

  Frm := TForm.Create(Application);
  try
    with Frm do
    begin
      Left := 0;
      Top := 0;
      BorderIcons := [];
      BorderStyle := bsDialog;
      Caption := CaptionForm; // 'Rejei'#231#227'o NFe';
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
      FreeAndnil(EventForm); ;
    End;

  Finally
    FreeAndnil(Frm);
  end;

end;

procedure consultaCliente(cgc, ie, status : String);
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
    FreeAndnil(Frm);
  end;

end;

{ TEventForm }

procedure TEventForm.OnShowSleepBtn(Sender : TObject);
begin
  Application.ProcessMessages;
  TLabel(TForm(Sender).FindComponent('lbMsg')).Repaint;

  try
    TCheckBox(TForm(Sender).FindComponent('ckbli')).Enabled := False;
    TButton(TForm(Sender).FindComponent('btnOk')).Enabled := False;
    TForm(Sender).DisableAlign;

    Sleep(8000);
  finally
    TForm(Sender).EnableAlign;
    TCheckBox(TForm(Sender).FindComponent('ckbli')).Enabled := true;
    TButton(TForm(Sender).FindComponent('btnOk')).Enabled := true;
  end;

  Application.ProcessMessages;
end;

procedure TEventForm.BtnClick(Sender : TObject);
begin
  if ckb.Checked = true then
  begin
    Frm.ModalResult := 1;
  end
  else
    ShowMessage('Obrigatório confirmar leitura da mensagem.');

end;

end.
