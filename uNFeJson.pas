unit uNFeJson;

interface

uses
  dialogs,
  ulkJSON,
  ACBrNFe,
  ACBrNFe.Classes,
  ACBrDFe.Conversao,
  pcnConversaoNFe,
  SysUtils;

procedure sendNFe();
function fulfillIde(ide: TIde): TIde;

implementation

procedure sendNFe();
var
  ACBrNFe1: TACBrNFe;
  ide: TIde;
begin
  ACBrNFe1 := TACBrNFe.Create(nil);
  ACBrNFe1.NotasFiscais.Clear;
  ACBrNFe1.NotasFiscais.Add;

  ide := fulfillIde(ACBrNFe1.NotasFiscais.Items[0].NFe.ide);
  ShowMessage(ide.natOp);
end;

function fulfillIde(ide: TIde): TIde;
begin
  ide.cUF := 35;
  ide.cNF := 1;
  ide.natOp := 'VENDA';
  ide.modelo := 55;
  ide.serie := 1;
  ide.nNF := 40508;
  ide.dEmi := now;
  ide.tpNF := TTipoNFe.tnSaida;
  ide.idDest := TpcnDestinoOperacao.doInterna;
  ide.cMunFG := 3550308;
  ide.tpImp := TACBrTipoImpressao.tiRetrato;
  ide.tpEmis := TACBrTipoEmissao.teNormal;
  ide.cDV := 0;
  ide.tpAmb := TACBrTipoAmbiente.taHomologacao;
  ide.finNFe := TpcnFinalidadeNFe.fnNormal;
  ide.indFinal := TpcnConsumidorFinal.cfNao;
  ide.indPres := TpcnPresencaComprador.pcNao;
  ide.procEmi := TACBrProcessoEmissao.peAplicativoContribuinte;
  ide.verProc := 'EMERION FATURA';

  Result := ide;
end;

end.
