unit uNFeJson;

interface

uses
  dialogs,
  ulkJSON,
  ACBrNFe,
  ACBrNFeNotasFiscais,
  ACBrNFe.Classes,
  ACBrDFe.Conversao,
  pcnConversaoNFe,
  SysUtils,
  uJsonHelper,
  uWebServiceTributos;

function isDateInFuture(const dateStr: string): Boolean;
function strToDate_YMD(const dateStr: string): TDateTime;

function sendNFe(ACBrNFe1: TACBrNFe; jsonString: String): Integer;
procedure fulfillIde(ide: TIde; jsonObj: TlkJSONobject);
procedure fulfillEmit(emit: TEmit; jsonObj: TlkJSONobject);
procedure fulfillDest(dest: TDest; jsonObj: TlkJSONobject);
procedure fulfillEntrega(entrega: TEntrega; jsonObj: TlkJSONobject);
procedure fulfillDet(det: TDetCollection; jsonObj: TlkJSONobject);
procedure fulfillIcmsTot(total: TTotal; jsonObj: TlkJSONobject);
procedure fulfillTransport(transport: TTransp; jsonObj: TlkJSONobject);
procedure fulfillExport(exporta: TExporta; jsonObj: TlkJSONobject);
procedure fulfillPayment(pag: TpagCollection; jsonObj: TlkJSONobject);
procedure fulfillCob(cobranca: TCobr; jsonObj: TlkJSONobject);

procedure fulfillIBSCBS000(imposto: TImposto);
procedure fulfillIBSCBS200(imposto: TImposto);

procedure fulfillRespTec(nfeObj: TNFe);

implementation

function sendNFe(ACBrNFe1: TACBrNFe; jsonString: String): Integer;
var
  nfeObj: TNFe;
  jsonObj: TlkJSONobject;
begin

  //TWebServiceTributos.calcularNovosImpostos; exit(1);

  ShowMessage('Enviando NFE com o modelo mais atual');

  jsonObj := TlkJSON.ParseText(jsonString) as TlkJSONobject;
  ACBrNFe1.NotasFiscais.Clear;
  ACBrNFe1.NotasFiscais.Add;

  nfeObj := ACBrNFe1.NotasFiscais.Items[0].NFe;

  fulfillIde(nfeObj.ide, jsonObj);
  fulfillEmit(nfeObj.emit, jsonObj);
  fulfillDest(nfeObj.dest, jsonObj);
  fulfillEntrega(nfeObj.entrega, jsonObj);
  fulfillDet(nfeObj.det, jsonObj);
  fulfillIcmsTot(nfeObj.total, jsonObj);
  fulfillTransport(nfeObj.Transp, jsonObj);
  fulfillExport(nfeObj.exporta, jsonObj);
  fulfillPayment(nfeObj.pag, jsonObj);
  fulfillCob(nfeObj.Cobr, jsonObj);
  nfeObj.InfAdic.infCpl := jsonObj.Field['infCpl'].Value;

  if (nfeObj.emit.EnderEmit.UF = 'SC') then
  begin
    fulfillRespTec(nfeObj);
  end;

  if (ACBrNFe1.Configuracoes.WebServices.Ambiente = TACBrTipoAmbiente.taHomologacao) then
    nfeObj.emit.xNome := 'NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL';

  Result := nfeObj.ide.nNF;
end;

procedure fulfillIde(ide: TIde; jsonObj: TlkJSONobject);
var
  ok: Boolean;
begin
  if (jsonObj.Field['nfContingencia'].Count > 0) then
  begin
    ide.dhCont := now;
    ide.xJust := jsonObj.Field['nfContingencia'].Field['xJust'].Value;
  end;

  ide.cUF := jsonObj.Field['ide'].Field['cUF'].Value;
  // ide.cNF := 1;
  ide.natOp := jsonObj.Field['ide'].Field['natOp'].Value;
  ide.indPag := StrToIndpagEX(jsonObj.Field['ide'].Field['TipCnd'].Value);
  ide.modelo := jsonObj.Field['ide'].Field['modelo'].Value;
  ide.serie := jsonObj.Field['ide'].Field['serie'].Value;
  ide.nNF := jsonObj.Field['ide'].Field['nNF'].Value;
  ide.dEmi := now;

  if (jsonObj.Field['ide'].Field['dSaiEnt'].Value <> '0000-00-00') then
    ide.dSaiEnt := jsonObj.Field['ide'].Field['dSaiEnt'].Value;

  ide.tpNF := StrToTpNF(jsonObj.Field['ide'].Field['tpNF'].Value);
  ide.idDest := StrToDestinoOperacao(ok, jsonObj.Field['ide'].Field['idDest'].Value);
  ide.cMunFG := jsonObj.Field['ide'].Field['cMunFG'].Value;
  ide.tpImp := StrToTpImp(jsonObj.Field['ide'].Field['tpImp'].Value);
  ide.tpEmis := StrToTipoEmissao(jsonObj.Field['ide'].Field['tpEmis'].Value);
  ide.cDV := jsonObj.Field['ide'].Field['cDV'].Value;
  ide.tpAmb := StrToTipoAmbiente(jsonObj.Field['ide'].Field['tpAmb'].Value);
  ide.finNFe := StrToFinNFe(ok, jsonObj.Field['ide'].Field['finNFe'].Value);
  ide.indFinal := StrToConsumidorFinal(ok, jsonObj.Field['ide'].Field['indFinal'].Value);
  ide.indPres := StrToPresencaComprador(ok, jsonObj.Field['ide'].Field['indPres'].Value);
  ide.procEmi := StrToprocEmi(jsonObj.Field['ide'].Field['procEmi'].Value);
  ide.verProc := jsonObj.Field['ide'].Field['verProc'].Value;
end;

procedure fulfillEmit(emit: TEmit; jsonObj: TlkJSONobject);
var
  ok: Boolean;
begin
  emit.CNPJCPF := jsonObj.Field['emit'].Field['CNPJ'].Value;
  emit.xNome := jsonObj.Field['emit'].Field['xNome'].Value;
  emit.xFant := jsonObj.Field['emit'].Field['xFant'].Value;
  emit.EnderEmit.xLgr := jsonObj.Field['emit'].Field['enderEmit'].Field['xLgr'].Value;
  emit.EnderEmit.nro := jsonObj.Field['emit'].Field['enderEmit'].Field['nro'].Value;
  emit.EnderEmit.xCpl := jsonObj.Field['emit'].Field['enderEmit'].Field['xCpl'].Value;
  emit.EnderEmit.xBairro := jsonObj.Field['emit'].Field['enderEmit'].Field['xBairro'].Value;
  emit.EnderEmit.cMun := jsonObj.Field['emit'].Field['enderEmit'].Field['cMun'].Value;
  emit.EnderEmit.xMun := jsonObj.Field['emit'].Field['enderEmit'].Field['xMun'].Value;
  emit.EnderEmit.UF := jsonObj.Field['emit'].Field['enderEmit'].Field['UF'].Value;
  emit.EnderEmit.CEP := jsonObj.Field['emit'].Field['enderEmit'].Field['CEP'].Value;
  emit.EnderEmit.cPais := jsonObj.Field['emit'].Field['enderEmit'].Field['cPais'].Value;
  emit.EnderEmit.xPais := jsonObj.Field['emit'].Field['enderEmit'].Field['xPais'].Value;
  emit.EnderEmit.fone := jsonObj.Field['emit'].Field['enderEmit'].Field['fone'].Value;
  emit.IE := jsonObj.Field['emit'].Field['ie'].Value;
  emit.IEST := jsonObj.Field['emit'].Field['IEST'].Value;
  emit.CRT := StrToCRT(ok, jsonObj.Field['emit'].Field['CRT'].Value);
end;

procedure fulfillDest(dest: TDest; jsonObj: TlkJSONobject);
begin
  dest.CNPJCPF := jsonObj.Field['dest'].Field['CPF_CNPJ'].Value;
  dest.idEstrangeiro := jsonObj.Field['dest'].Field['idEstrangeiro'].Value;
  dest.xNome := jsonObj.Field['dest'].Field['xNome'].Value;
  dest.EnderDest.xLgr := jsonObj.Field['dest'].Field['enderDest'].Field['xLgr'].Value;
  dest.EnderDest.nro := jsonObj.Field['dest'].Field['enderDest'].Field['nro'].Value;
  dest.EnderDest.xCpl := jsonObj.Field['dest'].Field['enderDest'].Field['xCpl'].Value;
  dest.EnderDest.xBairro := jsonObj.Field['dest'].Field['enderDest'].Field['xBairro'].Value;
  dest.EnderDest.cMun := jsonObj.Field['dest'].Field['enderDest'].Field['cMun'].Value;
  dest.EnderDest.xMun := jsonObj.Field['dest'].Field['enderDest'].Field['xMun'].Value;
  dest.EnderDest.UF := jsonObj.Field['dest'].Field['enderDest'].Field['UF'].Value;
  dest.EnderDest.CEP := jsonObj.Field['dest'].Field['enderDest'].Field['CEP'].Value;
  dest.EnderDest.cPais := jsonObj.Field['dest'].Field['enderDest'].Field['cPais'].Value;
  dest.EnderDest.xPais := jsonObj.Field['dest'].Field['enderDest'].Field['xPais'].Value;
  dest.EnderDest.fone := jsonObj.Field['dest'].Field['enderDest'].Field['fone'].Value;
  dest.IE := jsonObj.Field['dest'].Field['ie'].Value;
  dest.ISUF := jsonObj.Field['dest'].Field['ISUF'].Value;
  dest.Email := jsonObj.Field['dest'].Field['email'].Value;
  dest.indIEDest := StrToindIEDest(jsonObj.Field['dest'].Field['indIEDest'].Value);
  dest.IM := jsonObj.Field['dest'].Field['IM'].Value;
end;

procedure fulfillRespTec(nfeObj: TNFe);
begin
  nfeObj.infRespTec.CNPJ := '05557708000161';
  nfeObj.infRespTec.xContato := 'CARLOS ALBERTO FERREIRA DOS SANTOS';
  nfeObj.infRespTec.Email := 'carlos@emerion.com.br';
  nfeObj.infRespTec.fone := '01120217221';
end;

procedure fulfillEntrega(entrega: TEntrega; jsonObj: TlkJSONobject);
begin
  if (jsonObj.Field['entrega'].Count = 0) then
    exit;
  entrega.CNPJCPF := jsonObj.Field['entrega'].Field['CPF_CNPJ'].Value;
  entrega.xLgr := jsonObj.Field['entrega'].Field['xLgr'].Value;
  entrega.nro := jsonObj.Field['entrega'].Field['nro'].Value;
  entrega.xCpl := jsonObj.Field['entrega'].Field['xCpl'].Value;
  entrega.xBairro := jsonObj.Field['entrega'].Field['xBairro'].Value;
  entrega.cMun := jsonObj.Field['entrega'].Field['cMun'].Value;
  entrega.xMun := jsonObj.Field['entrega'].Field['xMun'].Value;
  entrega.UF := jsonObj.Field['entrega'].Field['UF'].Value;
end;

procedure fulfillDet(det: TDetCollection; jsonObj: TlkJSONobject);
var
  detArray, rastroArray, medArray, diArray, adiArray: TlkJSONlist;
  detItem, rastroItem, medItem, diItem, adiItem: TlkJSONobject;
  item: TDetCollectionItem;
  detRastroItem: TRastroCollectionItem;
  detMedItem: TMedCollectionItem;
  detDiItem: TDICollectionItem;
  detAdiItem: TADICollectionItem;
  i, j, k, l, m: Integer;
  ok: Boolean;
begin
  detArray := jsonObj.Field['det'] as TlkJSONlist;
  for i := 0 to detArray.Count - 1 do
  begin
    detItem := detArray.Child[i] as TlkJSONobject;
    item := det.New;

    item.Prod.nItem := detItem.Field['item'].Field['nItem'].Value;
    item.Prod.cProd := detItem.Field['item'].Field['cProd'].Value;
    item.Prod.cEAN := detItem.Field['item'].Field['cEAN'].Value;
    item.Prod.xProd := detItem.Field['item'].Field['xProd'].Value;
    item.Prod.NCM := detItem.Field['item'].Field['NCM'].Value;
    item.Prod.EXTIPI := detItem.Field['item'].Field['EXTIPI'].Value;
    item.Prod.CFOP := detItem.Field['item'].Field['CFOP'].Value;
    item.Prod.CEST := detItem.Field['item'].Field['CEST'].Value;
    item.Prod.uCom := detItem.Field['item'].Field['uCom'].Value;
    item.Prod.qCom := detItem.Field['item'].Field['qCom'].Value;
    item.Prod.vUnCom := detItem.Field['item'].Field['vUnCom'].Value;
    item.Prod.vProd := detItem.Field['item'].Field['vProd'].Value;
    item.Prod.cEANTrib := detItem.Field['item'].Field['cEANTrib'].Value;
    item.Prod.uTrib := detItem.Field['item'].Field['uTrib'].Value;
    item.Prod.qTrib := detItem.Field['item'].Field['qTrib'].Value;
    item.Prod.vUnTrib := detItem.Field['item'].Field['vUnTrib'].Value;
    item.Prod.vFrete := detItem.Field['item'].Field['vFrete'].Value;
    item.Prod.vSeg := detItem.Field['item'].Field['vSeg'].Value;
    item.Prod.vOutro := detItem.Field['item'].Field['vOutro'].Value;
    item.Prod.vDesc := detItem.Field['item'].Field['vDesc'].Value;
    item.Prod.nFCI := detItem.Field['item'].Field['nFCI'].Value;

    //fulfillIBSCBS000(item.Imposto);

    item.Prod.xPed := detItem.Field['item'].Field['xPed'].Value;
    item.Prod.nItemPed := detItem.Field['item'].Field['nItemPed'].Value;

    if (detItem.Field['item'].Field['cProdANP'].Value <> '') then
    begin
      item.Prod.comb.cProdANP := detItem.Field['item'].Field['cProdANP'].Value;
      item.Prod.comb.descANP := detItem.Field['item'].Field['descANP'].Value;
      if (jsonObj.Field['entrega'].Count > 0) then
        item.Prod.comb.UFcons := detItem.Field['entrega'].Field['UF'].Value
      else
        item.Prod.comb.UFcons := jsonObj.Field['dest'].Field['enderDest'].Field['UF'].Value;
      item.Prod.comb.CODIF := detItem.Field['item'].Field['CODIF'].Value;
    end;

    item.infAdProd := detItem.Field['infAdProd'].Value;

    if (detArray.Child[i].Field['di'].Count > 0) then
    begin
      diArray := detArray.Child[i].Field['di'] as TlkJSONlist;
      for l := 0 to diArray.Count - 1 do
      begin
        diItem := diArray.Child[l] as TlkJSONobject;
        detDiItem := item.Prod.di.New;

        detDiItem.nDi := diItem.Field['nDi'].Value;
        detDiItem.dDi :=  strToDate_YMD(diItem.Field['dDi'].Value);
        detDiItem.xLocDesemb := diItem.Field['xLocDesemb'].Value;
        detDiItem.UFDesemb := diItem.Field['UFDesemb'].Value;
        detDiItem.dDesemb := strToDate_YMD(diItem.Field['dDesemb'].Value);
        detDiItem.cExportador := diItem.Field['cExportador'].Value;

        if (diArray.Child[i].Field['adi'].Count > 0) then
        begin
          adiArray := diArray.Child[i].Field['adi'] as TlkJSONlist;
          for m := 0 to adiArray.Count - 1 do
          begin
            adiItem := adiArray.Child[m] as TlkJSONobject;
            detAdiItem := detDiItem.adi.New;
            detAdiItem.nSeqAdi := adiItem.Field['nSeqAdi'].Value;
            detAdiItem.nAdicao := adiItem.Field['nAdicao'].Value;
            detAdiItem.cFabricante := adiItem.Field['cFabricante'].Value;
            detAdiItem.vDescDI := adiItem.Field['vDescDI'].Value;
          end;
        end;
      end;
    end;

    item.Imposto.II.vBc := detItem.Field['imposto'].Field['importoImportacao'].Field['vBC'].Value;
    item.Imposto.II.vII := detItem.Field['imposto'].Field['importoImportacao'].Field['vII'].Value;
    item.Imposto.II.vDespAdu := detItem.Field['imposto'].Field['importoImportacao'].Field['vDespAdu'].Value;
    item.Imposto.II.vIOF := detItem.Field['imposto'].Field['importoImportacao'].Field['vIOF'].Value;
    item.Imposto.vTotTrib := detItem.Field['imposto'].Field['vTotTrib'].Value;

    if (detArray.Child[i].Field['rastro'].Count > 0) then
    begin
      rastroArray := detArray.Child[i].Field['rastro'] as TlkJSONlist;
      for j := 0 to rastroArray.Count - 1 do
      begin
        rastroItem := rastroArray.Child[j] as TlkJSONobject;
        detRastroItem := item.Prod.rastro.New;

        detRastroItem.nLote := rastroItem.Field['nLote'].Value;
        detRastroItem.qLote := rastroItem.Field['qLote'].Value;
        detRastroItem.dFab := rastroItem.Field['dFab'].Value;
        detRastroItem.dVal := rastroItem.Field['dVal'].Value;
      end;
    end;

    if (jsonObj.Field['det'].Field['med'].Count > 0) then
    begin
      medArray := detArray.Child[i].Field['med'] as TlkJSONlist;
      for k := 0 to medArray.Count - 1 do
      begin
        medItem := medArray.Child[k] as TlkJSONobject;
        detMedItem := item.Prod.med.New;

        detMedItem.nLote := medItem.Field['nLote'].Value;
        detMedItem.qLote := medItem.Field['qLote'].Value;
        detMedItem.dFab := medItem.Field['dFab'].Value;
        detMedItem.dVal := medItem.Field['dVal'].Value;
        detMedItem.vPMC := medItem.Field['vPMC'].Value;
      end;
    end;

    item.Imposto.ICMS.orig := StrToOrig(detItem.Field['imposto'].Field['icms'].Field['orig'].Value);
    if (Length(detItem.Field['imposto'].Field['icms'].Field['cst'].Value) = 2) then
    begin
      item.Imposto.ICMS.CST := StrToCSTICMS(detItem.Field['imposto'].Field['icms'].Field['cst'].Value);
      item.Imposto.ICMS.vBCST := detItem.Field['imposto'].Field['icms'].Field['vBCST'].Value;
      item.Imposto.ICMS.vICMSST := detItem.Field['imposto'].Field['icms'].Field['vICMSST'].Value;
    end
    else
    begin
      item.Imposto.ICMS.CSOSN := StrToCSOSNIcms(detItem.Field['imposto'].Field['icms'].Field['cst'].Value);
      item.Imposto.ICMS.pCredSN := detItem.Field['imposto'].Field['icms'].Field['pCredSN'].Value;
      item.Imposto.ICMS.vCredICMSSN := detItem.Field['imposto'].Field['icms'].Field['vCredICMSSN'].Value;
    end;
    // detItem.Field['imposto'].Field['icms'].Field['modBCST'].Value;
    // For some reason this value should not come from the file. It is a constant
    item.Imposto.ICMS.modBCST := TpcnDeterminacaoBaseIcmsST.dbisPrecoTabelado;
    item.Imposto.ICMS.pICMSST := detItem.Field['imposto'].Field['icms'].Field['pICMSST'].Value;
    item.Imposto.ICMS.pMVAST := detItem.Field['imposto'].Field['icms'].Field['pMVAST'].Value;
    item.Imposto.ICMS.pRedBC := detItem.Field['imposto'].Field['icms'].Field['pRedBC'].Value;
    item.Imposto.ICMS.vBc := detItem.Field['imposto'].Field['icms'].Field['vBC'].Value;
    item.Imposto.ICMS.pICMS := detItem.Field['imposto'].Field['icms'].Field['pICMS'].Value;
    item.Imposto.ICMS.vICMS := detItem.Field['imposto'].Field['icms'].Field['vICMS'].Value;
    item.Imposto.ICMS.vICMSDeson := detItem.Field['imposto'].Field['icms'].Field['vICMSDeson'].Value;

    if ((StrToCSTICMS(detItem.Field['imposto'].Field['icms'].Field['cst'].Value) = cst60) and
      (getTipoConsumidor(jsonObj) <> cfConsumidorFinal)) then
    begin
      if (detItem.Field['imposto'].Field['icms'].Field['pST'].Value > 0) then
      begin
        item.Imposto.ICMS.vBCSTRet := detItem.Field['imposto'].Field['icms'].Field['vBCSTRet'].Value;
        item.Imposto.ICMS.vICMSSubstituto := detItem.Field['imposto'].Field['icms'].Field['vICMSSubstituto'].Value;
        item.Imposto.ICMS.vICMSSTRet := detItem.Field['imposto'].Field['icms'].Field['vICMSSTRet'].Value;
        item.Imposto.ICMS.pST := detItem.Field['imposto'].Field['icms'].Field['pST'].Value;
      end;
    end;

    item.Imposto.ICMS.motDesICMS := StrTomotDesICMS(ok, detItem.Field['imposto'].Field['icms'].Field
      ['motDesICMS'].Value);

    if(detItem.Field['imposto'].Field['ICMSUFDest'] <> Nil) then
    begin
      if (item.Imposto.ICMS.CST = cst00) then
      begin
        if (detItem.Field['imposto'].Field['ICMSUFDest'].Field['vBCFCPUFDest'].Value > 0) then
        begin
          if not((StrToDestinoOperacao(ok, jsonObj.Field['ide'].Field['idDest'].Value) = doInterestadual) and
            (getTipoConsumidor(jsonObj) = cfConsumidorFinal)) then
          begin
            item.Imposto.ICMS.vBCFCP := detItem.Field['imposto'].Field['icms'].Field['vBC'].Value;
            item.Imposto.ICMS.pFCP := detItem.Field['imposto'].Field['ICMSUFDest'].Field['vBCFCPUFDest'].Value;
            item.Imposto.ICMS.vFCP := detItem.Field['imposto'].Field['ICMSUFDest'].Field['vICMSUFDest'].Value;
          end
          else
          begin
            item.Imposto.ICMS.vBCFCP := detItem.Field['imposto'].Field['icms'].Field['vBC'].Value;
            item.Imposto.ICMSUFDest.pFCPUFDest := detItem.Field['imposto'].Field['ICMSUFDest'].Field
              ['vBCFCPUFDest'].Value;
            item.Imposto.ICMSUFDest.vFCPUFDest := detItem.Field['imposto'].Field['ICMSUFDest'].Field['vFCPUFDest'].Value;
          end;
        end;
      end;

      item.Imposto.ICMSUFDest.vBCUFDest := detItem.Field['imposto'].Field['ICMSUFDest'].Field['vBCUFDest'].Value;
      item.Imposto.ICMSUFDest.vBCFCPUFDest := detItem.Field['imposto'].Field['ICMSUFDest'].Field['vBCUFDest'].Value;
      item.Imposto.ICMSUFDest.pICMSUFDest := detItem.Field['imposto'].Field['ICMSUFDest'].Field['pICMSUFDest'].Value;
      item.Imposto.ICMSUFDest.pICMSInter := detItem.Field['imposto'].Field['ICMSUFDest'].Field['pICMSInter'].Value;
      item.Imposto.ICMSUFDest.pICMSInterPart := detItem.Field['imposto'].Field['ICMSUFDest'].Field['pICMSInterPart'].Value;
      item.Imposto.ICMSUFDest.vICMSUFDest := detItem.Field['imposto'].Field['ICMSUFDest'].Field['vICMSUFDest'].Value;
      item.Imposto.ICMSUFDest.vICMSUFRemet := detItem.Field['imposto'].Field['ICMSUFDest'].Field['vICMSUFRemet'].Value;
    end;

    item.Imposto.IPI.vBc := detItem.Field['imposto'].Field['ipi'].Field['vBC'].Value;
    item.Imposto.IPI.pIPI := detItem.Field['imposto'].Field['ipi'].Field['pIPI'].Value;
    item.Imposto.IPI.vIPI := detItem.Field['imposto'].Field['ipi'].Field['vIPI'].Value;
    item.Imposto.IPI.CST := StrToCSTIPI(ok, detItem.Field['imposto'].Field['ipi'].Field['CST'].Value);

    item.Imposto.PIS.vBc := detItem.Field['imposto'].Field['pis'].Field['vBC'].Value;
    item.Imposto.PIS.pPIS := detItem.Field['imposto'].Field['pis'].Field['pIPI'].Value;
    item.Imposto.PIS.vPIS := detItem.Field['imposto'].Field['pis'].Field['vIPI'].Value;
    item.Imposto.PIS.CST := StrToCSTPIS(detItem.Field['imposto'].Field['pis'].Field['CST'].Value);

    item.Imposto.COFINS.vBc := detItem.Field['imposto'].Field['cofins'].Field['vBC'].Value;
    item.Imposto.COFINS.pCOFINS := detItem.Field['imposto'].Field['cofins'].Field['pIPI'].Value;
    item.Imposto.COFINS.vCOFINS := detItem.Field['imposto'].Field['cofins'].Field['vIPI'].Value;
    item.Imposto.COFINS.CST := StrToCSTCOFINS(detItem.Field['imposto'].Field['cofins'].Field['CST'].Value);
  end;
end;

procedure fulfillIBSCBS000(imposto: TImposto);
var
 novoImposto : TIBSCBS;
begin
  novoImposto := imposto.IBSCBS;

  novoImposto.CST := TCSTIBSCBS.cst000;
  novoImposto.cClassTrib := '';
  novoImposto.indDoacao := TIndicadorEx.tieNenhum;
  novoImposto.gIBSCBS.vBC := 0;
  novoImposto.gIBSCBS.gIBSUF.pIBSUF := 0;
  novoImposto.gIBSCBS.gIBSUF.vIBSUF := 0;
  novoImposto.gIBSCBS.gIBSMun.pIBSMun := 0;
  novoImposto.gIBSCBS.gIBSMun.vIBSMun := 0;
  novoImposto.gIBSCBS.vIBS := 0;
  novoImposto.gIBSCBS.gCBS.pCBS := 0;
  novoImposto.gIBSCBS.gCBS.vCBS := 0;
end;

procedure fulfillIBSCBS200(imposto: TImposto);
var
 novoImposto : TIBSCBS;
begin
  novoImposto := imposto.IBSCBS;

  novoImposto.CST := TCSTIBSCBS.cst000;
  novoImposto.cClassTrib := '';

  novoImposto.gIBSCBS.vBC := 0;

  novoImposto.gIBSCBS.gIBSUF.pIBSUF := 0;
  novoImposto.gIBSCBS.gIBSUF.gRed.pRedAliq := 0;
  novoImposto.gIBSCBS.gIBSUF.gRed.pAliqEfet := 0;
  novoImposto.gIBSCBS.gIBSUF.vIBSUF := 0;

  novoImposto.gIBSCBS.gIBSMun.pIBSMun := 0;
  novoImposto.gIBSCBS.gIBSMun.gRed.pRedAliq := 0;
  novoImposto.gIBSCBS.gIBSMun.gRed.pAliqEfet := 0;
  novoImposto.gIBSCBS.gIBSMun.vIBSMun := 0;

  novoImposto.gIBSCBS.vIBS := 0;

  novoImposto.gIBSCBS.gCBS.pCBS := 0;
  novoImposto.gIBSCBS.gCBS.gRed.pRedAliq := 0;
  novoImposto.gIBSCBS.gCBS.gRed.pAliqEfet := 0;
  novoImposto.gIBSCBS.gCBS.vCBS := 0;

  novoImposto.gIBSCBS.gTribRegular.CSTReg := TCSTIBSCBS.cst200;
  novoImposto.gIBSCBS.gTribRegular.cClassTribReg := '';
  novoImposto.gIBSCBS.gTribRegular.pAliqEfetRegIBSUF := 0;
  novoImposto.gIBSCBS.gTribRegular.vTribRegIBSUF := 0;
  novoImposto.gIBSCBS.gTribRegular.pAliqEfetRegIBSMun := 0;
  novoImposto.gIBSCBS.gTribRegular.vTribRegIBSMun := 0;
  novoImposto.gIBSCBS.gTribRegular.pAliqEfetRegCBS := 0;
  novoImposto.gIBSCBS.gTribRegular.vTribRegCBS := 0;
end;

procedure fulfillIcmsTot(total: TTotal; jsonObj: TlkJSONobject);
begin
  total.ICMSTot.vBc := jsonObj.Field['icmsTot'].Field['vBC'].Value;
  total.ICMSTot.vICMS := jsonObj.Field['icmsTot'].Field['vICMS'].Value;
  total.ICMSTot.vBCST := jsonObj.Field['icmsTot'].Field['vBCST'].Value;
  total.ICMSTot.vST := jsonObj.Field['icmsTot'].Field['vST'].Value;
  total.ICMSTot.vProd := jsonObj.Field['icmsTot'].Field['vProd'].Value;
  total.ICMSTot.vFrete := jsonObj.Field['icmsTot'].Field['vFrete'].Value;
  total.ICMSTot.vSeg := jsonObj.Field['icmsTot'].Field['vSeg'].Value;
  total.ICMSTot.vDesc := jsonObj.Field['icmsTot'].Field['vDesc'].Value;
  total.ICMSTot.vII := jsonObj.Field['icmsTot'].Field['vII'].Value;
  total.ICMSTot.vIPI := jsonObj.Field['icmsTot'].Field['vIPI'].Value;
  total.ICMSTot.vPIS := jsonObj.Field['icmsTot'].Field['vPIS'].Value;
  total.ICMSTot.vCOFINS := jsonObj.Field['icmsTot'].Field['vCOFINS'].Value;
  total.ICMSTot.vOutro := jsonObj.Field['icmsTot'].Field['vOutro'].Value;
  total.ICMSTot.vNF := jsonObj.Field['icmsTot'].Field['vNF'].Value;
  total.ICMSTot.vTotTrib := jsonObj.Field['icmsTot'].Field['vTotTrib'].Value;
  total.ICMSTot.vICMSDeson := jsonObj.Field['icmsTot'].Field['vICMSDeson'].Value;
  total.ICMSTot.vFCPUFDest := jsonObj.Field['icmsTot'].Field['vFCPUFDest'].Value;
  total.ICMSTot.vICMSUFDest := jsonObj.Field['icmsTot'].Field['vICMSUFDest'].Value;
  total.ICMSTot.vICMSUFRemet := jsonObj.Field['icmsTot'].Field['vICMSUFRemet'].Value;
end;

procedure fulfillTransport(transport: TTransp; jsonObj: TlkJSONobject);
var
  ok: Boolean;
  volItem: TVolCollectionItem;
begin
  transport.modFrete := StrTomodFrete(ok, jsonObj.Field['transporte'].Field['modFrete'].Value);
  transport.Transporta.CNPJCPF := jsonObj.Field['transporte'].Field['CNPJCPF'].Value;
  transport.Transporta.xNome := jsonObj.Field['transporte'].Field['xNome'].Value;
  transport.Transporta.IE := jsonObj.Field['transporte'].Field['ie'].Value;
  transport.Transporta.xEnder := jsonObj.Field['transporte'].Field['xEnder'].Value;
  transport.Transporta.xMun := jsonObj.Field['transporte'].Field['xMun'].Value;
  transport.Transporta.UF := jsonObj.Field['transporte'].Field['uf'].Value;

  volItem := transport.Vol.New;
  volItem.qVol := jsonObj.Field['transporte'].Field['qVol'].Value;
  volItem.esp := jsonObj.Field['transporte'].Field['esp'].Value;
  volItem.marca := jsonObj.Field['transporte'].Field['marca'].Value;
  volItem.pesoL := jsonObj.Field['transporte'].Field['pesoL'].Value;
  volItem.pesoB := jsonObj.Field['transporte'].Field['pesoB'].Value;
  volItem.nVol := jsonObj.Field['transporte'].Field['nVol'].Value;

  transport.veicTransp.placa := jsonObj.Field['transporte'].Field['placa'].Value;
  transport.veicTransp.placa := jsonObj.Field['transporte'].Field['ufPlaca'].Value;
end;

procedure fulfillExport(exporta: TExporta; jsonObj: TlkJSONobject);
begin
  if (jsonObj.Field['export'].Count = 0) then
    exit;

  exporta.UFSaidaPais := jsonObj.Field['export'].Field['UFSaidaPais'].Value;
  exporta.xLocExporta := jsonObj.Field['export'].Field['xLocExporta'].Value;
  exporta.xLocDespacho := jsonObj.Field['export'].Field['xLocDespacho'].Value;

end;

procedure fulfillPayment(pag: TpagCollection; jsonObj: TlkJSONobject);
var
  pagArray: TlkJSONlist;
  pagItem: TlkJSONobject;
  item: TpagCollectionItem;
  i: Integer;
  ok: Boolean;
begin
  pagArray := jsonObj.Field['pagamento'] as TlkJSONlist;

  if (pagArray = nil) or (pagArray.Count = 0) then
    begin
      item := pag.New;
      item.tPag := TpcnFormaPagamento.fpSemPagamento;
      item.vPag := 0;
      Exit;
    end;

  for i := 0 to pagArray.Count - 1 do
  begin
    pagItem := pagArray.Child[i] as TlkJSONobject;
    item := pag.New;

    if (StrToFinNFe(ok, jsonObj.Field['ide'].Field['finNFe'].Value) <> TpcnFinalidadeNFe.fnNormal) then
    begin
      item.tPag := TpcnFormaPagamento.fpSemPagamento;
      item.vPag := 0;
    end
    else
    begin
      item.tPag := TpcnFormaPagamento.fpDuplicataMercantil;
      item.vPag := pagItem.Field['vPag'].Value;
    end;
  end;
end;

procedure fulfillCob(cobranca: TCobr; jsonObj: TlkJSONobject);
var
  pagArray: TlkJSONlist;
  pagItem: TlkJSONobject;
  item: TDupCollectionItem;
  i: Integer;
begin
  pagArray := jsonObj.Field['pagamento'] as TlkJSONlist;

  if(pagArray.Count > 0) then
  begin
    if ((pagArray.Count > 1) or (isDateInFuture(pagArray.Child[0].Field['dVenc'].Value))) then
    begin
      for i := 0 to pagArray.Count - 1 do
      begin
        pagItem := pagArray.Child[i] as TlkJSONobject;
        item := cobranca.Dup.New;
        item.nDup := FormatFloat('000', pagItem.Field['numeroParcela'].Value);
        item.dVenc := strToDate_YMD(pagItem.Field['dVenc'].Value);
        item.vDup := pagItem.Field['vPag'].Value;
      end;
    end;
  end;

  cobranca.Fat.nFat := getStringSafe(jsonObj, 'nfat');
  cobranca.Fat.vOrig := getNumberSafe(jsonObj, 'voriginal');
  cobranca.Fat.vDesc := 0;
  cobranca.Fat.vLiq := getNumberSafe(jsonObj, 'vliq');
end;

function isDateInFuture(const dateStr: string): Boolean;
var
  D: TDateTime;
  FS: TFormatSettings;
begin
  FS := TFormatSettings.Create;
  FS.DateSeparator := '-';
  FS.ShortDateFormat := 'yyyy-mm-dd';

  Result := TryStrToDate(dateStr, D, FS) and (D > Date);
end;

function strToDate_YMD(const dateStr: string): TDateTime;
var
  FS: TFormatSettings;
begin
  FS := TFormatSettings.Create;
  FS.DateSeparator := '-';
  FS.ShortDateFormat := 'yyyy-mm-dd';

  Result := StrToDate(dateStr, FS);
end;

end.
