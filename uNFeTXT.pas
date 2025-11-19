unit uNFeTXT;

interface

uses
  uFuncoes,
  ACBrNFe,
  System.SysUtils,
  ACBrDFe.Conversao,
  pcnConversaoNFe,
  StdCtrls, Forms, Windows;

function sendNFe(ACBrNFe1: TACBrNFe; var Arquivo: TextFile; RChave:String; Memo1: TMemo; SN: Boolean;
Handle: HWND; TipoEnvio: Integer = 3): Integer;

implementation

var
strxNome: string;
 Chave, vchave: string;
 VNumLote: Integer;
 VCGeraisCaminhoArquivoRetorno, loja, nfat: String;
 VWebSAmbiente: Integer;


function sendNFe(ACBrNFe1: TACBrNFe; var Arquivo: TextFile; RChave:String; Memo1: TMemo; SN: Boolean;
Handle: HWND; TipoEnvio: Integer = 3): Integer;
var
Linha, codigoBarra, unidadeTributacao, numeroparcela, FCI: string;
Pfcp, Vfcp, vliq, voriginal, VSNAliq: Real;
begin
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
          // Codigo da UF do emitente do documento fiscal
          Ide.natOp := trim(Copy(Linha, 18, 60));
          // Descricao da natureza de operacao
          // Indicador da forma de pagamento 0-Pagamento a vista 1-Pagamento a prazo 2-Outros showmessage(copy(Linha,78,1));
          if Copy(Linha, 78, 1) = '0' then
            Ide.indPag := ipVista
          else if Copy(Linha, 78, 1) = '1' then
            Ide.indPag := ipPrazo
          else
            Ide.indPag := ipOutras;
          Ide.modelo := strtoint(Copy(Linha, 79, 2));
          // Codigo do Modelo do documento fiscal
          Ide.serie := strtoint(Copy(Linha, 81, 1));
          // Serie do documento fiscal
          Ide.nNF := strtoint(Copy(Linha, 82, 9));
          // Numero do documento fiscal
          VNumLote := strtoint(Copy(Linha, 82, 9)); // numero do lote de envio
          try
            // Ide.dEmi := strtodate(Copy(Linha, 99, 2) + '/' + Copy(Linha, 96, 2) + '/' + Copy(Linha, 91, 4));
            Ide.dEmi := now;
            // Data de emissao do documento fiscal
          except
          on E: Exception do
          begin
            Memo1.Lines.Text := 'Problemas com a Data de Emissao da nota';
            Memo1.Lines.SaveToFile(VCGeraisCaminhoArquivoRetorno + '\LogErro-' +
              IntToStr(VNumLote) + '.txt');
            Application.Terminate;
          end;
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
          // Codigo do Municipio de Ocorrencia do Fato Gerador
          // Formato de Impressao do DANFE
          if (Copy(Linha, 119, 1)) = '1' then
            Ide.tpImp := tiRetrato
          else
            Ide.tpImp := tiPaisagem;

          if TipoEnvio = 3 then
            Ide.tpEmis := teNormal
          else if TipoEnvio = 4 then
            Ide.tpEmis := teDPEC;

          // Forma de emissao da NF-e
          Ide.cDV := strtoint(Copy(Linha, 121, 1));
          // Digito verificador da Chave de Acesso da NF-e

          // Identificacao do Ambiente
          if (VWebSAmbiente) = 1 then
            Ide.tpAmb := taHomologacao
          else
            Ide.tpAmb := taProducao;

          // Finalidade de emissao da NF-e
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

          if (Ide.finNFe <> fnNormal) then
          begin
            with pag.New do
            begin
              tPag := fpSemPagamento;
              vPag := 0;
            end;
          end;

          Ide.procEmi := peAplicativoContribuinte;

          // Processo de emissao da NF-e
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

          // Versao do processo de emissao da NF-e
          with Ide.NFref.New do
          begin
            if Copy(Linha, 7, 3) <> '000' then
              refNFe := Copy(Linha, 7, 44);

            if trim(Copy(Linha, 87, 3)) <> '' then
            begin
              RefECF.modelo := ECFModRef2D;
              RefECF.nECF := Copy(Linha, 87, 3);
            end;

            if trim(Copy(Linha, 90, 6)) <> '' then
              RefECF.nCOO := Copy(Linha, 90, 6);

          end;
        end
        else
          // =============================EM0203=============================//
          if (Copy(Linha, 0, 6)) = 'EM0203' then
          begin

            // Razao social obrigatoria para ambiente de homologacao
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
            // copy(Linha, 35, 60); // Razao social ou Nome do emitente
            emit.xFant := trim(Copy(Linha, 95, 60)); // Nome fantasia
            emit.EnderEmit.xLgr := trim(Copy(Linha, 155, 60)); // Logradouro
            emit.EnderEmit.nro := trim(Copy(Linha, 215, 60)); // Numero
            emit.EnderEmit.xCpl := trim(Copy(Linha, 275, 60)); // Complemento
            emit.EnderEmit.xBairro := trim(Copy(Linha, 335, 60)); // Bairro
            emit.EnderEmit.cMun := strtoint(Copy(Linha, 395, 7));
            // Codigo do municipio
            emit.EnderEmit.xMun := trim(Copy(Linha, 402, 60));
            // Nome do municipio
            emit.EnderEmit.uf := (Copy(Linha, 462, 2)); // Sigla da UF
            ACBrNFe1.Configuracoes.WebServices.uf := (Copy(Linha, 462, 2));
            emit.EnderEmit.CEP := strtoint(Copy(Linha, 464, 8));
            // Codigo do CEP
            emit.EnderEmit.cPais := strtoint(Copy(Linha, 472, 4));
            // Codigo do Pais
            emit.EnderEmit.xPais := trim(Copy(Linha, 476, 60)); // Brasil
            emit.EnderEmit.fone := trim(Copy(Linha, 536, 10)); // Telefone
            emit.ie := trim(Copy(Linha, 546, 14)); // IE
            emit.IEST := trim(Copy(Linha, 560, 18)); // IEST

            if SN then
              emit.CRT := crtSimplesNacional;

            // =========================Responsavel Tecnico=======================//
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
              // Razao social ou nome do destinatario
              dest.EnderDest.xLgr := trim(Copy(Linha, 95, 60)); // Logradouro
              dest.EnderDest.nro := trim(Copy(Linha, 155, 60)); // Numero
              dest.EnderDest.xCpl := trim(Copy(Linha, 215, 60)); // Complemento
              dest.EnderDest.xBairro := trim(Copy(Linha, 275, 60)); // Bairro
              dest.EnderDest.cMun := strtoint(Copy(Linha, 335, 7));
              // Codigo do Municipio
              dest.EnderDest.xMun := trim(Copy(Linha, 342, 60));
              // Nome do Municipio
              dest.EnderDest.uf := (Copy(Linha, 402, 2)); // Sigla da UF
              dest.EnderDest.CEP := strtoint(Copy(Linha, 404, 8));
              // Codigo do Cep
              dest.EnderDest.cPais := strtoint(Copy(Linha, 412, 4));
              // Codigo do Pais
              dest.EnderDest.xPais := trim(Copy(Linha, 416, 60)); // Brasil
              dest.EnderDest.fone := trim(Copy(Linha, 476, 10)); // Telefone
              dest.ie := trim(Copy(Linha, 486, 14)); // IE
              dest.ISUF := trim(Copy(Linha, 500, 12)); // Inscricao SUFRAMA}

              dest.email := trim(Copy(Linha, 548, 60));
              // Email para Recepcao de XML
              ACBrNFe1.NotasFiscais[0].NFe.dest.email :=
                trim(Copy(Linha, 548, 60));

              dest.idEstrangeiro := trim(Copy(Linha, 512, 20));
              // Inscricao de Estrangeiro}

              case strtoint(Copy(Linha, 532, 1)) of
                1:
                  dest.indIEDest := inContribuinte;
                2:
                  dest.indIEDest := inIsento;
                9:
                  dest.indIEDest := inNaoContribuinte;
              end; // Indica Contribuinte

              dest.IM := trim(Copy(Linha, 533, 15)); // Inscricao municiapl
            end; // fim do EM0204

      // ==============================EM0205=============================//
      // =Endereco de Entrega

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

        with Det.New do
        begin
          Prod.nItem := strtoint(Copy(Linha, 9, 3));
          // Nro. do item
          Prod.cProd := trim(Copy(Linha, 12, 60));
          // Codigo do Produto ou servico
{$REGION 'Codigo Barra'}
          codigoBarra := trim(Copy(Linha, 72, 14)); // cEAN

          if trim(codigoBarra) = '' then
            codigoBarra := 'SEM GTIN';

          Prod.cEAN := trim(codigoBarra); // cEAN
{$ENDREGION}
          Prod.xProd := trim(Copy(Linha, 86, 120));
          // Descricao do produto ou servico
          Prod.NCM := trim(Copy(Linha, 206, 8)); // Codigo NCM
          Prod.EXTIPI := trim(Copy(Linha, 214, 3)); // EX_TIPI
          Prod.CFOP := trim(Copy(Linha, 219, 4));
          // Prod.CEST := trim(Copy(Linha, 219, 4));
          // Codigo fiscal da operacao

          if trim(Copy(Linha, 223, 6)) <> '' then
            Prod.uCom := trim(Copy(Linha, 223, 6)) // Unidade comercial
          else
            Prod.uCom := '0';

          Prod.qCom := StrToFloat(StringReplace((Copy(Linha, 229, 15)), '.',
            ',', [rfReplaceAll]));
          // Quantidade comercial
          Prod.vUnCom := StrToFloat(StringReplace((Copy(Linha, 244, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor unitario de comercializacao
          Prod.vProd := StrToFloat(StringReplace((Copy(Linha, 259, 15)), '.',
            ',', [rfReplaceAll]));
          // Valor Total Bruto dos Produtos ou Servicos
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
          // Valor Unitario de tributacao
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
          // FCI - Numero de controle da FCI - Ficha de Conteudo de Importacao

          /// //// Imposto de importacao

          if length(Linha) > 385 then
          begin
            if trim(Copy(Linha, 384, 15)) <> '' then
              Imposto.II.vBC := StrToFloat(StringReplace((Copy(Linha, 384, 15)),
                '.', ',', [rfReplaceAll]));
            // Valor da Base de Calculo para II
            if trim(Copy(Linha, 399, 15)) <> '' then
              Imposto.II.vII := StrToFloat(StringReplace((Copy(Linha, 399, 15)),
                '.', ',', [rfReplaceAll]));
            // Valor do Imposto de importacao

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
                if CharInset(Prod.CFOP[1], ['3']) then
                begin
                  with Prod.DI.New do
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
                      if CharInset(Prod.CFOP[1], ['3']) then
                        with adi.New do
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
              with Prod.rastro.New do
              // quando usar Medicamentos usa Prod.med, nosso caso e so rastreamento entao Pro.rastro
              begin
                nLote := trim(Copy(Linha, 7, 20));
                qLote := StrToFloat(StringReplace(trim(Copy(Linha, 27, 15)),
                  '.', ',', [rfReplaceAll]));
                dFab := strtodate(Copy(Linha, 42, 10));
                dVal := strtodate(Copy(Linha, 52, 10));
              end;

              Readln(Arquivo, Linha); // Ler a proxima linha
            end;
          end;

          if (Copy(Linha, 0, 6)) = 'EM4207' then
          begin
            while (Copy(Linha, 0, 6)) = 'EM4207' do
            begin
              with Prod.med.New do
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
                  'Origem do produto nao encontrada. Verifique e tente novamente.',
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
                  'Situacao tributaria do ICMS nao encontrada. Verifique e tente novamente.',
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
              // Modalidade de determinacao da BC do ICMS ST
              Imposto.ICMS.pRedBC :=
                StrToFloat(StringReplace((Copy(Linha, 17, 15)), '.', ',',
                [rfReplaceAll]));
              // Percential de reducao de BC do ICMS

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
                  'Origem do produto nao encontrada. Verifique e tente novamente.',
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
                  'Situacao tributaria do ICMS nao encontrada. Verifique e tente novamente.',
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
              // Modalidade de determinacao da BC do ICMS ST
              Imposto.ICMS.pRedBC :=
                StrToFloat(StringReplace((Copy(Linha, 17, 15)), '.', ',',
                [rfReplaceAll]));
              // Percential de reducao de BC do ICMS
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
            // Percentual do ICMS relativo ao Fundo de Combate a Pobreza (FCP) na UF de destino

            // Percentual do Fundo de Combate a Pobreza (FCP)
            // Imposto.ICMS.pFCP :=StrToFloat(StringReplace((Copy(Linha, 22, 15)), '.', ',', [rfReplaceAll]));
            // Valor do Fundo de Combate a Pobreza (FCP)
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

            Imposto.ICMSUFDest.vBCUFDest :=
              StrToFloat(StringReplace((Copy(Linha, 7, 15)), '.', ',',
              [rfReplaceAll])); // Valor da BC do ICMS na UF de destino

            Imposto.ICMSUFDest.vBCFCPUFDest := Imposto.ICMSUFDest.vBCUFDest;

            Imposto.ICMSUFDest.pICMSUFDest :=
              StrToFloat(StringReplace((Copy(Linha, 37, 15)), '.', ',',
              [rfReplaceAll])); // Aliquota interna da UF de destino

            Imposto.ICMSUFDest.pICMSInter :=
              StrToFloat(StringReplace((Copy(Linha, 52, 15)), '.', ',',
              [rfReplaceAll])); // Aliquota interestadual das UF envolvidas

            Imposto.ICMSUFDest.pICMSInterPart :=
              StrToFloat(StringReplace((Copy(Linha, 67, 15)), '.', ',',
              [rfReplaceAll]));

            Imposto.ICMSUFDest.vICMSUFDest :=
              StrToFloat(StringReplace((Copy(Linha, 97, 15)), '.', ',',
              [rfReplaceAll]));

            Imposto.ICMSUFDest.vICMSUFRemet :=
              StrToFloat(StringReplace((Copy(Linha, 112, 15)), '.', ',',
              [rfReplaceAll]));

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

            // CASE   Situacao tributaria do IPI
            try
              strtoint(Copy(Linha, 50, 2))
            except
              messagebox(Handle,
                'Situacao de IPI nao encontrada. Verifique e tente novamente.',
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

            // CASEEEEE 02 Situacao Tributaria do PIS       showmessage(copy(EM0209,12,2));
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
            // CASE Situacao Tributaria do COFINS       showmessage(copy(EM0209,52,2));

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
        // Valor Total dos produtos e servicos  showmessage(copy(linha,68,15));
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
        // Outras Despesas Acessorias showmessage(copy(linha,188,15));
        Total.ICMSTot.vOutro := StrToFloat(StringReplace((Copy(Linha, 187, 15)),
          '.', ',', [rfReplaceAll]));
        // Valor Total da NFe       showmessage(copy(linha,203,15));
        Total.ICMSTot.vNF := StrToFloat(StringReplace((Copy(Linha, 202, 15)),
          '.', ',', [rfReplaceAll]));

        // Informar a soma no cabecalho da nota
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

          Total.ICMSTot.vICMSUFDest :=
            StrToFloat(StringReplace((Copy(Linha, 262, 15)), '.', ',',
            [rfReplaceAll]));

          Total.ICMSTot.vICMSUFRemet :=
            StrToFloat(StringReplace((Copy(Linha, 277, 15)), '.', ',',
            [rfReplaceAll]));

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
        // Razao social ou nome                     showmessage(copy(linha,36,60));
        Transp.Transporta.ie := (Copy(Linha, 96, 14));
        // IE                                                showmessage(copy(linha,96,14));
        Transp.Transporta.xEnder := (Copy(Linha, 110, 60));
        // Endereco completo                     showmessage(copy(linha,110,60));
        Transp.Transporta.xMun := (Copy(Linha, 170, 60));
        // Nome do Municipio                     showmessage(copy(linha,170,60));
        Transp.Transporta.uf := (Copy(Linha, 230, 2));
        // Sigla da UF                     showmessage(copy(linha,230,2));

        with Transp.Vol.New do
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

      end; // fim do EM0211

      // =============================EM1211=============================//
      if (Copy(Linha, 0, 6)) = 'EM1211' then
      begin
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
          with pag.New do
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
          with pag.New do
          begin
            tPag := fpSemPagamento;
            vPag := 0;
          end;
        end
        else
        begin
          with pag.New do
          begin
            tPag := fpDuplicataMercantil; // fpDuplicata
            vPag := StrToFloat(StringReplace((Copy(Linha, 77, 15)), '.', ',',
              [rfReplaceAll]));
          end;
        end;

{$ENDREGION}
        if ((strtodate(Copy(Linha, 75, 2) + '/' + Copy(Linha, 72, 2) + '/' +
          Copy(Linha, 67, 4)) > Date) or (strtoIntDef(numeroparcela, 0) > 1))
        then
        begin
          with Cobr.Dup.New do
          begin
            numeroparcela := (Copy(Linha, 7, 60));
            nDup := LeftZero((Copy(numeroparcela, pos('-', numeroparcela) + 1,
              length(numeroparcela))), 3);
            // Numero da fatura
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
    end; // enquanto nao chegar ao fim do Arquivo
    // =============================EM0214=============================//
    if (Copy(Linha, 0, 6)) = 'EM0214' then
    begin

      InfAdic.infCpl := (Copy(Linha, 7, 2000));
    end; // fim do EM0214
  end; // Fim das configuracoes da NFe
end;

end.
