#include "protheus.ch"
#Include "TOTVS.ch"
#Include "REPORT.ch"

#define DMPAPER_A4 9 

/*/{Protheus.doc} DVDREL01
//TODO GERA RELATÓRIO DE NOTAS CANCELADAS CONFORME PERIODO INFORMADO
@description RELATÓRIO DE NOTAS FISCAIS CANCELADAS
@author David Rodrigues
@since 17/04/2019
@version 1.0
@type function
@history 20/05/2020,David Rodrigues, Ajuste para ser chamado em outras funções
@history 09/09/2020,David Rodrigues, Adicionado a coluna tipo
@history 24/07/2021,David Rodrigues, Ajustes para gerar planilha excel em chamadas externa
@history 04/03/2022,David Rodrigues, Ajuste pq estava puxando data no pergunte errado
/*/

//|------ ARRAY aXparam --------------------|
//| aXparam[1] // Data de 					|
//| aXparam[2] // Data ate 					|
//| aXparam[3] // Diretorio onde salvar		|
//| aXparam[4] // Nome do arquivo			|
//|-----------------------------------------|

User Function DVDREL01(xExcAuto,aXparam)
	
	Local oReport := nil
	
	Private lExcJob := Iif( xExcAuto == "DVDFFT02", .T., .F.)

	DEFAULT xExcAuto 	:= ''
	DEFAULT aXparam		:= {}

	oReport := ReportDef()

	//Chamado Externa
	If lExcJob
		If EMPTY(aXparam)
			Alert("aXparam Não pode ser vazio!","DVDREL01-DVDFFT02 VALID")
		Else
			ZCFGEXC01(aXparam)
			ZGERAPLAN(aXparam)
			ZGERAPDF(oReport,aXparam)
			
		EndIf
	else
		oReport:PrintDialog()
	EndIf
	

return()

Static Function ReportDef()
	
	Local oReport   := Nil
	Local oSection1 := Nil
	Local cNome 	:= "DVDREL01" 					//Nome físico do relatório
	Local cTitulo 	:= "Notas Canceladas" 			//Titulo do Relario
	Local cDescricao:= "Imprime a Relação de NF Canceladas"	//Descrição do relatório
	Local cPerg	    := PadR('DVDREL01',10)			//Pergunta utilizada em parâmetros

	oReport:=TReport():New(cNome,cTitulo,cPerg,{|oReport| ReportPrint( oReport ) }, cDescricao)

	//Se nao for chamado por job chama a pergunta padrão
	If !lExcJob
		Pergunte(cPerg,.F.)
	endif
	
	//Configurações do relatorio
	oReport:lParamPage := .F. 	//Se Imprime a pagina de parametros
	oReport:SetPortrait() 		//Relatorio em retrato
	oReport:LoadDefault()		//Layout Padão
	oReport:SetEdit(.F.)		//Se permite usuario alterar o layout
	oReport:SetTotalInLine(.T.) //Totalizadores em linha ou coluna
	oReport:SetLeftMargin(2)    //Define a margem à esquerda do relatório
	
	//primeira seção
	oSection1:= TRSection():New(oReport,"NFCANC", {"SF3"}, NIL, , .T.)
	TRCell():New(oSection1, "F3_DTCANC"  ,"SF3","DT. CANCEL",,12,,,"CENTER",,"CENTER",,,,,,)
	TRCell():New(oSection1, "F3_NFISCAL" ,"SF3","NUMERO",,10,,,,,,,,,,,)
	TRCell():New(oSection1, "F3_ESPECIE" ,"SF3","ESPECIE",,8,,,,,,,,,,,)
	TRCell():New(oSection1, "F3_TIPO"    ,"SF3","TIPO",,8,,,,,,,,,,,)
	TRCell():New(oSection1, "F3_CHVNFE"  ,"SF3","CHAVE",,50,,,,,,,,,,,)
	TRCell():New(oSection1, "F3_DESCRET" ,"SF3","RET.SEFAZ",,100,,,"LEFT",,"LEFT",,,.T.,,,)

	//Criar totalizador Com total de nfs canceladas
	TRFunction():New(oSection1:Cell("F3_DTCANC"),"TOTNFCANC" ,"COUNT",,"TOTAL NF CANCELADAS","@E 999999",,.F.,.T.)	

Return(oReport)

Static Function ReportPrint(oReport)

	Local oSection1 := oReport:Section(1)
	Local cAlias    := GetNextAlias()
	
	BeginSql alias cAlias

		SELECT F3_DTCANC,F3_NFISCAL,F3_ESPECIE,F3_TIPO,F3_CHVNFE,F3_DESCRET
		FROM %table:SF3% SF3
		WHERE F3_FILIAL = %xfilial:SF3%
		AND F3_DTCANC >= %exp:MV_PAR01%
		AND F3_DTCANC <= %exp:MV_PAR02%
		AND F3_DTCANC <> %exp:' '%
		AND SF3.%notDel%
		
		ORDER BY F3_DTCANC
	EndSql

	dbSelectArea(cAlias)
	(cAlias)->(dbGoTop())

	oReport:SetMeter((cAlias)->(LastRec()))

	oSection1:Init() //inicializa a primeira seção

	While !(cAlias)->( EOF() )
		
		If oReport:Cancel()
			Exit
		EndIf
		
		oReport:SetMsgPrint("Imprimindo " + alltrim( (cAlias)->F3_NFISCAL))
		oReport:IncMeter()

		//Associa os valores com as celulas
		oSection1:Cell("F3_DTCANC"   ):SetValue(STOD((cAlias)->F3_DTCANC ))
		
		oSection1:Cell("F3_NFISCAL"  ):SetValue((cAlias)->F3_NFISCAL)
		
		oSection1:Cell("F3_ESPECIE"):SetValue((cAlias)->F3_ESPECIE)

		oSection1:Cell("F3_TIPO"):SetValue((cAlias)->F3_TIPO)
		
		oSection1:Cell("F3_CHVNFE"):SetValue((cAlias)->F3_CHVNFE )
		
		oSection1:Cell("F3_DESCRET"):SetValue((cAlias)->F3_DESCRET)

		oSection1:Printline() //imprime a sessão

		(cAlias)->(dbSkip())

	EndDo
	
	(cAlias)->(dbCloseArea())
	
	oSection1:Finish() //finaliza a seção

Return( NIL )

/*/{Protheus.doc} ZCFGEXC01
	(long_description)
	@type  Static Function
	@author user
	@since 18/05/2020
	@version version
	@param param_name, param_type, param_descr
	@return return_var, return_type, return_description
	@example
	(examples)
	@see (links_or_references)
	/*/
Static Function ZCFGEXC01(aXparam)


	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Variaveis utilizadas para parametros                         ³
	//³ mv_par01 // Data de                                          ³
	//³ mv_par02 // Data ate                                         ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

	MV_PAR01 := aXparam[1]
	MV_PAR02 := aXparam[2]

Return Nil 

/*/{Protheus.doc} 
	(long_description)
	@type  Static Function
	@author user
	@since 18/05/2020
	@version version
	@param param_name, param_type, param_descr
	@return return_var, return_type, return_description
	@example
	(examples)
	@see (links_or_references)
	/*/
Static Function ZGERAPDF(oReport,aXparam)

	Local oPrint 			//Objeto FWMSPrinter
	Local cFilePDF

	oReport:nDevice 			:= 6                	// Geração formato .PDF
	oReport:nEnvironment		:= 1					//Ambiente selecionado. 1-Server e 2-Cliente.
	oReport:cDir	 			:= Alltrim(aXparam[3])	//Diretorio do arquivo
	oReport:cFile	 			:= Alltrim(aXparam[4])	//Nome do arquivo
	oReport:lParamPage  		:= .F. 					//Aponta que serão impressos os parâmetros do relatório
	
	oReport:LoadDefault()				//Layout Padão
	//oReport:oPage:setPaperSize(9) 		//Tamanho do papel 9=> A4
    //oReport:cFontBody  := "Courier new" //Fonte padrão
	//oReport:nFontBody  := 8 			 //Tam Fonte
	//oReport:nLineHeight := 33			 //Altura das Linhas


	//Inicia o Treport
	oReport:init()
	
	//Recebe objeto FWMSPrinter
	oPrint := oReport:oPrint
	
	//Define o Diretorio de impressão do PDF
	oPrint:cPathPDF 	:= Alltrim(cDir)
	oPrint:lViewPDF 	:= .F.
	//oPrint:SetParm( "-RFS") 	//diminui a divergência nos tamanhos de fonte
	
	//Verifica se o diretorio existe
	If !ExistDir( oPrint:cPathPDF )
		//Cria O diretorio se não existir
		MAKEDIR(oPrint:cPathPDF)
	EndIf
	
	//Gera o Local onde esta arquivo pdf
	cFilePDF := oPrint:cPathPDF + oReport:cFile + ".pdf"	
	
	//Apaga aquivo Para não gerar Relatorio antigo
	FERASE(cFilePDF)

	//Gera o relatorio
	ReportPrint(oReport)

	//Finaliza Treport
	oReport:Finish( )
	
	//Gera o arquivo Pdf
	oPrint:Print(.F.)

Return Nil

/*/Função gera relatorio formato planilha
	@type  Function
	@author David Rodrigues
	@since 14/07/2021
	/*/
Static Function ZGERAPLAN(aXparam)

	Local oReport 	:= ReportDef() //Defini novamente para evitar erro
	Local cFilePLAN	:= ''
	
	oReport:cDir	 			:= aXparam[3]			//Diretorio onde salvar
	oReport:cFile	 			:= Alltrim(aXparam[4])	//Nome do arquivo
	oReport:lParamPage  		:= .F. 					//Aponta que serão impressos os parâmetros do relatório
	oReport:nRemoteType 		:= 3  		        	// FORMA DE GERAÇÃO DO RELATÓRIO

	oReport:setFile(cDir+Alltrim(aXparam[4])) //Diretorio + nome do arquivo
	oReport:setEnvironment(1)						//Ambiente selecionado. 1-Server e 2-Cliente.
	oReport:setDevice(4)							// Geração formato .XLS .XML
	oReport:NEXCELPRINTTYPE := 2					
	oReport:NXMLOPENWITH	:= 0

	oReport:SetPreview(.F.)

	//Gera o Local onde esta arquivo .xls
	cFilePlan := cDir+cFile+'.xls'	
	
	//Apaga aquivo Para não gerar Planilha antiga
	FERASE(cFilePlan)

	//Inicia o Treport
	oReport:oReport:init()

	//Gera o relatorio
	ReportPrint(oReport)

	//Finaliza Treport e gerar arquivo .tmp 
	oReport:Finish( )

	//renomeia arquivo .tmp para .xls
	FRename( oReport:cFile, cFilePlan )

Return Nil
