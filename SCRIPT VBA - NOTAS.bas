Attribute VB_Name = "Módulo1"
Sub abre_notas_sap()
    ' declaração da variaveis
    Dim SAPGuiAuto As Object
    Dim SAPApp As Object
    Dim Connection As Object
    Dim Session As Object

    Dim ws As Worksheet
    Dim ult_linha As Long
    Dim linha As Long
    
    Dim campos As Variant
    Dim j As Integer
    Dim i As Integer
    
    Dim titulo As String
    Dim tipo_nota As String
    Dim embarcacao As String
    Dim equipamento As String
    Dim prioridade As String
    Dim descricao As String
    Dim horimetro As String
    Dim centro_localizador As String
    Dim area_operacional As String
    Dim origem As String
    Dim centro_custo As String
    Dim natureza_demanda As String
    Dim quilometragem As String
    Dim nota_sap As String
    Dim notificador As String
    Dim job_code As String
    Dim comp_code As String
    Dim pessoa_contato As String

    

    ' define a planilha
    planilha_main = "Notas"
    Set ws = ThisWorkbook.Sheets(planilha_main)

    ' busca a última linha preenchida
    ult_linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' coisa do sap, não sei oq é, mas rodou com isso
    On Error Resume Next
    Set SAPGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SAPGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set Session = Connection.Children(0)
    On Error GoTo 0


    Session.findById("wnd[0]").maximize

    ' loop para preencher os dados no SAP
    For i = 2 To ult_linha ' começa da linha 2 para ignorar cabeçalhos
    

        ' verifica se a célula na Coluna A está vazia, se estiver, o if é true e vai rodar o código do sap
        If Trim(ws.Cells(i, 1).Value) = "" Then
        
            tipo_nota = Sheets(planilha_main).Cells(i, 2).Value
            titulo = Sheets(planilha_main).Cells(i, 3).Value
            embarcacao = Sheets(planilha_main).Cells(i, 14).Value
            equipamento = Sheets(planilha_main).Cells(i, 4).Value
            prioridade = Sheets(planilha_main).Cells(i, 7).Value
            descricao = Sheets(planilha_main).Cells(i, 11).Value
            horimetro = Sheets(planilha_main).Cells(i, 9).Value
            centro_localizador = Sheets(planilha_main).Cells(i, 12).Value
            area_operacional = Sheets(planilha_main).Cells(i, 13).Value
            origem = Sheets(planilha_main).Cells(i, 10).Value
            centro_custo = Sheets(planilha_main).Cells(i, 15).Value
            natureza_demanda = Sheets(planilha_main).Cells(i, 18).Value
            comp_code = Sheets(planilha_main).Cells(i, 17).Value
            job_code = Sheets(planilha_main).Cells(i, 16).Value
            notificador = Sheets(planilha_main).Cells(i, 8).Value
            pessoa_contato = Sheets(planilha_main).Cells(i, 19).Value
            
            
            ' verificação para saber se tem célula vazia.
                colunas = Array(2, 3, 13, 4, 7, 11, 9, 12, 14, 10, 15, 18, 8)
               
                For j = LBound(colunas) To UBound(colunas)
                    If Trim(Sheets(planilha_main).Cells(i, colunas(j)).Value) = "" Then
                        MsgBox "Existe uma célula vazia na linha " & i & ". Verifique e tente novamente.", vbExclamation, "Erro"
                        Exit Sub
                    End If
                Next j
                                            
        
                ' Em tipo da nota ele coloca a inform~ção da coluna b
                On Error Resume Next ' Tratamento de erro caso o usuário rode o script sem estar na tela inicial da função iw51 no sap
                Session.findById("wnd[0]/usr/ctxtRIWO00-QMART").Text = tipo_nota
                If Err.Number <> 0 Then
                    MsgBox "Erro ao executar o script. Coloque o SAP no menu inicial da IW51 e tente novamente.", vbExclamation, "ATENÇÃO"
                    Err.Clear
                    Exit Sub
                End If
                On Error GoTo 0
                Session.findById("wnd[0]/usr/ctxtRIWO00-QMART").caretPosition = 2
                Session.findById("wnd[0]").sendVKey 0
                
                
                ' acessa a aba para criar a notificação.
                Session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-QMTXT").Text = titulo ' preenche o titulo da nota
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:1200/ctxtRIWO1-EQUNR").Text = equipamento ' preenchen o campo equipamento
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:1200/ctxtRIWO1-EQUNR").SetFocus
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:1200/ctxtRIWO1-EQUNR").caretPosition = 8
                'Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7540/cmbVIQMEL-PRIOK").Key = prioridade ' preenche o campo prioreidade
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:1200/ctxtRIWO1-EQUNR").SetFocus
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:1200/ctxtRIWO1-EQUNR").caretPosition = 8
                
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7515/ctxtVIQMEL-QMNAM").Text = UCase(notificador) 'colcoa o nome do notificador(pessoa que está abrindo a nota)
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7515/ctxtVIQMEL-QMNAM").caretPosition = 5
                
                Session.findById("wnd[0]").sendVKey 0
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02").Select
                
               '-- ABA REFERÊNCIA
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 9, 9
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/txtQMEL-YYHORIMETRO").Text = horimetro 'preenche o campo horímetro
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/txtQMEL-YYHORIMETRO").SetFocus

                
                
                'colocar o rpcoding no segundo quadrado em forma de número
                Session.findById("wnd[0]").maximize
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/ctxtVIQMEL-QMCOD").Text = origem 'preenche o campo rapaircoding
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/ctxtVIQMEL-QMCOD").SetFocus
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/ctxtVIQMEL-QMCOD").caretPosition = 4
                Session.findById("wnd[0]").sendVKey 0 ' enter para carregar as informações com base no centro localizador de repaircoding
                
                ' coloca a descrição
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/cntlTEXT/shellcont/shell").Text = descricao + vbCr + "" 'preenche o campo descrição
            
                
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/txtQMEL-YYHORIMETRO").caretPosition = 4
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06").Select
                

                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/ctxtILOA-SWERK").Text = centro_localizador '´preenche o campo centro localizador
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/ctxtILOA-BEBER").Text = area_operacional ' preenche o campo area operacional
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/txtILOA-EQFNR").Text = embarcacao 'preenche o campo ordenação
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/ctxtILOA-VKORG").Text = "3000" 'origem de vendas é fixo pelo código
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/ctxtILOA-VTWEG").Text = "04" ' canal distrib é fixo pelo código
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/ctxtILOA-SPART").Text = "PM" 'setor atividade é fixo pelo código
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/ctxtILOA-KOSTL").Text = centro_custo ' preenche o campo centro de custo
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/txtILOA-EQFNR").SetFocus
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB06/ssubSUB_GROUP_10:SAPLIQS0:7314/subILOA:SAPMILA0:7000/txtILOA-EQFNR").caretPosition = 12
                
                ' bloco para acessar "Ação sugerida" e colocar job code e component code
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").Select
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtQMEL-ZZJOB_CODE").Text = job_code ' preenche o campo job code
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtQMEL-ZZCOMPONENT_CODE").Text = comp_code ' preenche o campo comp code
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtQMEL-ZZCOMPONENT_CODE").SetFocus
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtQMEL-ZZCOMPONENT_CODE").caretPosition = 4
                Session.findById("wnd[0]").sendVKey 0
                
                ' acessa ampliação para preencher natureza da demanda
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21").Select
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0105/ctxtQMEL-ZZNATUREZA").Text = natureza_demanda ' preenche o campo natureza da demanda
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0105/ctxtQMEL-ZZNATUREZA").caretPosition = 1
                Session.findById("wnd[0]").sendVKey 0
                
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0105/txtQMEL-ZZKM").Text = 1 'quilometragem ' preenche o campo quilometragem
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0105/txtQMEL-ZZKM").SetFocus
                Session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0105/txtQMEL-ZZKM").caretPosition = 1
                
                
                
                'ABA PESSOA DE CONTATO. duas condicionais para verificar a pessoa de contato
                
                ' aqui o código só executa se houver alguma informação na célula de pessoa de contato
                If pessoa_contato <> "" Then
                    ' tratamento específico do split. verifica se o usuário preencheu corretamente a célula da coluna 'pessoa de contato'
                    On Error Resume Next
                    partes = Split(pessoa_contato, " ") ' divide o texto da celula em duas partes, no caso, nome e sobrenome
                    If Err.Number <> 0 Or UBound(partes) < 1 Then
                        MsgBox "Erro no preenchimento da coluna 'Pessoa de Contato'. Verifique e tente novamente!", vbExclamation, "ERRO" ' emite uma mensagem de orientação ao usuário
                        Err.Clear
                        Exit Sub ' para o código se o if true
                    End If
                    On Error GoTo 0
                
                    primeiro_nome = partes(0) ' armazena a primeira parte da célula aqui
                    sobrenome = partes(1) ' armazena a segunda parte da célula aqui
                    
                    'bloco para acessar a aba de pessoa de contato
                    Session.findById("wnd[0]").maximize
                    Session.findById("wnd[0]/tbar[1]/btn[5]").press
                    Session.findById("wnd[0]/usr/tblSAPLIPARTCTRL_0200/cmbIHPA-PARVW[0,4]").Key = "PC"
                    Session.findById("wnd[0]/usr/tblSAPLIPARTCTRL_0200/ctxtDIADR-NAME_LIST[2,4]").SetFocus
                    Session.findById("wnd[0]/usr/tblSAPLIPARTCTRL_0200/ctxtDIADR-NAME_LIST[2,4]").caretPosition = 0
                    Session.findById("wnd[0]").sendVKey 4
                    Session.findById("wnd[1]/tbar[0]/btn[17]").press
                    
                    
                   On Error Resume Next
                    ' tenta preencher os campos de busca com o nome e sobrenome
                    Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").Text = sobrenome
                    Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[3,24]").Text = primeiro_nome
                    Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    
                    ' Verifica se encontrou o nome
                    ' se encotrou o nome, ele pula o bloco do if e continua rodando... caso contrário, ele entra no if e para o código
                    If Session.findById("wnd[1]/usr/lbl[64,3]", False) Is Nothing Then
                        MsgBox "Pessoa de contato não encontrada. Verifique se digitou corretamente.", vbExclamation, "ATENÇÃO"
                        Exit Sub
                        
                    Else
                        ' Continua o código normalmente
                        Session.findById("wnd[1]/usr/lbl[64,3]").SetFocus
                        Session.findById("wnd[1]/usr/lbl[64,3]").caretPosition = 6
                        Session.findById("wnd[1]").sendVKey 2
                    End If
                    
                    On Error GoTo 0 ' Volta o tratamento de erro ao normal
                   
                    Session.findById("wnd[0]").sendVKey 0
                    Session.findById("wnd[0]").sendVKey 0
                               
                Else 'Caso 2 -executa caso a celula da coluna pessoa de contato esteja vazia.
                
                
                    ' aqui ele vai executar pegando uma pessoa de contato qualquer.
                    Session.findById("wnd[0]").maximize
                    Session.findById("wnd[0]/tbar[1]/btn[5]").press
                    Session.findById("wnd[0]/usr/tblSAPLIPARTCTRL_0200/cmbIHPA-PARVW[0,4]").Key = "PC"
                    Session.findById("wnd[0]/usr/tblSAPLIPARTCTRL_0200/ctxtDIADR-NAME_LIST[2,4]").SetFocus
                    Session.findById("wnd[0]/usr/tblSAPLIPARTCTRL_0200/ctxtDIADR-NAME_LIST[2,4]").caretPosition = 0
                    Session.findById("wnd[0]").sendVKey 4
                    Session.findById("wnd[1]").sendVKey 0
                    Session.findById("wnd[0]").sendVKey 0
                    Session.findById("wnd[0]/tbar[0]/btn[3]").press
                
                End If
                    
                ' bloco para pegar no número da nota gerada e colocar na planilha após criar a nota
                Session.findById("wnd[0]/tbar[0]/btn[3]").press
                nota_sap = Session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-QMNUM").Text
                Sheets(planilha_main).Cells(i, 1).Value = nota_sap 'cola na planilha o número da nota armazenada na variavel
                Session.findById("wnd[0]/tbar[0]/btn[11]").press
                
        End If
        
    Next i

    MsgBox "Processo concluído!", vbInformation ' mensagem emitida após finalizar de rodar todas as notas com sucesso

End Sub

Sub aba_equipamento()

Dim num_serie As String
Dim equipamento As String
Dim cliente As String
Dim num_cliente As String
Dim campos As Variant
Dim coluna As Variant

    ' define a planilha
    planilha_equipamento = "Equipamento"
    Set ws = ThisWorkbook.Sheets(planilha_equipamento)

    ' busca a última linha preenchida
    ult_linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' coisa do sap, não sei oq é, mas rodou com isso
    On Error Resume Next
    Set SAPGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SAPGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set Session = Connection.Children(0)
    On Error GoTo 0
       
    For i = 2 To ult_linha
    
        campos = Array(2, 3, 4)
        vazio = True
        For Each coluna In campos
            If Trim(Sheets(planilha_equipamento).Cells(i, coluna).Value) <> "" Then
                vazio = False
                Exit For
            End If
        Next coluna
                       
                If vazio Then ' pronto para automação no sap
                    num_serie = Sheets(planilha_equipamento).Cells(i, 1).Value
                    Session.findById("wnd[0]").maximize
                    Session.findById("wnd[0]/usr/txtSERNR-LOW").Text = num_serie
                    Session.findById("wnd[0]/usr/txtSERNR-LOW").SetFocus
                    Session.findById("wnd[0]/usr/txtSERNR-LOW").caretPosition = 8
                    Session.findById("wnd[0]/tbar[1]/btn[8]").press
                    Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\05").Select
                    equipamento = Session.findById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152A:SAPLITO0:1520/ctxtITOBATTR-EQUNR").Text
                    num_cliente = Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\05/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102A:SAPLITO0:1092/subSUB_1092A:SAPLIPAR:0201/tblSAPLIPARTCTRL_0200/ctxtIHPA-PARNR[1,0]").Text
                    cliente = Session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\05/ssubSUB_DATA:SAPLITO0:0102/subSUB_0102A:SAPLITO0:1092/subSUB_1092A:SAPLIPAR:0201/tblSAPLIPARTCTRL_0200/ctxtDIADR-NAME_LIST[2,0]").Text
                    Sheets(planilha_equipamento).Cells(i, 2).Value = equipamento
                    Sheets(planilha_equipamento).Cells(i, 3).Value = cliente
                    Sheets(planilha_equipamento).Cells(i, 4).Value = num_cliente
                    Session.findById("wnd[0]").maximize
                    Session.findById("wnd[0]/tbar[0]/btn[3]").press
                    Session.findById("wnd[0]/usr/txtSERNR-LOW").Text = ""
                    Session.findById("wnd[0]/usr/txtSERNR-LOW").SetFocus
                    Session.findById("wnd[0]/usr/txtSERNR-LOW").caretPosition = 0
                End If
                
            
    Next i
   MsgBox "Processo concluído!", vbInformation 'msg para informar que o código rodou
End Sub
