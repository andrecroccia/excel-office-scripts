// #######################################################################################
// SCRIPT: Processar_Importacao_GETIN
// Final: Adapta a importação GET IN para a Base Geral (Tabela1).
// Filtra: Apenas notas <= 6 na coluna "4ª Resposta".
// NOVO: Concatena Feedback Geral + Garçom, remove '55' do telefone, insere fórmulas de data.
// #######################################################################################

function main(workbook: ExcelScript.Workbook) {

    // #######################################################################
    // !!! CONFIGURAÇÕES CRÍTICAS E VALORES CONFIRMADOS !!!
    const NUMERO_TOTAL_COLUNAS_TABELA1 = 17;
    const PLANILHA_IMPORTACAO_NOME = "IMPORTAÇÃO_GETIN";
    const PLANILHA_DESTINO_NOME = "Base Geral";
    const TABELA_DESTINO_NOME = "Tabela1";
    // #######################################################################

    // 1. Definições e Verificações Iniciais
    const planilhaImportacao = workbook.getWorksheet(PLANILHA_IMPORTACAO_NOME);
    const planilhaDestino = workbook.getWorksheet(PLANILHA_DESTINO_NOME);

    if (!planilhaImportacao || !planilhaDestino) {
        console.log(`ERRO: Uma ou mais planilhas ('${PLANILHA_IMPORTACAO_NOME}', '${PLANILHA_DESTINO_NOME}') não foram encontradas.`);
        return;
    }

    const tabelaDestino = planilhaDestino.getTable(TABELA_DESTINO_NOME);
    if (!tabelaDestino) {
        console.log(`ERRO: A tabela '${TABELA_DESTINO_NOME}' não foi encontrada na planilha '${PLANILHA_DESTINO_NOME}'.`);
        return;
    }

    // Leitura dos dados da planilha de importação
    const usedRange = planilhaImportacao.getUsedRange();
    if (!usedRange || usedRange.getRowCount() <= 1) {
        console.log(`Planilha '${PLANILHA_IMPORTACAO_NOME}' está vazia ou só tem cabeçalho. Limpando...`);
        if (usedRange) {
            planilhaImportacao.getRange("A2:Z1000").clear(ExcelScript.ClearAppliedTo.contents);
        }
        return;
    }

    // Leitura dos valores
    const data = planilhaImportacao.getUsedRange().getValues();
    const headers = data[0] as string[];

    // Mapeamento dos índices das colunas de origem (GET IN)
    const colIndex: { [key: string]: number } = {
        data_avaliacao: headers.indexOf("Data e Hora"),
        loja: headers.indexOf("Unidade"),
        nota: headers.indexOf("4ª Resposta"), // CRÍTICO: Coluna de filtro
        marca: headers.indexOf("Nome"),
        telefone: headers.indexOf("Telefone"), // CRÍTICO: Nova coluna para WhatsApp
        feedback_geral: headers.indexOf("Feedback Geral"), // CRÍTICO: Parte 1 do Comentário
        feedback_garcom: headers.indexOf("Feedback do Garçom") // CRÍTICO: Parte 2 do Comentário
    };

    // Verificação de Colunas CRÍTICAS
    const requiredColumns = ["data_avaliacao", "loja", "nota", "feedback_geral", "feedback_garcom"];
    for (const key of requiredColumns) {
        if (colIndex[key] === -1) {
            console.log(`ERRO: Coluna CRÍTICA '${key}' não encontrada na importação GET IN. Verifique o cabeçalho.`);
            return;
        }
    }

    // --- 2. Processamento, Filtragem e Mapeamento ---

    const linhasParaAdicionar: (string | number)[][] = [];

    for (let i: number = 1; i < data.length; i++) {
        const row = data[i];

        const notaValor = row[colIndex.nota];
        const notaStringFormatada = String(notaValor).trim();

        if (notaStringFormatada !== "") {

            // Tenta converter para número decimal (float), tratando a vírgula como ponto decimal
            const nota = parseFloat(notaStringFormatada.replace(',', '.'));

            // FILTRA: Apenas notas válidas (não NaN) e <= 6
            if (!isNaN(nota) && nota <= 6) {

                // 2.1. OBTENÇÃO E TRATAMENTO DOS DADOS

                // --- 2.1.1. TRATAMENTO DE DATA ---
                // Data já está em MM/DD/YYYY ou é um número serial. Mantemos o valor.
                const dataOriginalValor = row[colIndex.data_avaliacao];
                let dataValorParaTabela: string | number;

                if (typeof dataOriginalValor === 'number') {
                    // Se for serial do Excel, mantém o número
                    dataValorParaTabela = dataOriginalValor;
                } else {
                    // Se for string, remove parte da hora e mantém a data (MM/DD/YYYY)
                    let dataString = String(dataOriginalValor).trim();
                    dataValorParaTabela = dataString.split(' ')[0];
                }
                // --- FIM DO TRATAMENTO DE DATA ---

                // --- 2.1.2. TRATAMENTO DE TELEFONE (WHATZAPP) ---
                let whatzappValor: string = "";
                if (colIndex.telefone !== -1) {
                    let telefoneStr = String(row[colIndex.telefone]).replace(/[^0-9]/g, ''); // Remove caracteres não numéricos
                    if (telefoneStr.startsWith("55") && telefoneStr.length >= 11) {
                        // Remove os dois primeiros dígitos ("55") se for um número de celular válido
                        whatzappValor = telefoneStr.substring(2);
                    } else {
                        whatzappValor = telefoneStr;
                    }
                }
                // --- FIM DO TRATAMENTO DE TELEFONE ---

                // --- 2.1.3. CONCATENAÇÃO DO RELATO ---
                const fg = String(row[colIndex.feedback_geral] || "").trim();
                const fgarc = String(row[colIndex.feedback_garcom] || "").trim();

                let relatoFinal = "";
                if (fg) {
                    relatoFinal += `[Geral] ${fg}`;
                }
                if (fgarc) {
                    if (relatoFinal) relatoFinal += " . "; // Adiciona separador se já tiver um relato
                    relatoFinal += `[Garçom] ${fgarc}`;
                }
                // --- FIM DA CONCATENAÇÃO ---

                // Outros campos
                const lojaOriginal = String(row[colIndex.loja]).trim();
                const clienteNome = colIndex.marca !== -1 ? String(row[colIndex.marca]).trim() : "Cliente Get In";

                // 2.2. CRIAÇÃO DA LINHA DE DESTINO
                // Array com 14 posições (A até N) - Omitimos O, P, Q para inserir fórmulas depois.
                const novaLinha: (string | number)[] = new Array(17).fill("");

                // Mapeamento FINAL para as Colunas A a N (0 a 13)

                novaLinha[0] = dataValorParaTabela; // Coluna A: DATA
                novaLinha[1] = "Fale com Dono";     // Coluna B: CANAL (Fixo)
                // novaLinha[2] = nota;                // Coluna C: NOTA (Opcional, mas útil para referência)
                //novaLinha[10] = clienteNome;         // Coluna D: CLIENTE/AVALIADOR
                novaLinha[4] = lojaOriginal;        // Coluna E: UNIDADE (Sem mapeamento)
                novaLinha[5] = "Salão";             // Coluna F: ORIGEM (Fixo)
                // novaLinha[6] = "Salão";             // Coluna G: TIPO (Fixo)
                // Coluna H (7): Vazia
                novaLinha[8] = relatoFinal;         // Coluna I: RELATO / Comentário (Concatenado)
                // Coluna J (9): PEDIDO (Vazia, não existe no mapeamento)
                novaLinha[10] = clienteNome;        // Coluna K: CLIENTE (Nome do Cliente, como o iFood usa "marca")
                novaLinha[11] = whatzappValor;      // Coluna L: WHATZAPP (Telefone tratado)
                // Colunas M (12) e N (13): Vazia

                linhasParaAdicionar.push(novaLinha);
            }
        }
    }

    // --- 3. Colar na Tabela1 e Inserir Fórmulas ---

    if (linhasParaAdicionar.length > 0) {

        // Adiciona as linhas (com apenas 14 colunas: A até N)
        tabelaDestino.addRows(null, linhasParaAdicionar);

        // 3.1. INSERE AS FÓRMULAS NAS COLUNAS O, P e Q (Ano, Mês, Semana)
        const rowCount = tabelaDestino.getRowCount();
        const startRow = rowCount - linhasParaAdicionar.length;

        // Define o intervalo onde as fórmulas serão aplicadas (Colunas O, P, Q das novas linhas)
        // Coluna O é a 15ª coluna, índice 14.
        const rangeFormulas = planilhaDestino.getRange(
            startRow + 1,                               // Linha de início (corpo da tabela)
            14,                                         // Coluna de início (O, índice 14)
            linhasParaAdicionar.length,                 // Número de linhas
            3                                           // Número de colunas (O, P, Q)
        );

        // Fórmulas em PT-BR (Notação R1C1 referente à Coluna A, que é 14 colunas para trás: C[-14])
        const formulasPTBR = [
            [`=ANO(R[0]C[-14])`, `=MÊS(R[0]C[-14])`, `=NUM.SEMANA(R[0]C[-14])`]
        ];

        // Aplica as fórmulas R1C1 para que se adaptem ao idioma Português
        rangeFormulas.setFormulaR1C1(formulasPTBR);

        console.log(`Processamento concluído: ${linhasParaAdicionar.length} avaliações negativas (Nota <= 6) do GET IN foram adicionadas à Tabela '${TABELA_DESTINO_NOME}' na planilha '${PLANILHA_DESTINO_NOME}'.`);
        console.log(`Fórmulas de ANO, MÊS e SEMANA, e tratamento de WhatsApp foram aplicados automaticamente.`);

    } else {
        console.log("Nenhum relato negativo (Nota <= 6) encontrado na importação do GET IN para adicionar à tabela.");
    }

    // 4. Limpa a aba IMPORTAÇÃO (exceto o cabeçalho)
    planilhaImportacao.getRange("A2:Z1000").clear(ExcelScript.ClearAppliedTo.contents);
}
