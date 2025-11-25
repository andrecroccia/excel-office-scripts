// #######################################################################################
// SCRIPT: Processar_Importacao (Versão FINAL com DATA NEUTRA)
// Configurações: Tabela1 = 17 colunas.
// Comportamento: A coluna de Data é transferida SEM formatação ou conversão.
// #######################################################################################

// FUNÇÃO AUXILIAR: Removida a função de parsing de data, pois agora a data é transferida diretamente.

function main(workbook: ExcelScript.Workbook) {

    // #######################################################################
    // !!! CONFIGURAÇÕES CRÍTICAS E VALORES CONFIRMADOS !!!
    const NUMERO_TOTAL_COLUNAS_TABELA1 = 17;
    const PLANILHA_IMPORTACAO_NOME = "IMPORTAÇÃO_GOOGLE";
    const PLANILHA_DESTINO_NOME = "Base Geral";
    const TABELA_DESTINO_NOME = "Tabela1";
    // As variáveis de formato de data foram removidas, pois não serão usadas.
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

    // Lê os dados da planilha de importação
    const usedRange = planilhaImportacao.getUsedRange();
    if (!usedRange || usedRange.getRowCount() <= 1) {
        console.log(`Planilha '${PLANILHA_IMPORTACAO_NOME}' está vazia ou só tem cabeçalho.`);
        if (usedRange) {
            planilhaImportacao.getRange("A2:Z1000").clear(ExcelScript.ClearAppliedTo.contents);
        }
        return;
    }

    // O bloco de formatação de Coluna C (Data) foi removido aqui.

    // Leitura dos valores
    const data = planilhaImportacao.getUsedRange().getValues();
    const headers = data[0] as string[];

    // Localiza os índices das colunas de origem (Google/HubLocal)
    const colIndex = {
        cod: headers.indexOf("Cod."),
        data: headers.indexOf("Data"),
        avaliador: headers.indexOf("Avaliador"),
        nota: headers.indexOf("Nota"),
        comentario: headers.indexOf("Comentário")
    };

    for (const key in colIndex) {
        if (colIndex[key] === -1) {
            console.log(`ERRO: Coluna '${key}' não encontrada na importação. Verifique se o cabeçalho está correto.`);
            return;
        }
    }

    // --- 2. Processamento, Filtragem e Mapeamento ---

    const linhasParaAdicionar: (string | number)[][] = [];

    for (let i: number = 1; i < data.length; i++) {
        const row = data[i];

        const notaValor = row[colIndex.nota];

        // FILTRAGEM: Garante que o valor da nota não é nulo/vazio.
        const notaStringFormatada = String(notaValor).trim();

        if (notaStringFormatada !== "") {
            const nota = Number(notaStringFormatada);

            // FILTRA: Apenas notas válidas (não NaN) e <= 3
            if (!isNaN(nota) && nota <= 3) {

                // CRUCIAL: A data é lida e usada DIRETAMENTE
                const dataOriginal = row[colIndex.data];

                // Cria a linha de destino com 17 colunas
                const novaLinha: (string | number)[] = new Array(NUMERO_TOTAL_COLUNAS_TABELA1).fill("");

                // Mapeamento:
                novaLinha[0] = dataOriginal;             // Coluna A (DATA - VALOR ORIGINAL)
                novaLinha[1] = "Google";                 // Coluna B (Canal)
                novaLinha[4] = row[colIndex.cod];        // Coluna E (Unidade)
                novaLinha[5] = "Salão";
                
                novaLinha[8] = row[colIndex.comentario]; // Coluna I (Comentário/RELATO)
                novaLinha[10] = row[colIndex.avaliador]; // Coluna K (CLIENTE/Avaliador)

                linhasParaAdicionar.push(novaLinha);
            }
        }
    }

    // 3. Colar na Tabela1 usando addRows
    if (linhasParaAdicionar.length > 0) {

        // 3.1. Adiciona as linhas (com 17 colunas)
        tabelaDestino.addRows(null, linhasParaAdicionar);

        // O bloco de formatação de Coluna A (Data) foi removido aqui.

        console.log(`Processamento concluído: ${linhasParaAdicionar.length} relatos negativos (Nota <= 3) foram adicionados à Tabela '${TABELA_DESTINO_NOME}'.`);

    } else {
        console.log("Nenhum relato negativo (Nota <= 3) encontrado na importação para adicionar à tabela.");
    }

    // 4. Limpa a aba IMPORTAÇÃO (exceto o cabeçalho)
    planilhaImportacao.getRange("A2:Z1000").clear(ExcelScript.ClearAppliedTo.contents);
}
