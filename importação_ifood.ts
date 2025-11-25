// #######################################################################################
// SCRIPT: Processar_Importacao_IFOOD
// Final: Adapta a importação iFood para a Base Geral (Tabela1).
// Filtra: Apenas notas <= 3.
// AJUSTE CRÍTICO: TRATAMENTO DE DATA COMO STRING: DD/MM/YYYY -> MM/DD/YYYY para padronização no Excel.
// #######################################################################################

// OBJETO DE MAPEAMENTO: Define as trocas de valores das unidades do iFood.
const UNIT_MAP: { [key: string]: string } = {
    "CAMARADA BURGERS - PARK JACAREPAGUA": "CAM - Park Jacarepaguá",
    "CAMARADA CAMARAO - PARK JACAREPAGUA": "CAM - Park Jacarepaguá",
    "CAMARADA BURGERS - NOVA AMERICA": "CAM - Nova América",
    "CAMARADA BURGERS - NEW YORK": "CAM - New York City Center",
    "CAMARADA BURGERS - METROPOLITANO": "CAM - Metropolitano",
    "CAMARADA BURGERS - GOIANIA": "CAM - Goiânia",
    "CAMARADA BURGERS - MAG SHOPPING": "CAM - Mag Shopping",
    "CAMARADA BURGERS - RIOSUL": "CAM - Rio Sul",
    "CAMARADA CAMARAO - RIBEIRAO PRETO": "CAM - Ribeirão",
    "CAMARADA BURGERS - MANAUS": "CAM - Manauara",
    "CAMARADA BURGERS - TERESINA": "CAM - Rio Poty",
    "CAMARADA BURGERS - RIO MAR RECIFE": "CAM - Riomar Recife",
    "CAMARADA CAMARAO - RIOSUL": "CAM - Rio Sul",
    "CAMARADA BURGERS - SAO LUIS": "CAM - Shopping da Ilha",
    "CAMARADA BURGERS - RIBEIRAO PRETO": "CAM - Ribeirão",
    "CAMARADA BURGERS - BRASILIA": "CAM - Shopping ID",
    "CAMARADA BURGERS - SALVADOR SHOPPING": "CAM - Salvador Shopping",
    "CAMARADA BURGERS - RIO MAR FORTALEZA": "CAM - Riomar Fortaleza",
    "CAMARADA BURGERS MACEIO": "CAM - Maceió",
    "CAMARADA BURGERS - JARDINS ARACAJU": "CAM - Jardins",
    "CAMARADA CAMARAO - JARDINS ARACAJU": "CAM - Jardins",
    "CAMARADA BURGERS - SALVADOR BARRA": "CAM - Barra",
    "CAMARADA CAMARAO MACEIO": "CAM - Maceió",
    "CAMARADA CAMARAO - RIO MAR RECIFE": "CAM - Riomar Recife",
    "CAMARADA BURGERS - BELEM": "CAM - Boulevard",
    "CAMARADA BURGERS - SHOPPING RECIFE": "CAM - Shopping Center Recife",
    "CAMARADA BURGERS - DOM PEDRO": "CAM - Parque Dom Pedro",
    "CAMARADA CAMARAO - RIO MAR FORTALEZA": "CAM - Riomar Fortaleza",
    "CAMARADA CAMARAO - SALVADOR BARRA": "CAM - Barra",
    "CAMARADA BURGERS SHOPPING LAR CENTER": "CAM - Lar Center",
    "CAMARADA BURGERS - RIO MAR ARACAJU": "CAM - Riomar Aracajú",
    "CAMARADA BURGERS - GRAND PLAZA": "CAM - Grand Plaza",
    "CAMARADA CAMARAO - MOGI DAS CRUZES": "CAM - Mogi",
    "CAMARADA CAMARAO - BRASILIA": "CAM - Shopping ID",
    "CAMARADA CAMARAO - RIO MAR ARACAJU": "CAM - Riomar Aracajú",
    "CAMARADA CAMARAO - SHOPPING RECIFE": "CAM - Shopping Center Recife",
    "CAMARADA CAMARAO - GRAND PLAZA": "CAM - Grand Plaza",
    "CAMARADA BURGERS - MOGI DAS CRUZES": "CAM - Mogi",
    "CAMARADA BURGERS - MOOCA": "CAM - Mooca",
    "CAMARADA CAMARAO - PARK JACAREPAGUACAM": "CAM - Park Jacarepaguá",
    "CAMARADA CAMARAO - MAG SHOPPING": "CAM - Mag Shopping",
    "CAMARADA BURGERS - TAMBORE": "CAM - Tamboré",
    "CAMARADA CAMARAO - CIDADE SAO PAULO": "CAM - Cidade São Paulo",
    "CAMARADA CAMARAO - NEW YORK": "CAM - New York City Center",
    "CAMARADA CAMARAO - METROPOLITANO": "CAM - Metropolitano",
    "CAMARADA CAMARAO - NOVA AMERICA": "CAM - Nova América",
    "CAMARADA CAMARAO - VITORIA": "CAM - Vitória",
    "CAMARADA CAMARAO - MANAUS": "CAM - Manauara",
    "CAMARADA CAMARAO - MOOCA": "CAM - Mooca",
    "CAMARADA CAMARAO - TERESINA": "CAM - Rio Poty",
    "CAMARADA CAMARAO - RIO DESIGN BARRA": "CAM - Rio Design Barra",
    "CAMARADA CAMARAO - DOM PEDRO": "CAM - Parque Dom Pedro",
    "CAMARADA CAMARAO - BELEM": "CAM - Boulevard",
    "CAMARADA BURGERS - SAO CAETANO DO SUL": "CAM - Park São Caetano",
    "CAMARADA CAMARAO - SALVADOR SHOPPING": "CAM - Salvador Shopping",
    "CAMARADA CAMARAO - SAO LUIS": "CAM - Shopping da Ilha",
    "CAMARADA BURGERS - VITORIA": "CAM - Vitória",
    "CAMARADA CAMARAO SHOPPING LAR CENTER": "CAM - Lar Center",
    "CAMARADA BURGERS - RIO DESIGN BARRA": "CAM - Rio Design Barra",
    "CAMARADA CAMARAO - TAMBORE": "CAM - Tamboré",
    "CAMARADA BURGERS - CIDADE SAO PAULO": "CAM - Cidade São Paulo",
    "CAMARADA CAMARAO - SAO CAETANO DO SUL": "CAM - Park São Caetano",
    "CAMARADA CAMARAO - GOIANIA": "CAM - Goiânia"
};

function main(workbook: ExcelScript.Workbook) {

    // #######################################################################
    // !!! CONFIGURAÇÕES CRÍTICAS E VALORES CONFIRMADOS !!!
    const NUMERO_TOTAL_COLUNAS_TABELA1 = 17;
    const PLANILHA_IMPORTACAO_NOME = "IMPORTAÇÃO_IFOOD";
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

    // Mapeamento dos índices das colunas de origem (iFood)
    const colIndex: { [key: string]: number } = {
        data_avaliacao: headers.indexOf("data_avaliacao"),
        marca: headers.indexOf("marca"),
        nota: headers.indexOf("nota"),
        comentario: headers.indexOf("comentario"),
        num_pdd: headers.indexOf("num_pdd"),
        loja: headers.indexOf("loja"),

        // Colunas opcionais
        agregacao: headers.indexOf("agregacao"),
        ID_loja: headers.indexOf("ID_loja"),
        modelo_logistico: headers.indexOf("modelo_logistico"),
        id_pdd: headers.indexOf("id_pdd")
    };

    // Verificação de Colunas CRÍTICAS (as que serão usadas)
    const requiredColumns = ["data_avaliacao", "loja", "nota", "comentario", "num_pdd"];
    for (const key of requiredColumns) {
        if (colIndex[key] === -1) {
            console.log(`ERRO: Coluna CRÍTICA '${key}' não encontrada na importação iFood. Verifique se o cabeçalho está correto.`);
            return;
        }
    }

    // --- 2. Processamento, Filtragem e Mapeamento ---

    const linhasParaAdicionar: (string | number)[][] = [];

    for (let i: number = 1; i < data.length; i++) {
        const row = data[i];

        const notaValor = row[colIndex.nota];

        // FILTRO: Tratamento de String Mais Robusto
        const notaStringFormatada = String(notaValor).trim();

        if (notaStringFormatada !== "") {

            // Tenta converter para número decimal (float), tratando a vírgula como ponto decimal
            const nota = parseFloat(notaStringFormatada.replace(',', '.'));

            // FILTRA: Apenas notas válidas (não NaN) e <= 3
            if (!isNaN(nota) && nota <= 3) {

                // 2.1. OBTENÇÃO E TRATAMENTO DOS DADOS

                const dataOriginalValor = row[colIndex.data_avaliacao];
                // Definimos o tipo como string ou number para evitar o objeto Date.
                let dataValorParaTabela: string | number;

                // --- INÍCIO DO TRATAMENTO DE DATA (TRATAR COMO STRING/GERAL) ---
                if (typeof dataOriginalValor === 'number') {
                    // Caso 1: Data Serial do Excel. Mantenha o número.
                    dataValorParaTabela = dataOriginalValor;
                } else if (typeof dataOriginalValor === 'string') {
                    let dataString = dataOriginalValor.trim();

                    // Passo 1: Se for formato ISO (YYYY-MM-DDTHH:MM...), remove a parte do tempo/timezone (T...)
                    if (dataString.includes('T')) {
                        dataString = dataString.split('T')[0]; // Deixa apenas YYYY-MM-DD
                    }

                    // Tenta detectar o formato PT-BR (DD/MM/YYYY)
                    const parts_br = dataString.split('/');

                    if (parts_br.length === 3) {
                        // Caso 2: Se for DD/MM/YYYY, reestrutura para o formato MM/DD/YYYY (Americano)
                        // EX: 17/10/2025 -> 10/17/2025
                        const day = parts_br[0];
                        const month = parts_br[1];
                        const year = parts_br[2];

                        // Formato MM/DD/YYYY (string) que o Excel deve reconhecer
                        dataValorParaTabela = `${month}/${day}/${year}`;
                    } else {
                        // Caso 3: É uma string em outro formato (provavelmente YYYY-MM-DD), mantenha.
                        dataValorParaTabela = dataString;
                    }
                } else {
                    // Valor inesperado, loga e pula
                    console.log(`AVISO: Data em formato inesperado (não é string nem número) na linha ${i + 1}: '${dataOriginalValor}'. Ignorando esta linha.`);
                    continue;
                }

                // Verifica se a data é minimamente válida (evita passar strings vazias)
                if (typeof dataValorParaTabela === 'string' && dataValorParaTabela.length < 5) {
                    console.log(`AVISO: Data inválida (curta demais) na linha ${i + 1}: '${dataOriginalValor}'. Ignorando esta linha.`);
                    continue;
                }
                // --- FIM DO TRATAMENTO DE DATA ---


                // LÊ O VALOR DA COLUNA 'loja' e usa .trim() para limpar
                const lojaOriginal = String(row[colIndex.loja]).trim();

                // Valor para Coluna D (Cliente/Avaliador)
                const marcaParaColunaD = colIndex.marca !== -1 ? String(row[colIndex.marca]).trim() : lojaOriginal;

                // Aplica a troca de unidade
                const unidadePadronizada = (UNIT_MAP[lojaOriginal] || lojaOriginal);

                // Pedido
                const numPedido = row[colIndex.num_pdd];

                // Comentário/Relato
                const comentario = row[colIndex.comentario];

                // 2.2. CRIAÇÃO DA LINHA DE DESTINO
                // Cria a linha de destino com 17 colunas (tamanho CORRETO)
                const novaLinha: (string | number)[] = new Array(NUMERO_TOTAL_COLUNAS_TABELA1).fill("");

                // Mapeamento FINAL para a Base Geral (Tabela1)

                // Coluna A (0): DATA - Insere o valor (string ou número)
                novaLinha[0] = dataValorParaTabela;

                // Coluna B (1): CANAL
                novaLinha[1] = "IFood";

                // Coluna C (2): NOTA - Vazia


                // Coluna E (4): UNIDADE (Valor Mapeado de 'loja')
                novaLinha[4] = unidadePadronizada;

                // Coluna F (5): ORIGEM
                novaLinha[5] = "Delivery";

                // Coluna G (6): TIPO
                novaLinha[6] = "Delivery";

                // Coluna I (8): RELATO / Comentário
                novaLinha[8] = comentario;

                // Coluna J (9): PEDIDO
                novaLinha[9] = numPedido;


                // Adiciona a linha processada
                linhasParaAdicionar.push(novaLinha);
            }
        }
    }

    // 3. Colar na Tabela1 usando addRows
    if (linhasParaAdicionar.length > 0) {

        // Adiciona as linhas (com 17 colunas)
        tabelaDestino.addRows(null, linhasParaAdicionar);

        console.log(`Processamento concluído: ${linhasParaAdicionar.length} relatos negativos (Nota <= 3) foram adicionados à Tabela '${TABELA_DESTINO_NOME}' na planilha '${PLANILHA_DESTINO_NOME}'.`);

    } else {
        console.log("Nenhum relato negativo (Nota <= 3) encontrado na importação para adicionar à tabela.");
    }

    // 4. Limpa a aba IMPORTAÇÃO (exceto o cabeçalho)
    planilhaImportacao.getRange("A2:Z1000").clear(ExcelScript.ClearAppliedTo.contents);
}
