// #######################################################################################
// SCRIPT: Limpar_Planilhas_Importacao (CORREÇÃO FINAL)
// FINALIDADE: Limpa o conteúdo das abas de importação (exceto cabeçalho).
// Alvos: IMPORTAÇÃO_IFOOD, IMPORTAÇÃO_GOOGLE, IMPORTAÇÃO_GETIN.
// CORREÇÃO: Uso de ExcelScript.ClearApplyTo.contents (removido o 'ed' extra).
// #######################################################################################

function main(workbook: ExcelScript.Workbook) {

    // Lista de nomes das planilhas de importação
    const PLANILHAS_DE_IMPORTACAO = [
        "IMPORTAÇÃO_IFOOD",
        "IMPORTAÇÃO_GOOGLE",
        "IMPORTAÇÃO_GETIN"
    ];

    let countLimpezas = 0;

    console.log("Iniciando limpeza das planilhas de importação...");

    for (const nomePlanilha of PLANILHAS_DE_IMPORTACAO) {

        const planilha = workbook.getWorksheet(nomePlanilha);

        if (!planilha) {
            console.log(`AVISO: Planilha '${nomePlanilha}' não encontrada. Pulando.`);
            continue;
        }

        try {

            const usedRange = planilha.getUsedRange();

            // Só tenta limpar se houver mais de uma linha (ou seja, dados abaixo do cabeçalho)
            if (usedRange && usedRange.getRowCount() > 1) {

                // Limpa o conteúdo de A2 em diante (mantém o cabeçalho na linha 1).
                const rangeParaLimpar = planilha.getRange("A2:Z1000");

                // *** CORREÇÃO APLICADA AQUI: ClearApplyTo ***
                rangeParaLimpar.clear(ExcelScript.ClearApplyTo.contents);

                console.log(`  - Planilha '${nomePlanilha}' limpa (conteúdo abaixo do cabeçalho removido).`);
                countLimpezas++;

            } else {
                console.log(`  - Planilha '${nomePlanilha}' já estava limpa.`);
            }

        } catch (e) {
            // Em caso de erro, apenas reporta o nome da planilha e o erro.
            console.log(`ERRO ao limpar a planilha '${nomePlanilha}': ${e.toString()}`);
        }
    }

    if (countLimpezas > 0) {
        console.log(`\nLimpeza concluída com sucesso. ${countLimpezas} planilhas foram limpas.`);
    } else {
        console.log("\nLimpeza concluída. Nenhuma alteração foi necessária.");
    }
}
