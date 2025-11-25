// FUNÇÃO AUXILIAR: Converte o número de série da data do Excel para uma string DD/MM/YYYY.
function excelDateToJSDate(excelSerial: number): string {
    // Retorna vazio se o valor não for um número ou for menor que 1
    if (typeof excelSerial !== 'number' || excelSerial <= 0) {
        return "";
    }

    // A data base do Excel é 30 de Dezembro de 1899, no fuso horário UTC
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));

    // Converte o número de série para milissegundos
    let ms = excelSerial * 24 * 60 * 60 * 1000;

    // Ajuste para o bug histórico do ano bissexto de 1900 no Excel:
    if (excelSerial > 60) {
        ms -= (24 * 60 * 60 * 1000);
    }

    const date = new Date(excelEpoch.getTime() + ms);

    // Formata a data para a string DD/MM/YYYY
    const day = String(date.getUTCDate()).padStart(2, '0');
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const year = date.getUTCFullYear();

    return `${day}/${month}/${year}`;
}


function main(workbook: ExcelScript.Workbook) {
    const activeSheet = workbook.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();

    if (!selectedRange) {
        console.log("ERRO: Nenhuma célula selecionada.");
        return;
    }

    const rowIndex = selectedRange.getRowIndex();

    // Índices das colunas: A=0, B=1, E=4, I=8, K=10, L=11.

    const queixa = activeSheet.getRangeByIndexes(rowIndex, 1, 1, 1).getText().toUpperCase();
    const colunaE = activeSheet.getRangeByIndexes(rowIndex, 4, 1, 1).getText();

    const rawDateValue = activeSheet.getRangeByIndexes(rowIndex, 0, 1, 1).getValue() as number;
    const colunaA = excelDateToJSDate(rawDateValue);

    const colunaK = activeSheet.getRangeByIndexes(rowIndex, 10, 1, 1).getText();
    const relato = activeSheet.getRangeByIndexes(rowIndex, 8, 1, 1).getText();

    // Coluna L (WhatsApp)
    let colunaL = activeSheet.getRangeByIndexes(rowIndex, 11, 1, 1).getText().trim();
    if (colunaL === "" || colunaL === null || colunaL === undefined) {
        colunaL = "(sem número)";
    }

    // MONTAGEM DA MENSAGEM
    const mensagem =
        "QUEIXA (" + queixa + ")\n" +
        colunaE + "\n" +
        colunaA + "\n" +
        colunaK + "\n" +
        colunaL + "\n\n" +
        "RELATO:  " + relato;

    // ABA "RELATO"
    let relatoSheet = workbook.getWorksheet("RELATO");
    if (!relatoSheet) {
        relatoSheet = workbook.addWorksheet("RELATO");
    }

    const celulaCopia = relatoSheet.getRange("A1");

    // COLAR E FORMATAR
    celulaCopia.setValue(mensagem);
    celulaCopia.getFormat().setWrapText(true);
    celulaCopia.getFormat().setRowHeight(300);

    // FINALIZAÇÃO
    relatoSheet.activate();
    celulaCopia.select();

    console.log("Mensagem gerada com tratamento de número vazio em 'WhatsApp'.");
}
