/**
 * Obtém a lista de setores cadastrados no sistema
 * @returns {string[]} Array com os nomes dos setores
 */
function obterSetores() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planilhaSetores = ss.getSheetByName("SETORES");
    
    if (!planilhaSetores) {
        console.error("Planilha SETORES não encontrada");
        return [];
    }

    const dados = planilhaSetores.getRange("A2:A" + planilhaSetores.getLastRow()).getValues();
    return dados.map(row => row[0]).filter(setor => setor !== "");
}

/**
 * Verifica se um setor está cadastrado no sistema
 * @param {string} setor - Nome do setor a ser verificado
 * @returns {boolean} True se o setor existir, False caso contrário
 */
function verificarSetorCadastrado(setor) {
    const setores = obterSetores();
    return setores.includes(setor.toUpperCase());
}

/**
 * Adiciona um novo setor ao sistema
 * @param {string} setor - Nome do setor a ser adicionado
 * @returns {boolean} True se adicionado com sucesso, False caso contrário
 */
function adicionarSetor(setor) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const planilhaSetores = ss.getSheetByName("SETORES");
    
    if (!planilhaSetores) {
        console.error("Planilha SETORES não encontrada");
        return false;
    }

    // Verifica se o setor já existe
    if (verificarSetorCadastrado(setor)) {
        return false;
    }

    // Adiciona o novo setor
    planilhaSetores.appendRow([setor.toUpperCase()]);
    return true;
} 