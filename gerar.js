const XLSX = require('xlsx');

// Função para calcular a média aritmética de um conjunto de notas
function noteNeedToPass(nota1, nota2) {
    const noteNeed = (42.0 - nota1 - (nota2 * 2)) / 3
    return noteNeed;
}

function noteSumWithWeights(nota1, nota2) {
    const noteSum = nota1 + (nota2 * 2)
    return noteSum
}

// Carregar as planilhas de entrada
const planilha1 = XLSX.readFile('mapao.xlsx');
const planilha2 = XLSX.readFile('mapao (2).xlsx');

// Inicializar a planilha de saída
const planilhaCombinada = XLSX.utils.book_new();

// Encontrar as disciplinas (cabeçalhos das colunas)
const disciplinas = [];
for (const planilha of [planilha1, planilha2]) {
    const sheet = planilha.Sheets[planilha.SheetNames[0]];
    const header = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0];
    for (let i = 2; i < header.length; i++) {
        if (header[i] === 'Média') {
            break; // Não importa mais a partir da coluna 'Média'
        }
        if (!disciplinas.includes(header[i])) {
            disciplinas.push(header[i]);
        }
    }
}

// Para cada disciplina, criar uma aba e inserir os nomes dos alunos e as notas correspondentes
for (const disciplina of disciplinas) {
    const sheet1 = planilha1.Sheets[planilha1.SheetNames[0]];
    const sheet2 = planilha2.Sheets[planilha2.SheetNames[0]];

    const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
    const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

    const indexOfDiscipline1 = data1[0].indexOf(disciplina);

    console.log(data1[0][indexOfDiscipline1]);

    // Filtrar apenas as notas da disciplina atual




    // Criar uma nova planilha para a disciplina
    const novaPlanilha = XLSX.utils.book_new();

    const novaSheet = XLSX.utils.aoa_to_sheet([[]]);;
    
    let tableOrigin = 5


    novaSheet['A' + tableOrigin] = { v: 'Nome do Aluno', t: 's' };
    novaSheet['B' + tableOrigin] = { v: '1º TRI',   t: 's' };
    novaSheet['C' + tableOrigin] = { v: '2º TRI',   t: 's' };
    novaSheet['D' + tableOrigin] = { v: 'A',   t: 's' };
    novaSheet['E' + tableOrigin] = { v: 'B',   t: 's' };
    novaSheet['F' + tableOrigin] = { v: '3º TRI',   t: 's' };
    novaSheet['G' + tableOrigin] = { v: 'C',   t: 's' };
    novaSheet['H' + tableOrigin] = { v: 'D',   t: 's' };
    novaSheet['I' + tableOrigin] = { v: 'Prova Final',   t: 's' };
    novaSheet['J' + tableOrigin] = { v: 'Média Final',   t: 's' };

    // Inserir os nomes dos alunos na primeira coluna
    for (let i = 1; i < data1.length; i++) {

        if (data1[i][1] === 'Média:')
            break;

        if (data1[i][1] === data2[i][1]) {
            novaSheet['A' + (i + tableOrigin)] = { v: data1[i][1], t: 's' };
        }
        else {
            novaSheet['A' + (i + tableOrigin)] = { v: 'Valores diferentes nas duas planilhas', t: 's' };
        }

        if (i == 0)
            continue;

        const nota1 = parseFloat(data1[i][indexOfDiscipline1])
        const nota2 = parseFloat(data2[i][indexOfDiscipline1])

        const needToPass = noteNeedToPass(nota1, nota2)
        const noteSum = noteSumWithWeights(nota1, nota2);

        novaSheet['B' + (i + tableOrigin)] = { v: data1[i][indexOfDiscipline1], t: 's' };
        novaSheet['C' + (i + tableOrigin)] = { v: data2[i][indexOfDiscipline1], t: 's' };
        novaSheet['D' + (i + tableOrigin)] = { v: noteSum.toFixed(1), t: 's' };
        novaSheet['E' + (i + tableOrigin)] = { v: needToPass.toFixed(1), t: 's' };

    }


    novaSheet['!ref'] = 'A1:Z99'


    // Adicionar a nova planilha à planilha de saída com o nome da disciplina
    XLSX.utils.book_append_sheet(planilhaCombinada, novaSheet, disciplina);
}

// Salvar a planilha de saída
XLSX.writeFile(planilhaCombinada, 'planilha_combinada.xlsx');

console.log('Planilha combinada gerada com sucesso.');