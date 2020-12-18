const axios = require('axios');
const excel = require('excel4node');

function parseData(dataRow) {
    let parsedData = {
        fase: parseInt(dataRow.nivel.numero),
        tempo: parseInt(dataRow.tempoFinal),
        erros: parseInt(dataRow.erros),
        idade: parseInt(dataRow.usuario.idade),
        escolaridade: parseInt(dataRow.usuario.escolaridade),
        tipoEscola: dataRow.usuario.tipoEscola,
        sexo: dataRow.usuario.sexo,
        transtorno: dataRow.usuario.transtorno
    }
    return parsedData;
}

async function main() {
    var workbook = new excel.Workbook();
    var worksheet = workbook.addWorksheet('Relatório');
    const titulos = ['Fase', 'Tempo', 'Erros', 'Idade',
        'Escolaridade', 'Tipo de Escola', 'Sexo', 'Transtorno'];

    var { data } = await axios.get('http://200.141.166.245:8080/resultado');

    data.sort((a, b) => a.nivel.numero - b.nivel.numero);

    for (let i = 2; i < titulos.length + 2; i++) {
        worksheet.cell(1, i).string(titulos[i - 2]);
    }

    for (let i = 2; i < data.length + 2; i++) {
        let parsedData = parseData(data[i - 2]);

        worksheet.cell(i, 2).number(parsedData.fase);
        worksheet.cell(i, 3).number(parsedData.tempo);
        worksheet.cell(i, 4).number(parsedData.erros);
        worksheet.cell(i, 5).number(parsedData.idade);
        worksheet.cell(i, 6).number(parsedData.escolaridade);
        worksheet.cell(i, 7).string(parsedData.tipoEscola);
        worksheet.cell(i, 8).string(parsedData.sexo);
        worksheet.cell(i, 9).string(parsedData.transtorno);

    }

    try {
        await workbook.write('Relatório.xlsx');
        console.log('Relatório gerado com sucesso!');
    } catch (error) {
        console.error(error);
    }
}

main();