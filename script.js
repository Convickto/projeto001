let responsavelTesouraria = "";
let responsavelOperacoes = "";

// Objeto para armazenar contagem acumulada dos problemas
const contagemAcumuladaProblemas = {
    "Camera nao funciona": 0,
    "Pino para fora": 0,
    "CF fora do patio": 0,
    "Bateria descarregada": 0,
    "Bag atrasada": 0
};

// Evento de envio do formulário
document.getElementById("dataForm").addEventListener("submit", function(event) {
    event.preventDefault();

    // Preencher os responsáveis apenas na primeira submissão
    if (responsavelTesouraria === "" || responsavelOperacoes === "") {
        responsavelTesouraria = prompt("Digite o nome do Responsável Tesouraria:");
        responsavelOperacoes = prompt("Digite o nome do Responsável Operações:");
        document.getElementById("responsaveisDisplay").innerHTML = `<strong>Responsável Tesouraria:</strong> ${responsavelTesouraria}<br><strong>Responsável Operações:</strong> ${responsavelOperacoes}`;
    }

    const rota = document.getElementById("rota").value;
    const re = document.getElementById("re").value;
    const carro = document.getElementById("carro").value;
    const rua = document.getElementById("rua").value;
    const horaInicial = document.getElementById("horaInicial").value;
    const horaFinal = document.getElementById("horaFinal").value;
    const observacoes = document.getElementById("observacoes").value;

    const checkboxes = document.querySelectorAll('.checklist input[type="checkbox"]');
    let marcados = [];
    checkboxes.forEach(function(checkbox) {
        if (checkbox.checked) {
            marcados.push(checkbox.value);
        }
    });
    const itensMarcados = marcados.join(';');

    const tableBody = document.getElementById("tableBody");
    const newRow = document.createElement("tr");
    newRow.innerHTML = `<td>${responsavelTesouraria}</td><td>${responsavelOperacoes}</td><td>${rota}</td><td>${re}</td><td>${carro}</td><td>${rua}</td><td>${horaInicial}</td><td>${horaFinal}</td><td>${itensMarcados}</td><td>${observacoes}</td>`;
    tableBody.appendChild(newRow);

    atualizarContagemProblemas(marcados);
    document.getElementById("dataForm").reset();
    checkboxes.forEach(checkbox => checkbox.checked = false);
});

// Função para atualizar e exibir contagem acumulada dos problemas
function atualizarContagemProblemas(itensMarcados) {
    itensMarcados.forEach(item => {
        if (contagemAcumuladaProblemas.hasOwnProperty(item)) {
            contagemAcumuladaProblemas[item]++;
        }
    });

    // Ordenar e exibir a contagem acumulada dos problemas
    const chartProblemas = document.getElementById("chartProblemas");
    chartProblemas.innerHTML = "<h3>Contagem de Problemas:</h3>";
    const contagemOrdenada = Object.entries(contagemAcumuladaProblemas)
        .sort((a, b) => b[1] - a[1]);
    contagemOrdenada.forEach(([problema, count]) => {
        chartProblemas.innerHTML += `<p>${problema}: ${count}</p>`;
    });
}

// Evento para exportar a tabela para Excel com a contagem dos problemas
document.getElementById("exportExcel").addEventListener("click", function() {
    const wb = XLSX.utils.book_new();
    
    // Criação de uma nova aba com a contagem dos problemas
    const dadosProblemas = [
        ["Resumo de Problemas:", ""],
        ...Object.entries(contagemAcumuladaProblemas).map(([problema, count]) => [problema, count]),
        ["", ""]
    ];

    const wsProblemas = XLSX.utils.aoa_to_sheet(dadosProblemas);
    const wsTabela = XLSX.utils.table_to_sheet(document.getElementById("dataTable"));

    // Combinar ambas as abas em uma só
    XLSX.utils.sheet_add_aoa(wsTabela, dadosProblemas, { origin: 0 });

    // Adicionar as abas ao workbook
    XLSX.utils.book_append_sheet(wb, wsTabela, "Relatório");
    XLSX.writeFile(wb, "relatorio_carregamento.xlsx");
});
