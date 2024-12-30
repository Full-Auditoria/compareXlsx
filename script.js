document.getElementById('compare-form').addEventListener('submit', async function (event) {
    event.preventDefault();
  
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];
    const compareBy = document.getElementById('compare-by').value;  // Captura a escolha do usuário (Nome ou Cuc)
  
    if (!file1 || !file2) {
      alert('Por favor, selecione os dois arquivos!');
      return;
    }
  
    const readExcelOrCsv = async (file) => {
      const extension = file.name.split('.').pop().toLowerCase();
  
      const data = await file.arrayBuffer();
      let parsedData;
  
      if (extension === 'csv') {
        parsedData = await parseCsv(data);
      } else {
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        parsedData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
      }
  
      return parsedData;
    };
  
    // Função para tratar CSV
    const parseCsv = (data) => {
      const text = new TextDecoder().decode(data);
      const rows = text.split('\n').map(row => row.split(','));
      return rows;
    };
  
    const data1 = await readExcelOrCsv(file1); // Planilha 1
    const data2 = await readExcelOrCsv(file2); // Planilha 2
  
    // Normaliza os cabeçalhos para minúsculas
    const normalizeHeader = (header) => header.toLowerCase().trim();
  
    const headers1 = data1[0].map(normalizeHeader);
    const headers2 = data2[0].map(normalizeHeader);
  
    const rows1 = data1.slice(1); // Dados sem cabeçalhos
    const rows2 = data2.slice(1);
  
    // Índices das colunas (insensível a maiúsculas/minúsculas)
    const cucIndex1 = headers1.indexOf("cuc");
    const nomeIndex1 = headers1.indexOf("nome");
    const cucIndex2 = headers2.indexOf("cuc");
    const nomeIndex2 = headers2.indexOf("nome");
  
    if (cucIndex1 === -1 || nomeIndex1 === -1 || cucIndex2 === -1 || nomeIndex2 === -1) {
      alert("Certifique-se de que ambas as planilhas possuem as colunas 'Cuc' e 'Nome'.");
      return;
    }
  
    // Normalizar valor
    const normalizeValue = (value) => {
      return value ? value.toString().toLowerCase().trim().replace(/\s+/g, ' ') : '';
    };
  
    // Remover duplicatas - utilizando Set
    const uniqueRows1 = Array.from(new Set(rows1.map(row => normalizeValue(compareBy === 'nome' ? row[nomeIndex1] : row[cucIndex1]))));
    const uniqueRows2 = Array.from(new Set(rows2.map(row => normalizeValue(compareBy === 'nome' ? row[nomeIndex2] : row[cucIndex2]))));
  
    // Comparação e coleta de resultados
    const results = [["Cuc", "Nome", "Resultado", "Linha PLAN A", "Linha PLAN B"]];
    rows1.forEach((row1, rowIndex1) => {
      const value1 = normalizeValue(compareBy === 'nome' ? row1[nomeIndex1] : row1[cucIndex1]);
      
      // Encontrar se o valor existe na outra planilha
      const matchIndex = rows2.findIndex((row2, rowIndex2) => normalizeValue(compareBy === 'nome' ? row2[nomeIndex2] : row2[cucIndex2]) === value1);
      
      if (matchIndex !== -1) {
        // Adicionar à lista de resultados
        results.push([
          normalizeValue(row1[cucIndex1]),
          normalizeValue(row1[nomeIndex1]),
          "Encontrado",
          rowIndex1 + 2, // Adicionando 2 para compensar o cabeçalho e o índice começar de 1
          matchIndex + 2
        ]);
      } else {
        results.push([
          normalizeValue(row1[cucIndex1]),
          normalizeValue(row1[nomeIndex1]),
          "Não encontrado",
          rowIndex1 + 2,
          ""
        ]);
      }
    });
  
    // Criar planilha com resultados
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(results);
    XLSX.utils.book_append_sheet(wb, ws, "Resultados");
  
    const resultExcel = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([resultExcel], { type: 'application/octet-stream' });
  
    // Download do arquivo
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Resultados.xlsx';
    link.click();
    URL.revokeObjectURL(url);
  
    // Feedback ao usuário
    document.getElementById('result').innerHTML = `
      <div class="alert alert-success" role="alert">
        Comparação concluída! O arquivo <strong>Resultados.xlsx</strong> foi gerado e baixado.
      </div>
    `;
  });
  