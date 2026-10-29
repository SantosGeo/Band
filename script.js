const inputArquivo = document.getElementById('arquivo');
const btnGerar = document.getElementById('generate');
const preview = document.getElementById('preview');
const btnSalve = document.getElementById('salve');



let dadosCompletos = null;

//Realiza o download apenas se o arquivo for importado

btnSalve.addEventListener('click', () =>{
    if(!dadosCompletos || dadosCompletos.length ===0){
        alert('Faça o download da lista');
        return;
    }

    const ws = XLSX.utils.aoa_to_sheet(dadosCompletos);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lista Completa");
    XLSX.writeFile(wb, "lista_votacao_completa.xlsx");
})

btnGerar.addEventListener('click', () => {
    const arquivo = inputArquivo.files[0];

    if (!arquivo) {
        alert('Selecione um arquivo xlsx ou csv');
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {

        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const primeiraAba = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[primeiraAba];

            const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            dadosCompletos = aoa;

            // Pega só o cabeçalho + as 10 primeiras linhas de dados
            const previewAoa = aoa.slice(0, 11); 

            const previewSheet = XLSX.utils.aoa_to_sheet(previewAoa);
            const htmlTabela = XLSX.utils.sheet_to_html(previewSheet);

            preview.innerHTML = `
                <div style="max-width: 95%; margin: 40px auto 0; overflow-x: auto;">
                    <h2 style="text-align:center; margin-bottom:15px; font-size:20px; color:#1a1a1a;">
                        Prévia da Lista (mostrando os 10 primeiros)
                    </h2>
                    ${htmlTabela}
                    <p style="text-align:center; margin-top:12px; font-size:13px; color:#666;">
                        Total de registros: <strong>${aoa.length - 1}</strong> 
                        ${aoa.length > 11 ? '(exibindo apenas os 10 primeiros na prévia)' : ''}
                    </p>
                </div>
            `;

            
            const tabela = preview.querySelector('table');
            if (tabela) {
                tabela.style.width = '100%';
                tabela.style.borderCollapse = 'collapse';
                tabela.style.fontSize = '14px';
                tabela.style.background = 'white';
                tabela.style.boxShadow = '0 2px 10px rgba(0,0,0,0.08)';
                tabela.style.borderRadius = '8px';

                tabela.querySelectorAll('th').forEach(th => {
                    th.style.background = '#f1f3f5';
                    th.style.fontWeight = '600';
                });
            }

        } catch (erro) {
            preview.innerHTML = `
                <p style="text-align:center; color:#c8102e;">
                    Verifique se é um Excel válido (.xlsx ou .xls).
                </p>
            `;
            console.error(erro);
        }

        btnSalve.style.display='block';
        btnSalve.innerHTML
    };

    reader.onerror = () => {
        preview.innerHTML = `
            <p style="text-align:center; margin-top:30%; color:#c8102e;">
                Erro ao carregar o arquivo.
            </p>
        `;
    };

    reader.readAsArrayBuffer(arquivo);
});