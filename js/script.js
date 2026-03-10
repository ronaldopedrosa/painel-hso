// =====================================================================
// VARIÁVEL GLOBAL DE DADOS
// =====================================================================
let rawData = []; 

// =====================================================================
// FUNÇÃO PARA LER MÚLTIPLOS ARQUIVOS EXCEL (LÊ CÉLULAS E NOME DO ARQUIVO)
// =====================================================================
async function carregarExcel(input) {
    const files = input.files;
    if (!files || files.length === 0) return;

    let novosDados = [];

    const lerArquivoIndividual = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            // CAPTURA O NOME DO ARQUIVO PARA IDENTIFICAR OS LOOPS
            const fileName = file.name.toUpperCase(); 

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});

                    let worksheet = workbook.Sheets["BASE_CONSOLIDADA"];
                    if (!worksheet) {
                        worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    }

                    const jsonExcel = XLSX.utils.sheet_to_json(worksheet);

                    function getValue(row, possibleNames) {
                        const rowKeys = Object.keys(row);
                        const normalize = (str) => str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
                        for (let name of possibleNames) {
                            const target = normalize(name);
                            const foundKey = rowKeys.find(key => normalize(key) === target);
                            if (foundKey) return row[foundKey];
                        }
                        return ""; 
                    }

                    const dadosProcessados = jsonExcel.map(linha => {
                        const rawSys = getValue(linha, ["Sala/Sistema", "Sistema", "Area", "Grupos de Instrumentos"]);
                        const tag = getValue(linha, ["TAG Hemobrás", "TAG", "Tag", "Instrumento", "Codigo"]);
                        const desc = getValue(linha, ["Descrição dos Equipamentos", "Descrição", "Descricao", "Nome"]);
                        const origem = getValue(linha, ["ORIGEM", "Origem"]);
                        
                        const sysText = String(rawSys).toUpperCase();
                        let finalSys = "Geral / Outros"; 

                        // === REGRAS DE TRADUÇÃO INTELIGENTES ===
                        if (sysText.includes("PW") || sysText.includes("LOOP") || sysText.includes("PURIFICADA") || fileName.includes("PW") || fileName.includes("LOOP")) {
                            if (sysText.includes("1") || fileName.includes("1")) finalSys = "PW_LOOP - 1";
                            else if (sysText.includes("2") || fileName.includes("2")) finalSys = "PW_LOOP - 2";
                            else if (sysText.includes("3") || fileName.includes("3")) finalSys = "PW_LOOP - 3";
                            else if (sysText.includes("4") || fileName.includes("4")) finalSys = "PW_LOOP - 4";
                            else if (sysText.includes("5") || fileName.includes("5")) finalSys = "PW_LOOP - 5";
                            else finalSys = "PW_LOOP"; 
                        }
                        else if (sysText.includes("HNO") || sysText.includes("HNA") || sysText.includes("QUIMICO") || sysText.includes("QUÍMICO")) finalSys = "Químicos";
                        else if (sysText.includes("HSO") || sysText.includes("SULFURICO") || sysText.includes("ACIDO") || sysText.includes("H2SO4")) finalSys = "Ácido Sulfúrico";
                        else if (sysText.includes("WW") || sysText.includes("EFLUENTE") || sysText.includes("ESGOTO") || sysText.includes("TRATAMENTO")) finalSys = "Efluentes";
                        else if (sysText.includes("CA-") || sysText.includes("-CA") || sysText.includes(" CAP ") || sysText.includes("AR COMP") || sysText.includes("COMPRIMIDO") || sysText === "CA" || sysText === "CAP") finalSys = "Ar Comprimido";
                        else if (rawSys.length > 2) finalSys = rawSys; 

                        return {
                            sistema: finalSys, 
                            tag: tag,
                            desc: desc,
                            local: getValue(linha, ["Local", "Area"]), 
                            calib: getValue(linha, ["Calibração (SIM ou NÃO)", "Calibracao", "Criticidade"]),
                            status_calib: getValue(linha, ["Status de qualificação", "Status", "Situação"]),
                            defeito: getValue(linha, ["Defeito", "Com defeito", "Falha"]), 
                            observacao: getValue(linha, ["Observação", "Observacao", "Obs", "Motivo"]), // NOVA COLUNA DE OBSERVAÇÃO
                            origem: origem
                        };
                    });

                    resolve(dadosProcessados.filter(item => item.tag !== "" || item.desc !== ""));

                } catch (error) {
                    console.error("Erro ao ler arquivo:", file.name, error);
                    resolve([]); 
                }
            };
            reader.readAsArrayBuffer(file);
        });
    };

    try {
        const promessas = Array.from(files).map(file => lerArquivoIndividual(file));
        const resultados = await Promise.all(promessas);
        rawData = resultados.flat();

        alert(`SUCESSO! ${rawData.length} equipamentos carregados.`);
        init();

    } catch (err) {
        alert("Ocorreu um erro ao processar os arquivos.");
        console.error(err);
    }
}

// =====================================================================
// INICIALIZAÇÃO E FILTROS
// =====================================================================

function init() {
    if (rawData.length === 0) {
        document.getElementById('tableBody').innerHTML = '<tr><td colspan="9" style="text-align:center; padding:20px; color:#0056b3;">📂 <b>Aguardando arquivos...</b></td></tr>';
        return;
    }
    populateLocals();
    populateSystems();
    applyFilters();
}

function populateLocals() {
    const select = document.getElementById('filterLocal');
    select.innerHTML = '<option value="">Todos os Locais</option>'; 
    const locals = [...new Set(rawData.map(i => i.local ? String(i.local).trim() : ""))].filter(l => l).sort();
    locals.forEach(loc => {
        const opt = document.createElement('option');
        opt.value = loc;
        opt.innerText = loc;
        select.appendChild(opt);
    });
}

function populateSystems() {
    const select = document.getElementById('filterSystem');
    select.innerHTML = '<option value="">Mostrar Tudo</option>';
    const systems = [...new Set(rawData.map(i => i.sistema ? String(i.sistema).trim() : ""))].filter(s => s).sort();
    
    systems.forEach(sys => {
        const opt = document.createElement('option');
        opt.value = sys;
        let emoji = "🔧";
        if (sys === "Efluentes") emoji = "💧";
        if (sys === "Ar Comprimido") emoji = "💨";
        if (sys === "Ácido Sulfúrico" || sys === "Químicos") emoji = "🧪"; 
        if (sys.startsWith("PW_LOOP")) emoji = "💦"; 
        opt.innerText = `${emoji} ${sys}`;
        select.appendChild(opt);
    });
}

function applyFilters() {
    const sys = document.getElementById('filterSystem').value;
    const search = document.getElementById('searchInput').value.toLowerCase();
    const loc = document.getElementById('filterLocal').value;
    const calibType = document.getElementById('filterCalib').value;
    
    const defElement = document.getElementById('filterDefeito');
    const defType = defElement ? defElement.value : "";

    const filtered = rawData.filter(item => {
        const iSys = String(item.sistema || "");
        const iTag = String(item.tag || "").toLowerCase();
        const iDesc = String(item.desc || "").toLowerCase();
        const iLoc = String(item.local || "").trim();
        const iCalib = String(item.calib || "").toUpperCase();
        const iStatus = String(item.status_calib || "").toUpperCase();
        const iDefeito = String(item.defeito || "").toUpperCase(); 

        const matchSys = sys === "" || iSys === sys;
        const matchSearch = iTag.includes(search) || iDesc.includes(search);
        const matchLoc = loc === "" || iLoc === loc;
        
        let matchCalib = true;
        if (calibType === "SIM") matchCalib = iCalib.startsWith("SIM");
        else if (calibType === "NÃO") matchCalib = !iCalib.startsWith("SIM");
        else if (calibType === "REALIZADO") matchCalib = iCalib.startsWith("SIM") && iStatus.includes("OK");
        else if (calibType === "PENDENTE") matchCalib = iCalib.startsWith("SIM") && !iStatus.includes("OK");

        let matchDefeito = true;
        if (defType === "SIM") matchDefeito = iDefeito.includes("SIM");
        else if (defType === "OK") matchDefeito = !iDefeito.includes("SIM");

        return matchSys && matchSearch && matchLoc && matchCalib && matchDefeito;
    });

    updateTable(filtered);
    updateKPIs(filtered);
}

// =====================================================================
// RENDERIZAÇÃO DA TABELA E GRÁFICOS
// =====================================================================

function updateTable(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = "";

    if(data.length === 0) {
        // Atualizado colspan para 9 por conta da nova coluna
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:30px; color:#999;">Nenhum dado encontrado.</td></tr>`;
        document.getElementById('tableFooter').innerText = "";
        return;
    }

    data.forEach(item => {
        const tagClean = item.tag;
        const isPending = (tagClean === "");
        
        const statusHtml = isPending 
            ? '<span class="status-pill st-pend">PENDENTE</span>' 
            : '<span class="status-pill st-ok">OK</span>';
        
        const iCalib = String(item.calib).toUpperCase();
        const isCalib = iCalib.startsWith("SIM");
        const calibIcon = isCalib ? `<span class="calib-yes">SIM</span>` : `<span class="calib-no">${item.calib || "NÃO"}</span>`;

        const isDefeito = String(item.defeito).toUpperCase().includes("SIM");
        const defeitoHtml = isDefeito 
            ? '<span style="background-color:#f8d7da; color:#721c24; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em; border: 1px solid #f5c6cb;">🔴 DEFEITO</span>' 
            : '<span style="color:#28a745; font-weight:bold; font-size:0.9em;">🟢 OK</span>';

        // === NOVA LÓGICA PARA A OBSERVAÇÃO ===
        // Se tiver texto na observação, ele mostra em itálico e com cor cinza escuro. Se não, mostra um "-" clarinho.
        const obsHtml = item.observacao 
            ? `<span style="font-size:0.85em; color:#444; font-style:italic;">${item.observacao}</span>` 
            : `<span style="color:#ccc;">-</span>`;

        let sysClass = "sys-ar"; 
        let sysIcon = "🔧";
        let extraStyle = ""; 
        const sysName = String(item.sistema);
        
        if (sysName === "Efluentes") { sysClass = "sys-eflu"; sysIcon = "💧"; } 
        else if (sysName === "Ar Comprimido") { sysClass = "sys-ar"; sysIcon = "💨"; } 
        else if (sysName === "Ácido Sulfúrico") { sysClass = "sys-eflu"; sysIcon = "🧪"; } 
        else if (sysName === "Químicos") { sysClass = ""; sysIcon = "🧪"; extraStyle = "background-color: #8b5cf6; color: white;"; } 
        else if (sysName.startsWith("PW_LOOP")) {
            sysClass = ""; sysIcon = "💦";
            if (sysName.includes("1")) extraStyle = "background-color: #00bfff; color: white;"; 
            else if (sysName.includes("2")) extraStyle = "background-color: #1e90ff; color: white;"; 
            else if (sysName.includes("3")) extraStyle = "background-color: #0073e6; color: white;"; 
            else if (sysName.includes("4")) extraStyle = "background-color: #0059b3; color: white;"; 
            else if (sysName.includes("5")) extraStyle = "background-color: #004080; color: white;"; 
            else extraStyle = "background-color: #00bfff; color: white;"; 
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><span class="sys-badge ${sysClass}" style="${extraStyle}">${sysIcon} ${sysName}</span></td>
            <td class="tag-text">${isPending ? '<span style="color:var(--danger)">-- S/ TAG --</span>' : tagClean}</td>
            <td>${item.desc}</td>
            <td>${item.local}</td>
            <td>${calibIcon}</td>
            <td>${statusHtml}</td>
            <td>${defeitoHtml}</td>
            <td>${obsHtml}</td> <td style="font-size:0.8em; color:#999;">${item.origem}</td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('tableFooter').innerText = `Exibindo ${data.length} registros`;
}

function updateKPIs(data) {
    let totalItems = 0;
    let pendingTags = 0;
    let totalCritical = 0; 
    let totalDone = 0;
    let totalDefeitos = 0; 

    data.forEach(item => {
        totalItems++;
        if(!item.tag) pendingTags++;
        
        if(item.calib && String(item.calib).toUpperCase().startsWith("SIM")) {
            totalCritical++;
            if (item.status_calib && String(item.status_calib).toUpperCase().includes("OK")) {
                totalDone++;
            }
        }

        if(String(item.defeito).toUpperCase().includes("SIM")) {
            totalDefeitos++;
        }
    });

    animateValue("kpiTotal", totalItems);
    animateValue("kpiPending", pendingTags);
    animateValue("kpiCalib", totalCritical);
    
    const kpiDefEl = document.getElementById("kpiDefeito");
    if(kpiDefEl) animateValue("kpiDefeito", totalDefeitos);

    const chartElement = document.getElementById("chartDonut");
    if (chartElement) {
        const missing = totalCritical - totalDone;
        let percent = 0;
        if (totalCritical > 0) percent = Math.round((totalDone / totalCritical) * 100);

        document.getElementById("countDone").innerText = totalDone;
        document.getElementById("countMissing").innerText = missing;
        document.getElementById("chartPercent").innerText = percent + "%";
        chartElement.style.background = `conic-gradient(#2e7d32 0% ${percent}%, #e0e0e0 ${percent}% 100%)`;
    }
}

function animateValue(id, value) {
    const el = document.getElementById(id);
    if(el) el.innerText = value;
}

function limparDados() {
    if (rawData.length === 0) {
        alert("O painel já está limpo!");
        return;
    }

    if (confirm("Tem certeza que deseja limpar todos os dados da tela?")) {
        rawData = [];
        document.getElementById('excelInput').value = "";
        init();
    }
}

window.onload = init;