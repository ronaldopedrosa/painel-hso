// =====================================================================
// VARIÁVEL GLOBAL DE DADOS
// =====================================================================
let rawData = []; 

// =====================================================================
// FUNÇÃO PARA LER MÚLTIPLOS ARQUIVOS EXCEL (FOCADO EM SALA/SISTEMA)
// =====================================================================
async function carregarExcel(input) {
    const files = input.files;
    if (!files || files.length === 0) return;

    let novosDados = [];

    const lerArquivoIndividual = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

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

                        // === REGRAS DE TRADUÇÃO INTELIGENTES (PREPARADO PARA LOOP 2) ===

                        // 1. ÁGUA PURIFICADA (PW_LOOP - 1 e PW_LOOP - 2)
                        if (sysText.includes("PW") || sysText.includes("LOOP") || sysText.includes("PURIFICADA")) {
                            if (sysText.includes("1")) {
                                finalSys = "PW_LOOP - 1";
                            } else if (sysText.includes("2")) {
                                finalSys = "PW_LOOP - 2";
                            } else {
                                finalSys = "PW_LOOP"; // Caso venha sem número
                            }
                        }
                        // 2. QUÍMICOS (HNO / HNA)
                        else if (sysText.includes("HNO") || sysText.includes("HNA") || sysText.includes("QUIMICO") || sysText.includes("QUÍMICO")) {
                            finalSys = "Químicos";
                        }
                        // 3. ÁCIDO SULFÚRICO (HSO)
                        else if (sysText.includes("HSO") || sysText.includes("SULFURICO") || sysText.includes("ACIDO") || sysText.includes("H2SO4")) {
                            finalSys = "Ácido Sulfúrico";
                        }
                        // 4. EFLUENTES (WW)
                        else if (sysText.includes("WW") || sysText.includes("EFLUENTE") || sysText.includes("ESGOTO") || sysText.includes("TRATAMENTO")) {
                            finalSys = "Efluentes";
                        }
                        // 5. AR COMPRIMIDO (CA/CAP)
                        else if (sysText.includes("CA-") || sysText.includes("-CA") || sysText.includes(" CAP ") || sysText.includes("AR COMP") || sysText.includes("COMPRIMIDO") || sysText === "CA" || sysText === "CAP") {
                            finalSys = "Ar Comprimido";
                        }
                        else if (rawSys.length > 2) {
                             finalSys = rawSys; 
                        }

                        return {
                            sistema: finalSys, 
                            tag: tag,
                            desc: desc,
                            local: getValue(linha, ["Local", "Area"]), 
                            calib: getValue(linha, ["Calibração (SIM ou NÃO)", "Calibracao", "Criticidade"]),
                            status_calib: getValue(linha, ["Status de qualificação", "Status", "Situação"]),
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
// INICIALIZAÇÃO
// =====================================================================

function init() {
    if (rawData.length === 0) {
        document.getElementById('tableBody').innerHTML = '<tr><td colspan="7" style="text-align:center; padding:20px; color:#0056b3;">📂 <b>Aguardando arquivos...</b><br>Selecione suas planilhas no botão acima.</td></tr>';
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
        if (sys === "Ácido Sulfúrico") emoji = "🧪";
        if (sys === "Químicos") emoji = "🧪"; 
        if (sys.startsWith("PW_LOOP")) emoji = "💦"; // Pega o 1 e o 2 automaticamente
        
        opt.innerText = `${emoji} ${sys}`;
        select.appendChild(opt);
    });
}

function applyFilters() {
    const sys = document.getElementById('filterSystem').value;
    const search = document.getElementById('searchInput').value.toLowerCase();
    const loc = document.getElementById('filterLocal').value;
    const calibType = document.getElementById('filterCalib').value;

    const filtered = rawData.filter(item => {
        const iSys = String(item.sistema || "");
        const iTag = String(item.tag || "").toLowerCase();
        const iDesc = String(item.desc || "").toLowerCase();
        const iLoc = String(item.local || "").trim();
        const iCalib = String(item.calib || "").toUpperCase();
        const iStatus = String(item.status_calib || "").toUpperCase();

        const matchSys = sys === "" || iSys === sys;
        const matchSearch = iTag.includes(search) || iDesc.includes(search);
        const matchLoc = loc === "" || iLoc === loc;
        
        let matchCalib = true;
        if (calibType === "SIM") matchCalib = iCalib.startsWith("SIM");
        else if (calibType === "NÃO") matchCalib = !iCalib.startsWith("SIM");
        else if (calibType === "REALIZADO") matchCalib = iCalib.startsWith("SIM") && iStatus.includes("OK");
        else if (calibType === "PENDENTE") matchCalib = iCalib.startsWith("SIM") && !iStatus.includes("OK");

        return matchSys && matchSearch && matchLoc && matchCalib;
    });

    updateTable(filtered);
    updateKPIs(filtered);
}

function updateTable(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = "";

    if(data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding:30px; color:#999;">Nenhum dado encontrado.</td></tr>`;
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
        const calibIcon = isCalib 
            ? `<span class="calib-yes">SIM</span>` 
            : `<span class="calib-no">${item.calib || "NÃO"}</span>`;

        let sysClass = "sys-ar"; 
        let sysIcon = "🔧";
        let extraStyle = ""; 
        const sysName = String(item.sistema);
        
        if (sysName === "Efluentes") { 
            sysClass = "sys-eflu"; 
            sysIcon = "💧"; 
        } else if (sysName === "Ar Comprimido") { 
            sysClass = "sys-ar"; 
            sysIcon = "💨"; 
        } else if (sysName === "Ácido Sulfúrico") {
            sysClass = "sys-eflu"; 
            sysIcon = "🧪";
        } else if (sysName === "Químicos") {
            sysClass = ""; 
            sysIcon = "🧪";
            extraStyle = "background-color: #8b5cf6; color: white;"; 
        } else if (sysName.startsWith("PW_LOOP")) {
            sysClass = ""; 
            sysIcon = "💦";
            // Lógica para diferenciar as cores do LOOP 1 e LOOP 2
            if (sysName.includes("1")) {
                extraStyle = "background-color: #00bfff; color: white;"; // Azul Claro Vibrante
            } else if (sysName.includes("2")) {
                extraStyle = "background-color: #1e90ff; color: white;"; // Dodger Blue (Azul um pouquinho mais escuro)
            } else {
                extraStyle = "background-color: #00bfff; color: white;"; 
            }
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><span class="sys-badge ${sysClass}" style="${extraStyle}">${sysIcon} ${sysName}</span></td>
            <td class="tag-text">${isPending ? '<span style="color:var(--danger)">-- S/ TAG --</span>' : tagClean}</td>
            <td>${item.desc}</td>
            <td>${item.local}</td>
            <td>${calibIcon}</td>
            <td>${statusHtml}</td>
            <td style="font-size:0.8em; color:#999;">${item.origem}</td>
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

    data.forEach(item => {
        totalItems++;
        if(!item.tag) pendingTags++;
        
        if(item.calib && String(item.calib).toUpperCase().startsWith("SIM")) {
            totalCritical++;
            if (item.status_calib && String(item.status_calib).toUpperCase().includes("OK")) {
                totalDone++;
            }
        }
    });

    animateValue("kpiTotal", totalItems);
    animateValue("kpiPending", pendingTags);
    animateValue("kpiCalib", totalCritical);

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