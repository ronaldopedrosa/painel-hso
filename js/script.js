// =====================================================================
// VARIÁVEL GLOBAL DE DADOS
// =====================================================================
let rawData = []; 

async function carregarExcel(input) {
    const files = input.files;
    if (!files || files.length === 0) return;

    const lerArquivoIndividual = (file) => {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    let worksheet = workbook.Sheets["BASE_CONSOLIDADA"] || workbook.Sheets[workbook.SheetNames[0]];

                    const jsonExcel = XLSX.utils.sheet_to_json(worksheet);

                    function getValue(row, possibleNames) {
                        const rowKeys = Object.keys(row);
                        const normalize = (str) => String(str).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "").toUpperCase();
                        for (let name of possibleNames) {
                            const target = normalize(name);
                            const foundKey = rowKeys.find(key => normalize(key) === target);
                            if (foundKey) return row[foundKey];
                        }
                        return ""; 
                    }

                    const dadosProcessados = jsonExcel.map(linha => ({
                        sistema: getValue(linha, ["Sala/Sistema", "Sala", "Sistema", "Area"]), 
                        tag: getValue(linha, ["TAG Hemobrás", "TAG", "Tag", "Instrumento", "Codigo"]),
                        desc: getValue(linha, ["Descrição dos Equipamentos", "Descrição", "Descricao", "Nome"]),
                        local: getValue(linha, ["Local", "Area"]), 
                        calib: getValue(linha, ["Calibração (SIM ou NÃO)", "Calibracao", "Criticidade"]),
                        status_calib: getValue(linha, ["Status de qualificação", "Status", "Situação"]),
                        certSM: getValue(linha, ["Certificado aprovado SM", "Certificado SM", "SM"]),
                        aprSVC: getValue(linha, ["Aprovado SVC", "Aprovado", "SVC"]),
                        defeito: getValue(linha, ["Defeito", "Com defeito", "Falha"]), 
                        observacao: getValue(linha, ["Observação", "Observacao", "Obs", "Motivo"]),
                        origem: getValue(linha, ["ORIGEM", "Origem"])
                    }));
                    resolve(dadosProcessados.filter(item => item.tag !== "" || item.desc !== ""));
                } catch (error) {
                    console.error("Erro ao ler arquivo:", file.name, error);
                    resolve([]); 
                }
            };
            reader.readAsArrayBuffer(file);
        });
    };

    const resultados = await Promise.all(Array.from(files).map(file => lerArquivoIndividual(file)));
    rawData = resultados.flat();
    alert(`SUCESSO! ${rawData.length} equipamentos carregados.`);
    init();
}

function init() {
    if (rawData.length === 0) {
        document.getElementById('tableBody').innerHTML = '<tr><td colspan="10" style="text-align:center; padding:20px; color:#0056b3;">📂 <b>Aguardando arquivos...</b></td></tr>';
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
        opt.value = loc; opt.innerText = loc;
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
        const sUp = sys.toUpperCase();
        if (sUp.includes("EFLUENTE") || sUp === "WW") emoji = "💧";
        else if (sUp.includes("AR COMP") || sUp === "CA" || sUp === "CAP") emoji = "💨";
        else if (sUp.includes("WFI") || sUp.includes("BWT") || sUp.includes("PW") || sUp.includes("LOOP")) emoji = "💦"; 
        opt.innerText = `${emoji} ${sys}`;
        select.appendChild(opt);
    });
}

function applyFilters() {
    const sys = document.getElementById('filterSystem').value;
    const search = document.getElementById('searchInput').value.toLowerCase();
    const loc = document.getElementById('filterLocal').value;
    const calibType = document.getElementById('filterCalib').value;
    const defType = document.getElementById('filterDefeito') ? document.getElementById('filterDefeito').value : "";

    const filtered = rawData.filter(item => {
        const iSys = String(item.sistema || "");
        const iTag = String(item.tag || "").toLowerCase();
        const iDesc = String(item.desc || "").toLowerCase();
        const iLoc = String(item.local || "").trim();
        const isCalib = String(item.calib || "").toUpperCase().startsWith("SIM");
        
        const statusQualif = String(item.status_calib || "").toUpperCase();
        const smText = String(item.certSM || "").toUpperCase();
        
        const isCalibradoOK = statusQualif === "OK";
        const isCertificadoAprovado = smText.includes("SIM") || smText.includes("OK") || smText.includes("APROVADO");

        // Lógica de Filtro Restrita 
        let matchCalib = true;
        if (calibType === "SIM") matchCalib = isCalib;
        else if (calibType === "NÃO") matchCalib = !isCalib;
        else if (calibType === "AGUARDANDO_CALIBRACAO") matchCalib = isCalib && !isCalibradoOK;
        else if (calibType === "CALIBRADOS") matchCalib = isCalib && isCalibradoOK;
        else if (calibType === "AGUARDANDO_SM") matchCalib = isCalib && !isCertificadoAprovado;
        else if (calibType === "CERTIFICADO_APROVADO") matchCalib = isCalib && isCertificadoAprovado;

        const iDefeito = String(item.defeito || "").toUpperCase();
        const matchDefeito = defType === "SIM" ? iDefeito.includes("SIM") : (defType === "OK" ? !iDefeito.includes("SIM") : true);

        return (sys === "" || iSys === sys) && 
               (iTag.includes(search) || iDesc.includes(search)) && 
               (loc === "" || iLoc === loc) && 
               matchCalib && matchDefeito;
    });

    updateTable(filtered);
    updateKPIs(filtered);
}

function updateTable(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = data.length === 0 ? '<tr><td colspan="10" style="text-align:center; padding:30px; color:#999;">Nenhum dado encontrado.</td></tr>' : "";

    data.forEach(item => {
        const isCalib = String(item.calib).toUpperCase().startsWith("SIM");
        const statusQualif = String(item.status_calib).toUpperCase();
        const smText = String(item.certSM).toUpperCase();
        
        // Regra: Progresso Calibração [cite: 15, 19]
        let progressoHtml = '<span style="color:#ccc;">-</span>';
        if (isCalib) {
            if (statusQualif === "OK") {
                progressoHtml = '<span style="background-color:#d4edda; color:#155724; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em;">Calibrado OK</span>';
            } else {
                progressoHtml = '<span style="background-color:#fff3cd; color:#856404; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em;">Aguardando Calibração</span>';
            }
        }

        // Regra: Status da Certificação [cite: 15, 19]
        let certHtml = '<span style="color:#ccc;">-</span>';
        if (isCalib) {
            if (smText.includes("SIM") || smText.includes("OK") || smText.includes("APROVADO")) {
                certHtml = '<span style="background-color:#cce5ff; color:#004085; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em;">Certificado Aprovado</span>';
            } else {
                certHtml = '<span style="background-color:#e2e3e5; color:#383d41; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em;">Aguardando Certificado</span>';
            }
        }

        const isDefeito = String(item.defeito).toUpperCase().includes("SIM");
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><span class="sys-badge" style="background-color:#007bff; color:white; padding:4px 8px; border-radius:8px;">${item.sistema}</span></td>
            <td>${item.tag || '<span style="color:red">S/ TAG</span>'}</td>
            <td>${item.desc}</td>
            <td>${item.local}</td>
            <td><b style="color:${isCalib ? '#28a745' : '#ccc'}">${isCalib ? 'SIM' : 'NÃO'}</b></td>
            <td style="text-align: center;">${progressoHtml}</td>
            <td style="text-align: center;">${certHtml}</td>
            <td>${isDefeito ? '<span style="color:red; font-weight:bold;">🔴 DEFEITO</span>' : '<span style="color:green;">🟢 OK</span>'}</td>
            <td><span style="font-size:0.85em; font-style:italic;">${item.observacao || "-"}</span></td>
            <td style="font-size:0.8em; color:#999;">${item.origem}</td>
        `;
        tbody.appendChild(tr);
    });
    document.getElementById('tableFooter').innerText = `Exibindo ${data.length} registros`;
}

function updateKPIs(data) {
    let totalItems = 0, totalCritical = 0, totalDone = 0, totalDefeitos = 0;

    data.forEach(item => {
        totalItems++;
        const isCalib = item.calib && String(item.calib).toUpperCase().startsWith("SIM");
        if(isCalib) {
            totalCritical++;
            // KPI baseado exclusivamente no "Status de Qualificação = OK" [cite: 15, 19]
            if (String(item.status_calib).toUpperCase() === "OK") totalDone++;
        }
        if(String(item.defeito).toUpperCase().includes("SIM")) totalDefeitos++;
    });

    document.getElementById("kpiTotal").innerText = totalItems;
    document.getElementById("kpiCalib").innerText = totalCritical;
    if(document.getElementById("kpiDefeito")) document.getElementById("kpiDefeito").innerText = totalDefeitos;

    const missing = totalCritical - totalDone;
    const percent = totalCritical > 0 ? Math.round((totalDone / totalCritical) * 100) : 0;
    document.getElementById("countDone").innerText = totalDone;
    document.getElementById("countMissing").innerText = missing;
    document.getElementById("chartPercent").innerText = percent + "%";
    document.getElementById("chartDonut").style.background = `conic-gradient(#28a745 0% ${percent}%, #e0e0e0 ${percent}% 100%)`;
}

window.onload = init;