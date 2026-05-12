// =====================================================================
// VARIÁVEL GLOBAL DE DADOS E LINKS
// =====================================================================
let rawData = []; 

// Lembre-se de usar o GID da aba BASE_CONSOLIDADA e output=tsv
const urlsPlanilhas = [
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRfPfs_Uz4mQBzyZTIKu29vN_fDO5vv4yH6dAZBaURnWDZuFMCmoEwfAK7fEJPCdWSsH4n2hXrh3Gbi/pub?gid=2117011242&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRDR5EBxE2JHZOd8TiOMs-Oz9KHeg1w4ocS_36sUFBEWWMo5cimM9TbwpudxksctlXeURf_SEDzgJTR/pub?gid=1545016409&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRV3tVzKZlXxgdh2QsvYY_qHXWkSoDcxauP0Qyzt2xg9LrzpjMGKY7yHBCvElhQtlh4Umg-mVg4UoPf/pub?gid=723136747&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQIPNrWohFi7BFFJGEjNKbRbEhwrxHMG-z_bv2AMQjIFQZ09mqna_Un1zqenVkrQhErWgS0Hn1v6zkd/pub?gid=673959138&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRQI9b7YFfkhYPyTvA1YONAER7UmNUaOfpOYTRNxnLbggPN1AxBM-tifEHZBLa-LCuTD1fPhkKVR15G/pub?gid=1205138882&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTeQ4_IAPTifLW9kaqQgNcJ0WNp_HN38gOD3Wri3GQj9sSiki4QbpabsJxcKSEBADnWc9NctzZAGIwg/pub?gid=1095960966&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTgl-xkmWaFM-7pLOGzrOe4mru6TJg0IJ69mLOi1BR7dcuU09hBhM92lCY5lLI9cJwfyjpJ88Zqi_Kh/pub?gid=439725172&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRuTbW9-7dRN09-uOMPUfalkMSrstx9-yXoU0HFeLXJ6-gMDDxe7yoAVwWmCpWqnbp3vrlVqpvvh-K0/pub?gid=1779431803&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vSkAbkv7fvMa7T5ikDAcoLx91k0Q6nH-_FG8mQRUgDdAZ0rewBP2Y9N-ZoyTK6fsYLeGFCBeRqxYMlf/pub?gid=1321070825&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vStQ-dQGMiRBu6xWyRRna8X6lb4BUCy3jYRR_ZLHKvLcgNMnJOyPLS3C993D8AXFAFwWz0_v1AanAEh/pub?gid=1321003855&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vT5euNfPajNV6mwpRH9rvMkvKb2bzw25JEteGzCfsyXH_rPBbUA74ctAo_3OnkqgGTEIw8eVMcgC36q/pub?gid=312367184&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ50A1GhQW0wVQej_9VpVvIP8LFH3yGeaOFDAoM7FkQ3ufxLyRsvHJ2PYZAUe9p88plkY2kgGs_4rKB/pub?gid=835352193&single=true&output=tsv",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTpMooQXPhirFK2vEz_GhEFhttO_t9XW6tLvD3pd-H3zp4lO44Ic0Bs7E5woh1lBRpXi8S4ZmTR7hfJ/pub?gid=673959138&single=true&output=tsv",

];  

async function carregarDadosOnline() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center; padding:30px; color:#0056b3;">⏳ Baixando dados das nuvens...</td></tr>';
    
    let todosOsDados = [];

    try {
        const promessas = urlsPlanilhas.map(url => {
            if(!url.startsWith("http")) return null;
            return fetch(url).then(res => res.text());
        });
        
        const resultadosTSV = await Promise.all(promessas.filter(p => p !== null));
        
        resultadosTSV.forEach((tsvText, index) => {
            const separador = tsvText.includes("\t") ? "\t" : ",";
            const linhas = tsvText.split(/\r?\n/); 
            
            console.log(`Planilha ${index + 1}: ${linhas.length} linhas lidas.`);
            
            for (let i = 1; i < linhas.length; i++) {
                const col = linhas[i].split(separador); 
                
                if (col.length > 10 && col[1] && col[1].trim() !== "") {
                    todosOsDados.push({
                        sistema: col[3] ? col[3].trim() : "",        // Coluna D (Sala/Sistema detalhado)
                        tag: col[1] ? col[1].trim() : "",            // Coluna B (TAG)
                        local: col[2] ? col[2].trim() : "",          // Coluna C (Local)
                        desc: col[4] ? col[4].trim() : "",           // Coluna E (Descrição)
                        calib: col[20] ? col[20].trim() : "",        // Coluna U (Calibração SIM/NÃO)
                        status_calib: col[27] ? col[27].trim() : "", // Coluna AB (Status de Qualificação)
                        certSM: col[29] ? col[29].trim() : "",       // Coluna AD (Certificado SM)
                        defeito: col[31] ? col[31].trim() : "",      // Coluna AF (Defeito)
                        observacao: col[32] ? col[32].trim() : "",   // Coluna AG (Motivo)
                        bpf: col[33] ? col[33].trim() : "",          // Coluna AH (BPF)
                        macro: col[34] ? col[34].trim() : "",        // Coluna AI (Macro Sistema/Sigla)
                        origem: col[35] ? col[35].trim() : ""        // Coluna AJ (Origem)
                    });
                }
            }
        });

        rawData = todosOsDados;
        
        if (rawData.length > 0) {
            alert(`SUCESSO! ${rawData.length} equipamentos sincronizados com a nuvem.`);
            init();
        } else {
            tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:30px; color:#856404;">⚠️ Nenhum dado encontrado. Verifique se as planilhas estão publicadas.</td></tr>`;
        }

    } catch (error) {
        console.error("Erro ao carregar dados online:", error);
        tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:30px; color:red;">❌ Erro de conexão. Detalhe: ${error.message}</td></tr>`;
    }
}

function init() {
    if (rawData.length === 0) return;
    populateLocals();
    populateSystems();
    applyFilters();
}

function populateLocals() {
    const select = document.getElementById('filterLocal');
    select.innerHTML = '<option value="">Todos os Locais</option>'; 
    const locals = [...new Set(rawData.map(i => i.local))].filter(l => l).sort();
    locals.forEach(loc => {
        const opt = document.createElement('option');
        opt.value = loc; opt.innerText = loc;
        select.appendChild(opt);
    });
}

function populateSystems() {
    const select = document.getElementById('filterSystem');
    select.innerHTML = '<option value="">Mostrar Tudo</option>';
    
    // Filtro sendo montado apenas com a Sigla Enxuta
    const systems = [...new Set(rawData.map(i => i.macro))].filter(s => s).sort();
    
    systems.forEach(sys => {
        const opt = document.createElement('option');
        opt.value = sys;
        let emoji = "🔧";
        const sUp = sys.toUpperCase();
        if (sUp.includes("WW")) emoji = "💧";
        else if (sUp.includes("CA") || sUp.includes("CAP")) emoji = "💨";
        else if (sUp.includes("WFI") || sUp.includes("PW") || sUp.includes("LOOP")) emoji = "💦"; 
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
        const iMacro = String(item.macro).toUpperCase();
        const iTag = String(item.tag).toLowerCase();
        const iDesc = String(item.desc).toLowerCase();
        const iLoc = String(item.local);
        
        const isCalib = String(item.calib).toUpperCase().startsWith("SIM");
        
        // NOVO: Flexibilidade para o campo BPF aceitar SIM, OK ou BPF
        const isBPF = String(item.bpf).toUpperCase().includes("SIM") || 
                      String(item.bpf).toUpperCase().includes("OK") || 
                      String(item.bpf).toUpperCase().includes("BPF");
                      
        const statusQualif = String(item.status_calib).toUpperCase();
        const smText = String(item.certSM).toUpperCase();
        
        const isCalibradoOK = statusQualif === "OK";
        const isCertificadoAprovado = smText.includes("SIM") || smText.includes("OK") || smText.includes("APROVADO");

        let matchCalib = true;
        if (calibType === "SIM") matchCalib = isCalib;
        else if (calibType === "NÃO") matchCalib = !isCalib;
        else if (calibType === "BPF") matchCalib = isBPF;
        else if (calibType === "AGUARDANDO_CALIBRACAO") matchCalib = isCalib && !isCalibradoOK;
        else if (calibType === "CALIBRADOS") matchCalib = isCalib && isCalibradoOK;
        else if (calibType === "AGUARDANDO_SM") matchCalib = isCalib && !isCertificadoAprovado;
        else if (calibType === "CERTIFICADO_APROVADO") matchCalib = isCalib && isCertificadoAprovado;

        const iDefeito = String(item.defeito).toUpperCase();
        const matchDefeito = defType === "SIM" ? iDefeito.includes("SIM") : (defType === "OK" ? !iDefeito.includes("SIM") : true);

        return (sys === "" || iMacro === sys) && 
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
        
        // NOVO: Flexibilidade para o campo BPF na Tabela
        const isBPF = String(item.bpf).toUpperCase().includes("SIM") || 
                      String(item.bpf).toUpperCase().includes("OK") || 
                      String(item.bpf).toUpperCase().includes("BPF");
                      
        const statusQualif = String(item.status_calib).toUpperCase();
        const smText = String(item.certSM).toUpperCase();
        
        let progressoHtml = '<span style="color:#ccc;">-</span>';
        if (isCalib) {
            progressoHtml = statusQualif === "OK" 
                ? '<span style="background-color:#d4edda; color:#155724; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em;">Calibrado OK</span>'
                : '<span style="background-color:#fff3cd; color:#856404; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em;">Aguardando Calibração</span>';
        }

        let certHtml = '<span style="color:#ccc;">-</span>';
        if (isCalib) {
            certHtml = (smText.includes("SIM") || smText.includes("OK") || smText.includes("APROVADO"))
                ? '<span style="background-color:#cce5ff; color:#004085; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em;">Certificado Aprovado</span>'
                : '<span style="background-color:#e2e3e5; color:#383d41; padding:4px 8px; border-radius:12px; font-weight:bold; font-size:0.85em;">Aguardando Certificado</span>';
        }

        const isDefeito = String(item.defeito).toUpperCase().includes("SIM");
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <span class="sys-badge" style="background-color:#007bff; color:white; padding:4px 8px; border-radius:8px;" title="${item.sistema}">
                    ${item.macro || '<span style="font-size:0.8em">S/ MACRO</span>'}
                </span>
            </td>
            <td>
                ${item.tag || '<span style="color:red">S/ TAG</span>'}
                ${isBPF ? '<br><span style="font-size:0.7em; background:#8b5cf6; color:white; padding:2px 4px; border-radius:4px;">🛡️ BPF</span>' : ''}
            </td>
            <td><div style="max-width:250px; white-space:normal;" title="${item.desc}">${item.desc}</div></td>
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
    let totalItems = 0, totalCalib = 0, totalDone = 0, totalDefeitos = 0, totalBPF = 0;

    data.forEach(item => {
        totalItems++;
        if (String(item.calib).toUpperCase().startsWith("SIM")) {
            totalCalib++;
            if (String(item.status_calib).toUpperCase() === "OK") totalDone++;
        }
        if (String(item.defeito).toUpperCase().includes("SIM")) totalDefeitos++;
        
        // NOVO: Flexibilidade para o cálculo do KPI BPF
        const bpfText = String(item.bpf).toUpperCase();
        if (bpfText.includes("SIM") || bpfText.includes("OK") || bpfText.includes("BPF")) {
            totalBPF++;
        }
    });

    document.getElementById("kpiTotal").innerText = totalItems;
    document.getElementById("kpiCalib").innerText = totalCalib; 
    document.getElementById("kpiBPF").innerText = totalBPF; 
    if(document.getElementById("kpiDefeito")) document.getElementById("kpiDefeito").innerText = totalDefeitos;

    const percent = totalCalib > 0 ? Math.round((totalDone / totalCalib) * 100) : 0;
    document.getElementById("countDone").innerText = totalDone;
    document.getElementById("countMissing").innerText = totalCalib - totalDone;
    document.getElementById("chartPercent").innerText = percent + "%";
    document.getElementById("chartDonut").style.background = `conic-gradient(#28a745 0% ${percent}%, #e0e0e0 ${percent}% 100%)`;
}