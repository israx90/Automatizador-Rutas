
/* ========================================================================
   STATE — keeps parsed data so we can re-render when settings change
   ======================================================================== */
let currentDiagramData = null;
let currentSourceFileName = '';  // Stores uploaded DOCX filename for smart download naming

/** Multi-document state: { 'pagina_inicio': {data, fileName}, 'hito_1': ..., 'hito_5': ... } */
let allDocuments = {};
let currentHitoTab = null; // 'pagina_inicio' | 'hito_1' ... 'hito_5' | null (single-file mode)

/**
 * Generates a smart filename from the uploaded DOCX name.
 * Example: "04 Hito 3 Herramienta docente CGU-511 .docx" → "cgu-511_ruta_h3"
 * Logic:
 *  1. Extract the course code (pattern: letters-digits, e.g. CGU-511)
 *  2. Extract the hito number (pattern: Hito N)
 *  3. Combine as: code_ruta_hN
 *  Falls back to "ruta_aprendizaje" if patterns aren't found.
 */
function generateSmartFilename() {
    try {
        if (!currentSourceFileName) return 'ruta_aprendizaje';
        
        const name = currentSourceFileName.replace(/\.docx$/i, '').trim();
        
        // Extract course code (e.g. CGU-511, EIN-611, AOP-511)
        const skipWords = /^(hito|herramienta|docente|de|la|el|los|las|y)$/i;
        let codeMatch = null;
        const codeRegex = /([A-Za-z]{2,5})[\s\-_]+(\d{2,5})/gi;
        let m;
        while ((m = codeRegex.exec(name)) !== null) {
            if (!skipWords.test(m[1])) { codeMatch = m; break; }
        }
        const hitoMatch = name.match(/hito\s*(\d+)/i);
        
        // Página de Inicio document (no hito number, or has "pagina de inicio")
        const isPaginaInicio = currentDiagramData && currentDiagramData.diagram && currentDiagramData.diagram.type === 'paginaInicio';
        
        if (codeMatch) {
            const courseCode = (codeMatch[1] + '-' + codeMatch[2]).toLowerCase();
            if (isPaginaInicio) {
                return `${courseCode}_ruta_inicio`;
            }
            if (hitoMatch) {
                return `${courseCode}_ruta_h${hitoMatch[1]}`;
            }
            return courseCode + '_ruta';
        }
        
        return name.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '') || 'ruta_aprendizaje';
    } catch (e) {
        console.error("Filename generation error:", e);
        return 'ruta_aprendizaje';
    }
}

function generateCardFilename(hitoNum) {
    try {
        if (!currentSourceFileName) return `aprendizaje_h${hitoNum}`;
        const name = currentSourceFileName.replace(/\.docx$/i, '').trim();
        const skipWords = /^(hito|herramienta|docente|de|la|el|los|las|y)$/i;
        let codeMatch = null;
        const codeRegex = /([A-Za-z]{2,5})[\s\-_]+(\d{2,5})/gi;
        let m;
        while ((m = codeRegex.exec(name)) !== null) {
            if (!skipWords.test(m[1])) { codeMatch = m; break; }
        }
        if (codeMatch) {
            const courseCode = (codeMatch[1] + '-' + codeMatch[2]).toLowerCase();
            return `${courseCode}_aprendizaje_h${hitoNum}`;
        }
        return `aprendizaje_h${hitoNum}`;
    } catch (e) {
        return `aprendizaje_h${hitoNum}`;
    }
}

/* ========================================================================
   FILE UPLOAD (supports single and multiple DOCX files)
   ======================================================================== */
document.getElementById('docx-upload').addEventListener('change', handleFileSelect);

function handleFileSelect(event) {
    const files = Array.from(event.target.files);
    if (files.length === 0) return;

    // Reset file input so browser doesn't retain files
    event.target.value = '';

    currentActiveTab = 'diagram'; // Reset to diagram tab on new file

    const container = document.getElementById('diagram');
    container.innerHTML = `
        <div class="placeholder" style="animation: pulse 2s infinite">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="animation: spin 1s linear infinite"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
            <div>Analizando ${files.length > 1 ? files.length + ' documentos' : 'documento'}...</div>
        </div>
    `;
    document.getElementById('download-btn').disabled = true;

    if (files.length === 1) {
        // ── Single file mode (backward compatible) ──
        const file = files[0];
        currentSourceFileName = file.name;
        allDocuments = {};
        currentHitoTab = null;
        hideHitoTabs();

        if (file.name.toLowerCase().endsWith('.json')) {
            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    currentDiagramData = JSON.parse(e.target.result);
                    renderDiagram(currentDiagramData);
                    enableDownloads();
                } catch (err) {
                    console.error(err);
                    alert('Error al leer el archivo JSON.');
                    container.innerHTML = `<div class="placeholder"><div>Error al procesar el JSON</div></div>`;
                }
            };
            reader.readAsText(file);
        } else {
            const reader = new FileReader();
            reader.onload = function (e) {
                mammoth.extractRawText({ arrayBuffer: e.target.result })
                    .then(result => {
                        currentDiagramData = parseDocxToJson(result.value);
                        console.log('Parsed JSON:', JSON.stringify(currentDiagramData, null, 2));
                        renderDiagram(currentDiagramData);
                        enableDownloads();
                    })
                    .catch(err => {
                        console.error(err);
                        showUploadError(container);
                    });
            };
            reader.readAsArrayBuffer(file);
        }
    } else {
        // ── Multi-file mode: process all DOCX files ──
        allDocuments = {};
        currentHitoTab = null;

        const docxFiles = files.filter(f => f.name.toLowerCase().endsWith('.docx'));
        if (docxFiles.length === 0) {
            alert('No se encontraron archivos DOCX.');
            return;
        }

        // Extract the course code from the first filename (e.g., "CGU-511")
        const codeMatch = docxFiles[0].name.match(/[A-Z]{2,4}-?\d{3,4}/i);
        currentSourceFileName = codeMatch ? codeMatch[0].toUpperCase() : docxFiles[0].name;

        const promises = docxFiles.map(file => {
            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = function (e) {
                    mammoth.extractRawText({ arrayBuffer: e.target.result })
                        .then(result => {
                            const data = parseDocxToJson(result.value);
                            const type = detectDocumentType(file.name, data);
                            resolve({ file: file.name, type, data });
                        })
                        .catch(err => {
                            console.warn(`Error parsing ${file.name}:`, err);
                            resolve(null);
                        });
                };
                reader.readAsArrayBuffer(file);
            });
        });

        Promise.all(promises).then(results => {
            results.filter(Boolean).forEach(r => {
                allDocuments[r.type] = { data: r.data, fileName: r.file };
            });

            if (Object.keys(allDocuments).length === 0) {
                showUploadError(container);
                return;
            }

            console.log('Multi-doc loaded:', Object.keys(allDocuments));

            // Pick the first available tab
            const tabOrder = ['pagina_inicio', 'hito_1', 'hito_2', 'hito_3', 'hito_4', 'hito_5'];
            currentHitoTab = tabOrder.find(t => allDocuments[t]) || Object.keys(allDocuments)[0];

            const active = allDocuments[currentHitoTab];
            currentDiagramData = active.data;
            currentSourceFileName = active.fileName;

            renderHitoTabs();
            renderDiagram(currentDiagramData);
            enableDownloads();
        });
    }
}

/**
 * Auto-detect document type from filename and parsed content.
 * Returns: 'pagina_inicio' | 'hito_1' ... 'hito_5'
 */
function detectDocumentType(fileName, parsedData) {
    const fn = fileName.toLowerCase();

    // Detect Página de Inicio
    if (/p[aá]gina.*inicio/i.test(fn) || /01\s*p[aá]gina/i.test(fn)) {
        return 'pagina_inicio';
    }

    // Detect Hito by filename (e.g., "02 Hito 1", "Hito 3", "03 Hito 2")
    const hitoFileMatch = fn.match(/hito\s*(\d)/i);
    if (hitoFileMatch) {
        return `hito_${hitoFileMatch[1]}`;
    }

    // Fallback: detect from parsed content
    if (parsedData && parsedData.diagram) {
        if (parsedData.diagram.type === 'paginaInicio') return 'pagina_inicio';
        // Try to find hitoNum from the parsed diagram nodes
        const nodes = parsedData.diagram.nodes;
        if (nodes && nodes[0]) {
            const hitoBox = nodes[0];
            const titleMatch = (hitoBox.text?.title || '').match(/Hito\s+(\d)/i);
            if (titleMatch) return `hito_${titleMatch[1]}`;
        }
    }

    // Last resort: use position indicator from filename (02, 03, 04...)
    const posMatch = fn.match(/^(\d{2})\s/);
    if (posMatch) {
        const pos = parseInt(posMatch[1]);
        if (pos === 1) return 'pagina_inicio';
        if (pos >= 2 && pos <= 6) return `hito_${pos - 1}`;
    }

    return 'hito_1'; // ultimate fallback
}

/**
 * Render the primary Hito tab bar (only in multi-doc mode)
 */
function renderHitoTabs() {
    const tabBar = document.getElementById('hito-tabs');
    if (!tabBar) return;

    if (!currentHitoTab || Object.keys(allDocuments).length <= 1) {
        tabBar.style.display = 'none';
        return;
    }

    tabBar.style.display = 'flex';

    const tabLabels = {
        'pagina_inicio': 'Pag. Inicio',
        'hito_1': 'Hito 1',
        'hito_2': 'Hito 2',
        'hito_3': 'Hito 3',
        'hito_4': 'Hito 4',
        'hito_5': 'Hito 5',
    };

    const tabOrder = ['pagina_inicio', 'hito_1', 'hito_2', 'hito_3', 'hito_4', 'hito_5'];
    let html = '';
    tabOrder.forEach(key => {
        if (!allDocuments[key]) return;
        const isActive = currentHitoTab === key;
        const label = tabLabels[key] || key;
        const icon = key === 'pagina_inicio'
            ? '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>'
            : '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>';
        html += `<button class="tab-btn ${isActive ? 'active' : ''}" data-hito-tab="${key}">${icon} ${label}</button>`;
    });

    tabBar.innerHTML = html;

    tabBar.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const key = btn.getAttribute('data-hito-tab');
            switchHitoTab(key);
        });
    });
}

/**
 * Switch to a different Hito document
 */
function switchHitoTab(key) {
    if (!allDocuments[key]) return;
    currentHitoTab = key;
    currentDiagramData = allDocuments[key].data;
    currentSourceFileName = allDocuments[key].fileName;
    currentActiveTab = 'diagram'; // Always reset to diagram view

    renderHitoTabs();
    renderDiagram(currentDiagramData);
}

function hideHitoTabs() {
    const tabBar = document.getElementById('hito-tabs');
    if (tabBar) tabBar.style.display = 'none';
}

function enableDownloads() {
    document.getElementById('download-btn').disabled = false;
    document.getElementById('download-svg-btn').disabled = false;
    document.getElementById('download-zip-btn').disabled = false;
}

function showUploadError(container) {
    container.innerHTML = `
        <div class="placeholder">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242"/><path d="M12 12v9"/><path d="m8 16 4-4 4 4"/></svg>
            <div>Error al procesar el documento</div>
            <span style="color: #ef4444">Por favor revisa el formato e inténtalo de nuevo.</span>
        </div>
    `;
}

/* ========================================================================
   PARSER — Página de Inicio (Home Page / Course Overview)
   ======================================================================== */
function parsePaginaInicio(lines) {
    // Find the "Ruta de aprendizaje asignatura" section
    let rutaIdx = -1;
    for (let i = 0; i < lines.length; i++) {
        if (/Ruta\s+de\s+aprendizaje\s+asignatura/i.test(lines[i])) {
            rutaIdx = i;
            break;
        }
    }
    if (rutaIdx === -1) return null;

    // Extract course name from "Nombre de la asignatura: XXX"
    let courseName = '';
    for (let i = rutaIdx + 1; i < lines.length; i++) {
        const nameMatch = lines[i].match(/Nombre\s+de\s+la\s+asignatura\s*:\s*(.*)/i);
        if (nameMatch) {
            courseName = nameMatch[1].trim();
            break;
        }
    }

    // Try to extract a descriptive course name from the "Asignatura (Código y Nombre)" field
    let descriptiveName = '';
    for (let i = 0; i < rutaIdx; i++) {
        if (/Asignatura\s*\(C.digo\s+y\s+Nombre\)/i.test(lines[i])) {
            for (let j = i + 1; j < rutaIdx; j++) {
                if (lines[j].length > 0 && !/Carga\s+Horaria/i.test(lines[j])) {
                    descriptiveName = lines[j].trim();
                    break;
                }
            }
            break;
        }
    }

    let displayTitle = descriptiveName || courseName || 'Asignatura';

    // Parse Hito blocks: "Hito NN: Tipo" followed by subtitle on next line
    const hitos = [];
    for (let i = rutaIdx + 1; i < lines.length; i++) {
        const hitoMatch = lines[i].match(/^Hito\s+(\d+)\s*:\s*(.*)/i);
        if (hitoMatch) {
            const num = hitoMatch[1].replace(/^0+/, '') || hitoMatch[1];
            const type = hitoMatch[2].trim();
            let subtitle = '';
            if (i + 1 < lines.length && !/^Hito\s+\d+/i.test(lines[i + 1])) {
                subtitle = lines[i + 1].trim();
                i++;
            }
            hitos.push({ num, type, subtitle, cardBody: '', cardQuestion: '' });
        }
    }

    if (hitos.length === 0) return null;

    // ── Extract "Aprendizajes esperados" for tarjeta cards ──
    let aeIdx = -1;
    for (let i = 0; i < rutaIdx; i++) {
        if (/Aprendizajes\s+esperados/i.test(lines[i])) { aeIdx = i; break; }
    }

    if (aeIdx >= 0) {
        // Parse blocks between "HITO NN" markers
        // The DOCX groups them: HITO 01 / HITO 02 then competency texts, then Autoevalúate lines
        // Or HITO 05 alone with its texts
        const hitoBlocks = []; // { nums: [1,2], competencies: [], questions: [] }
        let currentBlock = null;

        for (let i = aeIdx + 1; i < rutaIdx; i++) {
            const hMarker = lines[i].match(/^HITO\s+(\d+)/i);
            if (hMarker) {
                const num = parseInt(hMarker[1]);
                // Check if next line is also a HITO marker (grouped pair)
                if (currentBlock && currentBlock.nums.length === 1 &&
                    i + 1 < rutaIdx && /^HITO\s+\d+/i.test(lines[i + 1])) {
                    // This is the start of a new paired block - save current first
                    if (currentBlock.competencies.length > 0 || currentBlock.questions.length > 0) {
                        hitoBlocks.push(currentBlock);
                    }
                }
                if (!currentBlock || currentBlock.competencies.length > 0 || currentBlock.questions.length > 0) {
                    // Start new block
                    if (currentBlock && (currentBlock.competencies.length > 0 || currentBlock.questions.length > 0)) {
                        hitoBlocks.push(currentBlock);
                    }
                    currentBlock = { nums: [num], competencies: [], questions: [] };
                } else {
                    // Add to existing block (paired hitos: HITO 01 / HITO 02)
                    currentBlock.nums.push(num);
                }
                continue;
            }

            if (!currentBlock) continue;

            const autoMatch = lines[i].match(/^Autoevalúate\s*:\s*(.*)/i);
            const preguntaMatch = lines[i].match(/^Pregunta\s+reflexiva\s*:\s*(.*)/i);
            if (autoMatch) {
                currentBlock.questions.push(autoMatch[1].trim());
            } else if (preguntaMatch) {
                currentBlock.questions.push(preguntaMatch[1].trim());
            } else if (lines[i].length > 10 && !/^Copia\s+de\s+la\s+herramienta/i.test(lines[i]) && !/^Redacta\s+un\s+p.rrafo/i.test(lines[i]) && !/^Si\s+consideras\s+que/i.test(lines[i])) {
                // Long enough to be a competency text (skip instructions)
                currentBlock.competencies.push(lines[i].trim());
            }
        }
        if (currentBlock && (currentBlock.competencies.length > 0 || currentBlock.questions.length > 0)) {
            hitoBlocks.push(currentBlock);
        }

        // Assign competencies and questions to the right hitos
        hitoBlocks.forEach(block => {
            block.nums.forEach((num, idx) => {
                const hito = hitos.find(h => parseInt(h.num) === num);
                if (hito) {
                    hito.cardBody = (block.competencies[idx] || block.competencies[0] || '').trim();
                    hito.cardQuestion = (block.questions[idx] || block.questions[0] || '').trim();
                }
            });
        });
    }

    return {
        diagram: {
            type: 'paginaInicio',
            title: displayTitle,
            courseCode: courseName,
            nodes: hitos.map((h, idx) => ({
                id: `hito_${idx}`,
                type: 'paginaInicioHito',
                text: {
                    hitoNum: h.num,
                    hitoType: h.type,
                    subtitle: h.subtitle
                },
                card: {
                    body: h.cardBody,
                    question: h.cardQuestion
                }
            }))
        }
    };
}

/* ========================================================================
   PARSER — Individual Hito Documents
   ======================================================================== */
/**
 * Auto-fix common Spanish punctuation, spacing & accent errors from DOCX sources.
 * Runs on every parsed text line to catch issues like "¿ Cómo" → "¿Cómo".
 */
function normalizeSpanishText(text) {
    if (!text) return text;
    return text
        // Remove space after opening ¿ and ¡
        .replace(/([¿¡])\s+/g, '$1')
        // ── Interrogative/exclamatory accent corrections ──
        // "¿Que " / "¿que " → "¿Qué " (also handles ¡Que)
        .replace(/([¿¡])(?:Q|q)ue\b/g, (m, p) => p + (p === '¿' ? 'Qué' : 'Qué'))
        // "¿Por que" → "¿Por qué"
        .replace(/([¿¡])[Pp]or que\b/g, '$1Por qué')
        // "¿Como " → "¿Cómo"
        .replace(/([¿¡])[Cc]omo\b/g, '$1Cómo')
        // "¿Cuando " → "¿Cuándo"
        .replace(/([¿¡])[Cc]uando\b/g, '$1Cuándo')
        // "¿Donde " → "¿Dónde"
        .replace(/([¿¡])[Dd]onde\b/g, '$1Dónde')
        // "¿Cual " → "¿Cuál" / "¿Cuales" → "¿Cuáles"
        .replace(/([¿¡])[Cc]ual\b/g, '$1Cuál')
        .replace(/([¿¡])[Cc]uales\b/g, '$1Cuáles')
        // "¿Quien " → "¿Quién" / "¿Quienes" → "¿Quiénes"
        .replace(/([¿¡])[Qq]uien\b/g, '$1Quién')
        .replace(/([¿¡])[Qq]uienes\b/g, '$1Quiénes')
        // "¿Cuanto/a/os/as" → "¿Cuánto/a/os/as"
        .replace(/([¿¡])[Cc]uant([oa]s?)\b/g, '$1Cuánt$2')
        // ── Spacing fixes ──
        // Remove space before colon (but keep the colon attached): "Autor : X" → "Autor: X"
        .replace(/\s+:/g, ':')
        // Remove space before closing punctuation: "texto , más" → "texto, más"
        .replace(/\s+([.,;!?\)])/g, '$1')
        // Ensure space after colon/semicolon/period/comma if followed by a letter
        .replace(/([.,;:!?\)])([A-Za-záéíóúñÁÉÍÓÚÑ])/g, '$1 $2')
        // Fix double periods (not ellipsis) → single period
        .replace(/\.\.(?!\.)/g, '.')
        // Remove space after opening quote: "" texto" → ""texto"
        .replace(/(["«""])\s+/g, '$1')
        // Remove space before closing quote: "texto "" → "texto""
        .replace(/\s+(["»""])/g, '$1')
        // Collapse multiple spaces into one
        .replace(/  +/g, ' ')
        // Trim
        .trim();
}

function parseDocxToJson(text) {
    const lines = text.split('\n').map(l => normalizeSpanishText(l.trim())).filter(l => l.length > 0);

    // ── Early detection: Is this a "Página de Inicio" document? ──
    const paginaResult = parsePaginaInicio(lines);
    if (paginaResult) return paginaResult;

    let startIdx = -1;
    let isHito1 = false;
    
    // First, look for any Hito header: "HITO NN KEYWORD(S): subtitle"
    // Matches: AVANCE y DESARROLLO, INVESTIGACIÓN, INTEGRACIÓN, DIAGNÓSTICO, etc.
    // Uses LAST occurrence to skip TOC entries
    // Tolerates missing space: "HITO 4AVANCE" as well as "HITO 4 AVANCE"
    // IMPORTANT: requires a colon ":" in the line OR a known keyword to avoid
    // matching internal section titles like "Hito 02 Formas de Organización..."
    const knownHitoKeywords = /AVAN|DIAGN|INVESTIG|INTEGRA|DESARROL|INNOVACI/i;
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/^HITO\s*\d+\s*(?!Semana)[A-ZÁÉÍÓÚÑa-záéíóúñ]/i.test(line) &&
            !/^HITO\s*\d+\s*Semana/i.test(line) &&
            !/Página\s+Principal/i.test(line) &&
            (line.includes(':') || knownHitoKeywords.test(line))) {
            startIdx = i;
        }
    }
    
    // If not found, look for Hito 1 specific format from screenshot (find LAST occurrence to skip TOC)
    if (startIdx === -1) {
        for (let i = 0; i < lines.length; i++) {
            if (/Ruta\s+de\s+aprendizaje\s+Hito\s+0?1/i.test(lines[i])) {
                startIdx = i;
                isHito1 = true;
            }
        }
    }

    if (startIdx === -1) {
        // Fallback: Return a completely empty UNKNOWN structure so the user can just manually edit it
        return {
            diagram: {
                title: "Hito X",
                nodes: [{
                    id: 'hito_main',
                    type: 'hitoBox',
                    text: { title: "Hito X" },
                    children: [{
                        id: 'avance_main',
                        type: 'mainNode',
                        text: { title: 'Avance y desarrollo:', subtitle: 'UNKNOWN' },
                        children: [{
                            id: 'week_0',
                            type: 'finalWeekNode',
                            text: { title: 'Semana X', subtitle: 'UNKNOWN' },
                            children: [{
                                id: 'act_0_0',
                                type: 'finalActivityNode',
                                text: { code: 'SXX', title: 'UNKNOWN:', subtitle: 'UNKNOWN' }
                            }]
                        }]
                    }]
                }]
            }
        };
    }

    let hitoNum = '3';
    const hitoMatch = lines[startIdx].match(/HITO\s*(\d+)/i);
    if (hitoMatch) {
        hitoNum = hitoMatch[1].replace(/^0+/, '');
        if (hitoNum === '1') isHito1 = true;
    } else if (isHito1) {
        hitoNum = '1';
    }

    let avanceSubtitle = '';
    let avanceTitle = 'Avance y desarrollo:';

    if (isHito1) {
        avanceTitle = 'Diagnóstico:';
        // First, try to extract subtitle from the header line itself
        // e.g. "Hito 01 DIAGNÓSTICO: Conozcamos tus derechos políticos"
        const headerDiagMatch = lines[startIdx].match(/^HITO\s+\d+\s+[A-ZÁÉÍÓÚÑa-záéíóúñ\s]+:\s*(.*)/i);
        if (headerDiagMatch && headerDiagMatch[1]) {
            avanceSubtitle = toTitleCaseES(headerDiagMatch[1].replace(/_/g, '').trim());
        }
        // If not found in header, look for a separate "Hito 01: subtitle" line
        if (!avanceSubtitle) {
            for (let i = startIdx + 1; i < lines.length; i++) {
                if (/^Semana\s+\d+/i.test(lines[i])) break; // Stop if we hit weeks
                
                const diagMatch = lines[i].match(/^Hito\s+0?1\s*:\s*(.*)/i);
                if (diagMatch && diagMatch[1]) {
                    avanceSubtitle = toTitleCaseES(diagMatch[1].replace(/_/g, '').trim());
                    break;
                }
            }
        }
    } else {
        // Dynamically extract avance title from the header line
        // e.g. "HITO 03 INVESTIGACIÓN: Mentalidad emprendedora..." → title="Investigación:", subtitle="Mentalidad..."
        // e.g. "HITO 3 AVANCE y DESARROLLO: Sistema..."            → title="Avance y Desarrollo:", subtitle="Sistema..."
        // Tolerates missing space: "HITO 4AVANCE y DESARROLLO: ..." → still works
        const headerMatch = lines[startIdx].match(/^HITO\s*\d+\s*(.+?)\s*:\s*(.*)/i);
        if (headerMatch) {
            const rawTitle = headerMatch[1].trim();
            avanceTitle = toTitleCaseES(rawTitle) + ':';
            avanceSubtitle = toTitleCaseES((headerMatch[2] || '').trim());
        }
        
        // Also collect any lines between the header and first "Semana XX:" as continuation
        for (let i = startIdx + 1; i < lines.length; i++) {
            if (/^Semana\s+\d+/i.test(lines[i])) break;
            if (avanceSubtitle) avanceSubtitle += ' ';
            avanceSubtitle += lines[i];
        }
    }

    let semanasStart = startIdx + 1;
    for (let i = startIdx + 1; i < lines.length; i++) {
        if (/^Semana\s+\d+/i.test(lines[i])) {
            semanasStart = i;
            break;
        }
    }

    const semanas = [];
    let currentSemana = null;

    for (let i = semanasStart; i < lines.length; i++) {
        const line = lines[i];
        if (/^(¡Qué interesante|Amplía tu vocabulario|Formulario Hito|Continúa el llenado)/i.test(line)) break;

        const weekMatch = line.match(/^Semana\s+(\d+)\s*:?\s*(.*)/i);
        if (weekMatch) {
            let num = weekMatch[1].replace(/^0+/, '') || weekMatch[1];
            let subtitle = weekMatch[2] ? weekMatch[2].trim() : '';
            // Strip trailing page numbers (e.g. from TOC: "Diagnóstico Inicial3" → "Diagnóstico Inicial")
            subtitle = subtitle.replace(/\d+$/, '').trim();
            if (!subtitle && i + 1 < lines.length && !/^(S\d{2}|Semana)/i.test(lines[i + 1])) {
                subtitle = lines[i + 1].replace(/\d+$/, '').trim();
                i++;
            }
            if (!subtitle) subtitle = 'UNKNOWN'; // User requested fallback
            currentSemana = { num, subtitle: toTitleCaseES(subtitle), activities: [] };
            semanas.push(currentSemana);
            continue;
        }

        const actMatch = line.match(/^(S\d{2})\s+(.*)/i);
        if (actMatch && currentSemana) {
            const code = actMatch[1].toUpperCase();
            let rawName = actMatch[2].trim();

            let extras = [];
            let j = i + 1;
            while (j < lines.length &&
                !/^(S\d{2}|Semana)/i.test(lines[j]) &&
                !/^(¿Por qué|¡Qué interesante|Amplía tu vocabulario|Formulario Hito|Continúa el llenado|FINAL DEL)/i.test(lines[j])) {
                extras.push(lines[j]);
                j++;
            }
            i = j - 1;

            let title, subtitle;
            if (rawName.toLowerCase() === 'contenido') {
                title = 'Contenido';
                subtitle = '';
            } else if (/^pregunta\s+reflexiva/i.test(rawName)) {
                title = 'Autoevalúate:';
                subtitle = '';
            } else if (rawName.includes(':')) {
                title = rawName.split(':')[0].trim() + ':';
                subtitle = rawName.substring(rawName.indexOf(':') + 1).trim();
                if (extras.length > 0) subtitle += ' ' + extras.join(' ');
            } else {
                // Try to split at known activity keyword boundary (handles missing colon)
                const actKeywords = /^(Afianzamiento|Taller(?:\s+(?:de\s+)?\w+)?|Foro|Proyecto|Gamificación|Aplicación|Desempeño(?:\s+presencial)?|Autoevalúate|Pregunta\s+reflexiva|Solucionario)\s*(.*)/i;
                const kwMatch = rawName.match(actKeywords);
                if (kwMatch) {
                    title = kwMatch[1].trim() + ':';
                    subtitle = (kwMatch[2] || '').trim();
                    if (extras.length > 0) subtitle += (subtitle ? ' ' : '') + extras.join(' ');
                } else {
                    title = rawName ? rawName + ':' : 'UNKNOWN';
                    subtitle = extras.length > 0 ? extras.join(' ') : 'UNKNOWN';
                }
            }
            
            if (!subtitle && title !== 'Contenido' && title !== 'Autoevalúate:') subtitle = 'UNKNOWN';

            currentSemana.activities.push({ code, title: toTitleCaseES(title), subtitle: toTitleCaseES(subtitle.trim()) });
            continue;
        }

        if (line.includes("FINAL DEL HITO")) break;
    }

    // Post-process: override semana subtitles with "Título de la Semana XX:" if found
    // (these have the definitive/correct titles vs the abbreviated ruta section)
    for (let i = 0; i < lines.length; i++) {
        const tituloMatch = lines[i].match(/^T[ií]tulo\s+de\s+la\s+Semana\s+(\d+)\s*:\s*(.*)/i);
        if (tituloMatch) {
            const semNum = tituloMatch[1].replace(/^0+/, '') || tituloMatch[1];
            const newTitle = tituloMatch[2].trim();
            if (newTitle) {
                const sem = semanas.find(s => s.num === semNum);
                if (sem) sem.subtitle = toTitleCaseES(newTitle);
            }
        }
    }

    if (semanas.length === 0) {
        semanas.push({
            num: 'X', subtitle: 'UNKNOWN', activities: [{ code: 'SXX', title: 'UNKNOWN:', subtitle: 'UNKNOWN' }]
        });
    }

    const totalWeeks = semanas.length;

    const weekNodes = semanas.map((sem, idx) => {
        const isFinal = idx === totalWeeks - 1;
        
        // Ensure at least one activity exists so the box renders something to click
        if (!sem.activities || sem.activities.length === 0) {
            sem.activities = [{ code: 'SXX', title: 'UNKNOWN:', subtitle: 'UNKNOWN' }];
        }
        
        const weekObj = {
            id: `week_${idx}`,
            type: isFinal ? 'finalWeekNode' : 'weekNode',
            text: { title: `Semana ${sem.num}`, subtitle: sem.subtitle },
            children: sem.activities.map((act, aIdx) => {
                const isContent = act.title === 'Contenido';
                let type;
                if (isFinal) type = isContent ? 'finalContentNode' : 'finalActivityNode';
                else type = isContent ? 'contentNode' : 'activityNode';
                const actObj = {
                    id: `act_${idx}_${aIdx}`,
                    type,
                    text: { code: act.code, title: act.title, subtitle: act.subtitle || undefined }
                };
                return actObj;
            })
        };
        return weekObj;
    });

    return {
        diagram: {
            title: `Hito ${hitoNum}`,
            isHito1: isHito1,
            nodes: [{
                id: 'hito_main',
                type: 'hitoBox',
                text: { title: `Hito ${hitoNum}` },
                children: [{
                    id: 'avance_main',
                    type: 'mainNode',
                    text: { title: avanceTitle, subtitle: avanceSubtitle || 'UNKNOWN' },
                    children: weekNodes
                }]
            }]
        }
    };
}

/* ========================================================================
   TEXT UTILITIES — English Italicization & Title Case
   ======================================================================== */

/**
 * Dictionary of English words commonly used in Spanish educational/business texts.
 * These will be rendered in italics in SVG output.
 */
const ENGLISH_ITALICS = new Set([
    // Business & Strategy
    'planner', 'canvas', 'pitch', 'startup', 'startups', 'feedback', 'coaching',
    'mentoring', 'networking', 'brainstorming', 'benchmarking', 'crowdfunding',
    'lean', 'sprint', 'sprints', 'scrum', 'agile', 'stakeholder', 'stakeholders',
    'mindset', 'insight', 'insights', 'workshop', 'workshops', 'storytelling',
    'branding', 'naming', 'briefing', 'crowdsourcing', 'outsourcing', 'coworking',
    'know-how', 'empowerment', 'hub', 'cluster', 'co-working',
    // Design & Innovation
    'design', 'thinking', 'growth', 'hacking', 'prototype', 'prototyping',
    'wireframe', 'wireframes', 'mockup', 'mockups', 'layout',
    // Project Management
    'planning', 'roadmap', 'backlog', 'stand-up', 'standup', 'milestone',
    'milestones', 'deadline', 'deadlines', 'deliverable', 'deliverables',
    'framework', 'frameworks', 'checklist', 'guidelines', 'performance',
    'target', 'targets', 'meeting', 'meetings', 'team', 'teams',
    'leader', 'leaders', 'manager',
    // Marketing & Sales
    'engagement', 'influencer', 'influencers', 'inbound', 'outbound',
    'buyer', 'customer', 'journey', 'brochure', 'flyer', 'banner',
    'display', 'spot', 'jingle',
    // Tech
    'software', 'hardware', 'cloud', 'streaming', 'podcast', 'podcasts',
    'big', 'data', 'machine', 'learning', 'blockchain', 'fintech', 'edtech',
    'e-commerce', 'ecommerce', 'dropshipping', 'marketplace',
    // Academic
    'paper', 'papers', 'abstract', 'poster', 'posters', 'review',
    'peer', 'state-of-the-art',
    // General business English
    'training', 'trading', 'holding', 'ranking', 'core', 'smart',
    'retail', 'wholesale', 'supply', 'chain', 'stock',
    'demo', 'testing', 'deploy', 'deployment',
]);

/** Subject/course names that are English but should NOT be italicized */
const ENGLISH_NO_ITALIC = new Set([
    'marketing', 'business',
]);

/**
 * Wraps English words in SVG <tspan font-style="italic"> for rendering.
 * Consecutive English words share one tspan (e.g. "Design Thinking").
 */
function italicizeSvgText(text) {
    if (!text) return text;
    const parts = text.split(/(\s+)/);
    let out = '';
    let inItalic = false;
    for (const part of parts) {
        if (/^\s*$/.test(part)) {
            out += part;
            continue;
        }
        // Strip leading/trailing punctuation for dictionary lookup
        const clean = part.replace(/^[¿¡"'«"(\[]+|[.,;:!?"'»")\]]+$/g, '');
        const isEng = clean.length > 1 &&
            ENGLISH_ITALICS.has(clean.toLowerCase()) &&
            !ENGLISH_NO_ITALIC.has(clean.toLowerCase());
        if (isEng && !inItalic) {
            out += '<tspan font-style="italic">';
            inItalic = true;
        } else if (!isEng && inItalic) {
            out += '</tspan>';
            inItalic = false;
        }
        out += part;
    }
    if (inItalic) out += '</tspan>';
    return out;
}

/**
 * Checks if a single word should be italic (for per-word rendering like justify).
 */
function isEnglishWord(word) {
    if (!word) return false;
    const clean = word.replace(/^[¿¡"'«"(\[]+|[.,;:!?"'»")\]]+$/g, '');
    return clean.length > 1 &&
        ENGLISH_ITALICS.has(clean.toLowerCase()) &&
        !ENGLISH_NO_ITALIC.has(clean.toLowerCase());
}

/* ========================================================================
   RICH TEXT EDITOR MODULE
   ======================================================================== */
const FONT_OPTIONS = [
    { value: "'Montserrat', sans-serif", label: 'Montserrat' },
    { value: "'Inter', sans-serif", label: 'Inter' },
    { value: "'Roboto', sans-serif", label: 'Roboto' },
    { value: "'Open Sans', sans-serif", label: 'Open Sans' },
    { value: "'Poppins', sans-serif", label: 'Poppins' },
    { value: "'Lato', sans-serif", label: 'Lato' },
    { value: "Arial, sans-serif", label: 'Arial' },
];

function createRichFieldHTML(fieldId, htmlContent, options = {}) {
    const { placeholder = 'Texto...', font = '', size = '' } = options;
    const fontOpts = FONT_OPTIONS.map(f =>
        `<option value="${f.value}" ${f.value === font ? 'selected' : ''}>${f.label}</option>`
    ).join('');
    return `
    <div class="rich-text-field" data-field-id="${fieldId}">
      <div class="rich-toolbar">
        <select class="rt-font" title="Fuente">${fontOpts}</select>
        <input type="number" class="rt-size" value="${size || 14}" min="8" max="72" step="1" title="Tamaño (px)">
        <span class="rt-separator"></span>
        <button class="rt-btn rt-bold" title="Bold (Ctrl+B)" type="button">B</button>
        <button class="rt-btn rt-italic" title="Italic (Ctrl+I)" type="button">I</button>
        <span class="rt-separator"></span>
        <button class="rt-btn rt-case" data-case="title" title="Title Case" type="button">Aa</button>
        <button class="rt-btn rt-case" data-case="lower" title="minúsculas" type="button">aa</button>
        <button class="rt-btn rt-case" data-case="upper" title="MAYÚSCULAS" type="button">AA</button>
      </div>
      <div class="rich-editor" contenteditable="true" data-field-id="${fieldId}" data-placeholder="${placeholder}" spellcheck="true" lang="es">${htmlContent || ''}</div>
    </div>`;
}

function parseHtmlToSegments(html) {
    if (!html || !html.trim()) return [];
    const div = document.createElement('div');
    div.innerHTML = html;
    const segments = [];
    function walk(node, style) {
        if (node.nodeType === 3) {
            const text = node.textContent;
            if (text) segments.push({ text, ...style });
            return;
        }
        if (node.nodeType !== 1) return;
        const tag = node.tagName.toLowerCase();
        const s = { ...style };
        if (tag === 'b' || tag === 'strong') s.bold = true;
        if (tag === 'i' || tag === 'em') s.italic = true;
        if (tag === 'br') { segments.push({ text: '\n' }); return; }
        if (tag === 'div' || tag === 'p') {
            if (segments.length > 0 && segments[segments.length - 1].text !== '\n') segments.push({ text: '\n' });
        }
        const cs = node.style;
        if (cs && (cs.fontWeight === 'bold' || parseInt(cs.fontWeight) >= 600)) s.bold = true;
        if (cs && cs.fontStyle === 'italic') s.italic = true;
        for (const child of node.childNodes) walk(child, s);
    }
    walk(div, {});
    while (segments.length > 0 && segments[segments.length - 1].text === '\n') segments.pop();
    return segments;
}

function segmentsToPlainText(segments) { return segments.map(s => s.text).join(''); }

function escXml(t) { return t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function renderRichSvgText(segments, x, y, maxWidth, defaults, measureCtx) {
    if (!segments || segments.length === 0) return { svg: '', height: 0 };
    const SAFETY = 1.08, LH = 1.35;
    const fontSize = defaults.size, fontFamily = defaults.font, defWeight = defaults.weight || '400';
    const words = [];
    for (const seg of segments) {
        if (seg.text === '\n') { words.push({ text: '\n', lb: true }); continue; }
        const parts = seg.text.split(/(\s+)/);
        for (const p of parts) { if (p.length > 0) words.push({ text: p, bold: seg.bold, italic: seg.italic }); }
    }
    const lines = [[]];
    let lw = 0;
    for (const w of words) {
        if (w.lb) { lines.push([]); lw = 0; continue; }
        const fw = w.bold ? '700' : defWeight;
        const fi = w.italic ? 'italic ' : '';
        measureCtx.font = `${fi}${fw} ${fontSize}px ${fontFamily}`;
        const ww = measureCtx.measureText(w.text).width * SAFETY;
        const isSpc = /^\s+$/.test(w.text);
        if (!isSpc && lw + ww > maxWidth && lines[lines.length - 1].length > 0) {
            const cl = lines[lines.length - 1];
            while (cl.length && /^\s+$/.test(cl[cl.length - 1].text)) cl.pop();
            lines.push([]);
            lw = 0;
        }
        lines[lines.length - 1].push(w);
        lw += ww;
    }
    let svgStr = '', curY = y;
    const slotH = fontSize * LH;
    for (const line of lines) {
        if (line.length === 0) { curY += slotH; continue; }
        const cY = curY + slotH / 2;
        let tspans = '', buf = '', prevKey = null;
        for (const w of line) {
            const key = (w.bold ? 'b' : '') + (w.italic ? 'i' : '');
            if (prevKey !== null && prevKey !== key) { tspans += _buildRichTspan(buf, prevKey, defWeight); buf = ''; }
            buf += w.text;
            prevKey = key;
        }
        if (buf) tspans += _buildRichTspan(buf, prevKey, defWeight);
        svgStr += `<text x="${x}" y="${cY}" font-family="${fontFamily}" font-size="${fontSize}" fill="${defaults.color}" dominant-baseline="central">${tspans}</text>`;
        curY += slotH;
    }
    return { svg: svgStr, height: curY - y };
}

function _buildRichTspan(text, styleKey, defWeight) {
    const isBold = styleKey && styleKey.includes('b');
    const isItalic = styleKey && styleKey.includes('i');
    let attrs = ` font-weight="${isBold ? '700' : defWeight}"`;
    if (isItalic) attrs += ` font-style="italic"`;
    const escaped = escXml(text);
    const content = isItalic ? escaped : italicizeSvgText(escaped);
    return `<tspan${attrs}>${content}</tspan>`;
}

function bindRichTextEvents(container, getNode, rerender) {
    container.querySelectorAll('.rich-text-field').forEach(field => {
        const fieldId = field.getAttribute('data-field-id');
        const toolbar = field.querySelector('.rich-toolbar');
        const editor = field.querySelector('.rich-editor');
        if (!toolbar || !editor) return;

        toolbar.querySelector('.rt-bold')?.addEventListener('click', (e) => {
            e.preventDefault(); editor.focus();
            document.execCommand('bold', false, null);
            _updateRtBtnState(toolbar);
            _onRichChange(editor, fieldId, getNode, rerender);
        });
        toolbar.querySelector('.rt-italic')?.addEventListener('click', (e) => {
            e.preventDefault(); editor.focus();
            document.execCommand('italic', false, null);
            _updateRtBtnState(toolbar);
            _onRichChange(editor, fieldId, getNode, rerender);
        });
        toolbar.querySelectorAll('.rt-case').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                const sel = window.getSelection();
                if (!sel.rangeCount) return;
                const text = sel.toString();
                if (!text) return;
                const ct = btn.getAttribute('data-case');
                let result;
                if (ct === 'upper') result = text.toUpperCase();
                else if (ct === 'lower') result = text.toLowerCase();
                else result = toTitleCaseES(text.toUpperCase());
                editor.focus();
                document.execCommand('insertText', false, result);
                _onRichChange(editor, fieldId, getNode, rerender);
            });
        });
        toolbar.querySelector('.rt-font')?.addEventListener('change', (e) => {
            const info = getNode(fieldId);
            if (info && info.node) { info.node.customFont = e.target.value; rerender(); }
        });
        toolbar.querySelector('.rt-size')?.addEventListener('change', (e) => {
            const info = getNode(fieldId);
            if (info && info.node) {
                const sv = parseInt(e.target.value);
                if (info.fieldKey === 'title') info.node.customTitleSize = sv;
                else if (info.fieldKey === 'subtitle') info.node.customSubtitleSize = sv;
                else if (info.fieldKey === 'description') info.node.customDescriptionSize = sv;
                rerender();
            }
        });
        editor.addEventListener('input', () => _onRichChange(editor, fieldId, getNode, rerender));
        editor.addEventListener('mouseup', () => _updateRtBtnState(toolbar));
        editor.addEventListener('keyup', () => _updateRtBtnState(toolbar));
        editor.addEventListener('keydown', (e) => {
            if (e.ctrlKey || e.metaKey) {
                if (e.key === 'b') { e.preventDefault(); document.execCommand('bold'); _updateRtBtnState(toolbar); _onRichChange(editor, fieldId, getNode, rerender); }
                if (e.key === 'i') { e.preventDefault(); document.execCommand('italic'); _updateRtBtnState(toolbar); _onRichChange(editor, fieldId, getNode, rerender); }
            }
        });
    });
}

function _onRichChange(editor, fieldId, getNode, rerender) {
    const info = getNode(fieldId);
    if (!info || !info.node) return;
    const html = editor.innerHTML;
    const plain = editor.textContent;
    if (info.fieldKey === 'title') { info.node.richTitle = html; info.node.customTitle = plain; }
    else if (info.fieldKey === 'subtitle') { info.node.richSubtitle = html; info.node.customSubtitle = plain; }
    else if (info.fieldKey === 'description') { info.node.richDescription = html; info.node.customDescription = plain; }
    rerender();
}

function _updateRtBtnState(toolbar) {
    try {
        toolbar.querySelector('.rt-bold')?.classList.toggle('active', document.queryCommandState('bold'));
        toolbar.querySelector('.rt-italic')?.classList.toggle('active', document.queryCommandState('italic'));
    } catch (e) { /* ignore */ }
}

/**
 * Sentence Case for Spanish text (Tipo Oración).
 * ONLY transforms text that is ALL CAPS or mostly caps (>70%).
 * If text already has mixed casing (e.g. "ADN del emprendimiento"), it's left as-is.
 * "AVANCE Y DESARROLLO" → "Avance y desarrollo"
 * "MÁS ALLÁ DEL PRODUCTO" → "Más allá del producto"
 * "ADN del emprendimiento" → "ADN del emprendimiento" (unchanged)
 */
function toTitleCaseES(text) {
    if (!text) return text;
    // Count uppercase vs lowercase letters (ignore punctuation, numbers, spaces)
    const letters = text.replace(/[^a-záéíóúñüA-ZÁÉÍÓÚÑÜ]/g, '');
    if (!letters) return text;
    const upperCount = (letters.match(/[A-ZÁÉÍÓÚÑÜ]/g) || []).length;
    const ratio = upperCount / letters.length;
    // Only apply if text is mostly uppercase (>70%) — leave mixed case as-is
    if (ratio <= 0.7) return text;
    const lower = text.toLowerCase();
    // Find first actual letter (skip ¿, ¡, quotes, etc.)
    const match = lower.match(/^([^a-záéíóúñü]*)(.*)/i);
    if (!match) return lower;
    const prefix = match[1];
    const rest = match[2];
    if (!rest) return lower;
    return prefix + rest.charAt(0).toUpperCase() + rest.slice(1);
}

/**
 * Render an arrowhead SVG element at the given position.
 * @param {'triangle'|'open'|'circle'|'diamond'|'none'} type
 * @param {number} tipX - X position of the arrowhead tip
 * @param {number} tipY - Y position of the arrowhead tip
 * @param {number} size - Size of the arrowhead
 * @param {string} color - Fill/stroke color
 * @param {number} strokeW - Stroke width for 'open' type
 * @param {'right'|'down'} dir - Direction the arrow points
 */
function renderArrowhead(type, tipX, tipY, size, color, strokeW, dir) {
    if (type === 'none') return '';
    const s = size;
    const h = s / 2;
    if (dir === 'down') {
        switch (type) {
            case 'open':
                return `<polygon points="${tipX - h},${tipY - s} ${tipX},${tipY} ${tipX + h},${tipY - s}" fill="none" stroke="${color}" stroke-width="${strokeW}" stroke-linejoin="round" />`;
            case 'circle':
                return `<circle cx="${tipX}" cy="${tipY - h}" r="${h}" fill="${color}" />`;
            case 'diamond':
                return `<polygon points="${tipX},${tipY} ${tipX + h},${tipY - h} ${tipX},${tipY - s} ${tipX - h},${tipY - h}" fill="${color}" />`;
            default: // triangle
                return `<polygon points="${tipX - h},${tipY - s} ${tipX},${tipY} ${tipX + h},${tipY - s}" fill="${color}" />`;
        }
    } else { // right
        switch (type) {
            case 'open':
                return `<polygon points="${tipX - s},${tipY - h} ${tipX},${tipY} ${tipX - s},${tipY + h}" fill="none" stroke="${color}" stroke-width="${strokeW}" stroke-linejoin="round" />`;
            case 'circle':
                return `<circle cx="${tipX - h}" cy="${tipY}" r="${h}" fill="${color}" />`;
            case 'diamond':
                return `<polygon points="${tipX},${tipY} ${tipX - h},${tipY - h} ${tipX - s},${tipY} ${tipX - h},${tipY + h}" fill="${color}" />`;
            default: // triangle
                return `<polygon points="${tipX - s},${tipY - h} ${tipX},${tipY} ${tipX - s},${tipY + h}" fill="${color}" />`;
        }
    }
}

/* ========================================================================
   RENDERER (100% Native SVG)
   ======================================================================== */
class SVGRenderer {
    constructor(data, container) {
        this.data = data;
        this.container = container;
        this.svgStr = '';
        this.defsStr = '';
        this.pathsStr = '';
        this.nodesStr = '';

        // Read active settings
        this.font = document.getElementById('s-font-family').value.replace(/['"]/g,'') || "Montserrat, sans-serif";
        this.colors = {
            orange: val('s-color-orange'),
            main: val('s-color-main'),
            week: val('s-color-week'),
            content: val('s-color-content'),
            activity: val('s-color-activity'),
            text: val('s-color-text'),
            conn: val('s-conn-color'),
            border: val('s-border-color')
        };
        
        this.fonts = {
            hito: parseInt(val('s-hito-font')),
            title: parseInt(val('s-title-font')),
            subtitle: parseInt(val('s-subtitle-font')),
            actTitle: parseInt(val('s-act-title-font')),
            actSub: parseInt(val('s-act-sub-font')),
            code: parseInt(val('s-code-font'))
        };
        
        this.dims = {
            mainW: parseInt(val('s-main-width')),
            weekW: parseInt(val('s-week-width')),
            actW: parseInt(val('s-act-width')),
            pad: parseInt(val('s-node-padding')),
            br: parseInt(val('s-border-radius')),
            bw: parseInt(val('s-border-width')),
            cw: parseInt(val('s-conn-width')),
            cr: parseInt(val('s-corner-radius')),
            arSize: parseInt(val('s-arrow-size')),
            arType: val('s-arrow-type')
        };
        
        this.spacing = {
            weekGap: parseInt(val('s-week-gap')),
            actGap: parseInt(val('s-act-gap')),
            weekToAct: parseInt(val('s-week-to-act'))
        };

        // Output dimensions
        this.totalHeight = 0;
        this.totalWidth = parseInt(val('s-canvas-width'));
        this.padding = parseInt(val('s-canvas-padding'));
        
        this.boxes = {}; 
        
        // Setup hidden canvas to measure text
        this.canvas = document.createElement("canvas");
        this.ctx = this.canvas.getContext("2d");
    }

    // Safety multiplier: Canvas measureText uses system fonts which may be
    // narrower than the Google Fonts (Montserrat) used by SVG. Adding 8%
    // ensures text always fits within calculated boxes.
    static MEASURE_SAFETY = 1.08;

    measureText(text, fontSize, fontWeight = 'normal') {
        this.ctx.font = `${fontWeight} ${fontSize}px ${this.font}`;
        return this.ctx.measureText(text).width * SVGRenderer.MEASURE_SAFETY;
    }

    wrapText(text, fontSize, fontWeight, maxWidth) {
        if (!text) return [];
        this.ctx.font = `${fontWeight} ${fontSize}px ${this.font}`;
        const safeMax = maxWidth / SVGRenderer.MEASURE_SAFETY;

        // Pre-process: force line break before "Primera/Segunda/Tercera/Cuarta parte"
        // e.g. "Texto normal. Primera parte" → ["Texto normal.", "Primera parte"]
        const paragraphs = text.toString().split(/\.\s*(?=(Primera|Segunda|Tercera|Cuarta|Quinta|Sexta)\s+parte)/i);
        // The split with lookahead keeps the matched part, but creates empty groups.
        // Clean up: re-join by processing manually
        const segments = [];
        const cleanText = text.toString();
        const partRegex = /\.\s*((Primera|Segunda|Tercera|Cuarta|Quinta|Sexta)\s+parte)/gi;
        let lastIdx = 0;
        let match;
        while ((match = partRegex.exec(cleanText)) !== null) {
            const beforePart = cleanText.substring(lastIdx, match.index + 1).trim(); // include the "."
            if (beforePart) segments.push(beforePart);
            lastIdx = match.index + 1;
        }
        const remainder = cleanText.substring(lastIdx).trim();
        if (remainder) segments.push(remainder);
        if (segments.length === 0) segments.push(cleanText);

        // Word-wrap each segment independently
        let allLines = [];
        segments.forEach(segment => {
            const words = segment.split(' ');
            let currentLine = words[0];
            for (let i = 1; i < words.length; i++) {
                const word = words[i];
                const width = this.ctx.measureText(currentLine + " " + word).width;
                if (width < safeMax) {
                    currentLine += " " + word;
                } else {
                    allLines.push(currentLine);
                    currentLine = word;
                }
            }
            allLines.push(currentLine);
        });
        return allLines;
    }

    drawBox(id, nodeData, x, y, width, bg, titleLines, subLines, tSize, sSize, code = null, codeSize = null, isHito = false) {
        const LH = 1.35;  // universal line-height multiplier
        
        // Support custom overrides
        let actualBg = (nodeData && nodeData.customBgColor) ? nodeData.customBgColor : bg;
        
        // Custom width override
        let actualWidth = (nodeData && nodeData.customWidth) ? nodeData.customWidth : width;
        
        // Custom border color
        const borderColor = (nodeData && nodeData.customBorderColor) ? nodeData.customBorderColor : this.colors.border;
        
        // Per-container font override
        const nodeFont = (nodeData && nodeData.customFont) ? nodeData.customFont.replace(/['"]/g, '') : this.font;
        const nodeTitleSize = (nodeData && nodeData.customTitleSize) || tSize;
        const nodeSubSize = (nodeData && nodeData.customSubtitleSize) || sSize;
        
        const textX = x + this.dims.pad;
        
        // ── Phase 1: Measure total text content height ──
        // We render text starting at y=0, then shift everything to center vertically
        let contentH = 0;
        let textSVG = '';
        let curY = 0; // relative Y (starts at 0, will be offset later)
        
        // Helper: add a <text> element centered vertically in a line slot
        const addText = (tx, fontSize, weight, content, extras = '') => {
            const slotH = fontSize * LH;
            const centerY = curY + slotH / 2;
            textSVG += `<text x="${tx}" y="${centerY}" font-family="${nodeFont}" font-size="${fontSize}" fill="${this.colors.text}" font-weight="${weight}" dominant-baseline="central"${extras}>${italicizeSvgText(content)}</text>`;
            curY += slotH;
            return slotH;
        };
        
        // Check per-field visibility flags
        const showTitle = !(nodeData && nodeData.hiddenTitle);
        const showSub = !(nodeData && nodeData.hiddenSubtitle);
        
        // Check for rich text data
        const hasRichTitle = nodeData && nodeData.richTitle;
        const hasRichSub = nodeData && nodeData.richSubtitle;

        if (isHito) {
            actualWidth = this.measureText(titleLines[0], nodeTitleSize, '800') * 1.05 + (this.dims.pad * 2);
            curY += addText(textX, nodeTitleSize, '800', titleLines[0]);
        } else {
            if (showTitle) {
                if (hasRichTitle) {
                    const segments = parseHtmlToSegments(nodeData.richTitle);
                    const maxW = actualWidth - this.dims.pad * 2;
                    const codeOffset = code ? (this.measureText(code, codeSize || this.fonts.code, '300') + 8) : 0;
                    if (code) {
                        const lineSize = Math.max(codeSize || this.fonts.code, nodeTitleSize);
                        const slotH = lineSize * LH;
                        const centerY = curY + slotH / 2;
                        textSVG += `<text x="${textX}" y="${centerY}" font-family="${nodeFont}" font-size="${codeSize || this.fonts.code}" fill="${this.colors.text}" font-weight="300" dominant-baseline="central" letter-spacing="0.5">${code}</text>`;
                        const result = renderRichSvgText(segments, textX + codeOffset, curY, maxW - codeOffset, { font: nodeFont, size: nodeTitleSize, weight: '700', color: this.colors.text }, this.ctx);
                        textSVG += result.svg;
                        curY += Math.max(slotH, result.height);
                    } else {
                        const result = renderRichSvgText(segments, textX, curY, maxW, { font: nodeFont, size: nodeTitleSize, weight: '700', color: this.colors.text }, this.ctx);
                        textSVG += result.svg;
                        curY += result.height;
                    }
                } else {
                    if (code) {
                        let codeW = this.measureText(code, codeSize, '300') + 8;
                        const lineSize = Math.max(codeSize, nodeTitleSize);
                        const slotH = lineSize * LH;
                        const centerY = curY + slotH / 2;
                        textSVG += `<text x="${textX}" y="${centerY}" font-family="${nodeFont}" font-size="${codeSize}" fill="${this.colors.text}" font-weight="300" dominant-baseline="central" letter-spacing="0.5">${code}</text>`;
                        textSVG += `<text x="${textX + codeW}" y="${centerY}" font-family="${nodeFont}" font-size="${nodeTitleSize}" fill="${this.colors.text}" font-weight="700" dominant-baseline="central">${italicizeSvgText(titleLines[0])}</text>`;
                        curY += slotH;
                        for (let i = 1; i < titleLines.length; i++) {
                            curY += addText(textX + codeW, nodeTitleSize, '700', titleLines[i]);
                        }
                    } else {
                        titleLines.forEach(line => {
                            curY += addText(textX, nodeTitleSize, '700', line);
                        });
                    }
                }
            }

            if (showSub) {
                if (hasRichSub) {
                    if (showTitle) curY += 3;
                    const segments = parseHtmlToSegments(nodeData.richSubtitle);
                    const result = renderRichSvgText(segments, textX, curY, actualWidth - this.dims.pad * 2, { font: nodeFont, size: nodeSubSize, weight: '400', color: this.colors.text }, this.ctx);
                    textSVG += result.svg;
                    curY += result.height;
                } else if (subLines && subLines.length > 0) {
                    if (showTitle) curY += 3;
                    subLines.forEach(line => {
                        curY += addText(textX, nodeSubSize, '400', line);
                    });
                }
            }
        }

        // ── Phase 2: Calculate box height & vertical centering offset ──
        contentH = curY;
        const minBoxH = contentH + this.dims.pad * 2;
        const boxH = Math.max(minBoxH, height || 0);
        
        // Offset to vertically center content within the box
        const textOffsetY = y + (boxH - contentH) / 2;

        // Escape text for the onclick handler
        const safeData = nodeData ? encodeURIComponent(JSON.stringify(nodeData)) : '{}';
        
        // Wrap everything in an interactive group with centering transform
        let gStr = `<g class="interactive-node" style="cursor: pointer;" data-node-id="${id}">`;
        gStr += `<rect x="${x}" y="${y}" width="${actualWidth}" height="${boxH}" rx="${this.dims.br}" fill="${actualBg}" stroke="${borderColor}" stroke-width="${this.dims.bw}" />`;
        gStr += `<g transform="translate(0, ${textOffsetY})">`;
        gStr += textSVG;
        gStr += `</g>`;
        gStr += `</g>`;

        this.nodesStr += gStr;

        this.boxes[id] = { x, y, w: actualWidth, h: boxH, r: x + actualWidth, b: y + boxH, cx: x + actualWidth/2, cy: y + boxH/2 };
        return boxH;
    }



    drawPath(dStr, tipX, tipY) {
        const as = this.dims.arSize;
        const type = this.dims.arType || 'triangle';
        // Draw the line
        this.pathsStr += `<path d="${dStr}" fill="none" stroke="${this.colors.conn}" stroke-width="${this.dims.cw}" stroke-linecap="round" stroke-linejoin="round" />`;
        // Draw arrowhead based on type
        this.pathsStr += renderArrowhead(type, tipX, tipY, as, this.colors.conn, this.dims.cw, 'right');
    }

    render() {
        const root = this.data.diagram.nodes[0];
        const avance = root.children[0];
        // Filter out hidden weeks structurally
        const weeks = avance.children.filter(w => !w.hidden);

        let currentY = this.padding;
        const startX = this.padding;

        // 1. Draw Hito
        const hitoText = root.customTitle || root.text.title;
        const hitoW = this.measureText(hitoText, this.fonts.hito, '800') + (this.dims.pad * 2);
        const hH = this.drawBox(root.id, root, startX, currentY, hitoW, this.colors.orange, [hitoText], [], this.fonts.hito, 0, null, 0, true);
        
        currentY += hH + 25;

        // 2. Draw Avance
        const tx1 = startX + 25;
        const rx = this.dims.cr;
        const avX = tx1 + rx + 15;
        const avTitleStr = avance.customTitle || avance.text.title;
        const avSubStr = avance.customSubtitle || avance.text.subtitle;
        const avTitleLines = this.wrapText(avTitleStr, this.fonts.title, '700', this.dims.mainW - this.dims.pad*2);
        const avSubLines = this.wrapText(avSubStr, this.fonts.subtitle, '400', this.dims.mainW - this.dims.pad*2);
        const avH = this.drawBox(avance.id, avance, avX, currentY, this.dims.mainW, this.colors.main, avTitleLines, avSubLines, this.fonts.title, this.fonts.subtitle);
        
        currentY += avH + 35;

        // 3. Draw Weeks & Activities
        const tx2 = avX + 25;
        const weekX = tx2 + rx + 15;
        const actX = Math.max(weekX + this.dims.weekW + this.spacing.weekToAct, weekX + this.dims.weekW + rx * 2 + 5);
        let calcActW = this.dims.actW;
        if (calcActW <= 0) calcActW = this.totalWidth - actX - this.padding; 

        const isHito1Diagram = this.data.diagram.isHito1;

        weeks.forEach((week, i) => {
            const isFinal = (i === weeks.length - 1);
            // Hito 1: all weeks use the same 'week' color (no orange override for final)
            const wBg = (isFinal && !isHito1Diagram) ? this.colors.orange : this.colors.week;
            
            const wTitleStr = week.customTitle || week.text.title;
            const wSubStr = week.customSubtitle || week.text.subtitle;
            const wTitleLines = this.wrapText(wTitleStr, this.fonts.title, '700', this.dims.weekW - this.dims.pad*2);
            const wSubLines = this.wrapText(wSubStr, this.fonts.subtitle, '400', this.dims.weekW - this.dims.pad*2);
            
            // Filter out hidden activities
            const visibleActs = week.children.filter(act => !act.hidden);
            
            const actsConfig = visibleActs.map(act => {
                const isContent = act.text.title.toLowerCase().startsWith('contenido');
                // Hito 1: all activities use their themed color (no orange override for final)
                // Non-Hito1: final week activities are orange, content nodes use main blue
                let aBg;
                if (isFinal && !isHito1Diagram) {
                    aBg = this.colors.orange;
                } else if (isHito1Diagram) {
                    aBg = isContent ? this.colors.content : this.colors.activity;
                } else {
                    aBg = isContent ? this.colors.main : this.colors.activity;
                }
                
                const aTitleStr = act.customTitle || act.text.title;
                const aSubStr = act.customSubtitle || act.text.subtitle;

                // Code takes some space
                let maxTitleW = calcActW - this.dims.pad*2;
                if(act.text.code) maxTitleW -= (this.measureText(act.text.code, this.fonts.code, '300') + 10);
                
                const aTitleLines = this.wrapText(aTitleStr, this.fonts.actTitle, '700', maxTitleW);
                const aSubLines = this.wrapText(aSubStr, this.fonts.actSub, '400', calcActW - this.dims.pad*2);
                return { aBg, aTitleLines, aSubLines, code: act.text.code, actRef: act };
            });

            // Virtual render acts to get height
            let estimatedSubH = 0;
            actsConfig.forEach(cfg => {
                const LH = 1.35;
                let eh = this.dims.pad * 2; // top + bottom padding
                eh += Math.max(cfg.aTitleLines.length * this.fonts.actTitle * LH, cfg.code ? this.fonts.code * LH : 0);
                if(cfg.aSubLines.length > 0) eh += 3 + cfg.aSubLines.length * this.fonts.actSub * LH;
                estimatedSubH += eh + this.spacing.actGap;
            });
            estimatedSubH -= this.spacing.actGap;

            const wId = week.id;
            const wH = this.drawBox(wId, week, weekX, currentY, this.dims.weekW, wBg, wTitleLines, wSubLines, this.fonts.title, this.fonts.subtitle);
            
            // Re-center Week or Acts
            let actStartY = currentY;
            let finalWeekY = currentY;
            if (estimatedSubH > wH) {
                finalWeekY += (estimatedSubH - wH) / 2;
                this.nodesStr = this.nodesStr.replace(`id="temp-${wId}"`, ''); // hack: we'd need a multi-pass approach, for simplicity let's just draw exactly at top
                // For perfect centering we just redraw the box properties in this.boxes, though SVG text is already baked.
                // It's cleaner in SVG to just top-align or pre-calculate. 
                // We'll proceed with top alignment as it's common in diagrams.
            }

            let maxActB = 0;
            actsConfig.forEach((cfg, j) => {
                const aId = cfg.actRef.id;
                const aH = this.drawBox(aId, cfg.actRef, actX, actStartY, calcActW, cfg.aBg, cfg.aTitleLines, cfg.aSubLines, this.fonts.actTitle, this.fonts.actSub, cfg.code, this.fonts.code);
                maxActB = actStartY + aH;
                actStartY += aH + this.spacing.actGap;
            });

            const branchB = Math.max(currentY + wH, maxActB);
            this.boxes[wId].acts = actsConfig.map((cfg) => this.boxes[cfg.actRef.id]);
            
            currentY = branchB + this.spacing.weekGap;
        });

        this.totalHeight = currentY + this.padding;
        this.addRouting(weeks);

        // No <defs> needed — arrowheads are inline polygons for max compatibility

        const isTransparent = document.getElementById('s-canvas-transparent').checked;
        const bgFill = isTransparent ? 'none' : document.getElementById('s-canvas-bg').value;

        this.svgStr = `<svg xmlns="http://www.w3.org/2000/svg" width="${this.totalWidth}" height="${this.totalHeight}" viewBox="0 0 ${this.totalWidth} ${this.totalHeight}">
            <rect width="100%" height="100%" fill="${bgFill}" />
            ${this.defsStr}
            ${this.pathsStr}
            ${this.nodesStr}
        </svg>`;

        this.container.innerHTML = this.svgStr;

        // Delegated click listener for interactive SVG nodes
        this.container.addEventListener('click', (e) => {
            const g = e.target.closest('[data-node-id]');
            if (g) {
                const nodeId = g.getAttribute('data-node-id');
                window.openNodeEditor(nodeId);
            }
        });
    }

    addRouting(weeks) {
        const as = this.dims.arSize;
        const r = this.dims.cr;

        // L1: Hito -> Avance
        const hId = this.data.diagram.nodes[0].id;
        const aId = this.data.diagram.nodes[0].children[0].id;
        const h = this.boxes[hId];
        const a = this.boxes[aId];
        const tx1 = h.x + 25;
        const cr1 = Math.max(1, Math.min(r, Math.abs(a.cy - h.b) / 2, Math.abs(a.x - tx1) - as - 1));
        this.drawPath(`M${tx1},${h.b} V${a.cy - cr1} A${cr1},${cr1} 0 0 0 ${tx1+cr1},${a.cy} H${a.x - as}`, a.x, a.cy);

        // L2: Avance -> Weeks
        const tx2 = a.x + 25; 
        weeks.forEach((week) => {
            const w = this.boxes[week.id];
            const cr2 = Math.max(1, Math.min(r, Math.abs(w.cy - a.b) / 2, Math.abs(w.x - tx2) - as - 1));
            this.drawPath(`M${tx2},${a.b} V${w.cy - cr2} A${cr2},${cr2} 0 0 0 ${tx2+cr2},${w.cy} H${w.x - as}`, w.x, w.cy);
        });

        // L3: Week -> Activities
        weeks.forEach((week) => {
            const w = this.boxes[week.id];
            const acts = w.acts;
            if(!acts || acts.length === 0) return;

            const braceWidth = acts[0].x - w.r;
            const braceX = w.r + braceWidth / 2;

            acts.forEach(act => {
                const dy = act.cy - w.cy;
                if (Math.abs(dy) < 2) {
                    this.drawPath(`M${w.r},${w.cy} H${act.x - as}`, act.x, act.cy);
                } else {
                    const cr3 = Math.max(1, Math.min(r, Math.abs(dy) / 2, braceWidth / 2 - 1));
                    if (dy > 0) {
                        this.drawPath(`M${w.r},${w.cy} H${braceX - cr3} A${cr3},${cr3} 0 0 1 ${braceX},${w.cy + cr3} V${act.cy - cr3} A${cr3},${cr3} 0 0 0 ${braceX+cr3},${act.cy} H${act.x - as}`, act.x, act.cy);
                    } else {
                        this.drawPath(`M${w.r},${w.cy} H${braceX - cr3} A${cr3},${cr3} 0 0 0 ${braceX},${w.cy - cr3} V${act.cy + cr3} A${cr3},${cr3} 0 0 1 ${braceX+cr3},${act.cy} H${act.x - as}`, act.x, act.cy);
                    }
                }
            });
        });
    }
}

/* ========================================================================
   RENDERER — Página de Inicio (Home Page / Course Overview)
   ======================================================================== */
class PaginaInicioRenderer {
    constructor(data, container) {
        this.data = data;
        this.container = container;

        // Read active settings (reuse same inputs)
        this.font = document.getElementById('s-font-family').value.replace(/['"]/g,'') || "Montserrat, sans-serif";
        this.colors = {
            orange: val('s-color-orange'),
            main: val('s-color-main'),
            text: val('s-color-text'),
            conn: val('s-conn-color'),
            border: val('s-border-color'),
            subBox: val('s-diag-sub-color'),
            subBoxBg: val('s-diag-sub-bg') || 'none'
        };
        this.fonts = {
            hito: parseInt(val('s-hito-font')),
            title: parseInt(val('s-title-font')),
            subtitle: parseInt(val('s-subtitle-font'))
        };
        this.dims = {
            pad: parseInt(val('s-node-padding')),
            br: parseInt(val('s-border-radius')),
            bw: parseInt(val('s-border-width')),
            cw: parseInt(val('s-conn-width')),
            cr: parseInt(val('s-corner-radius')),
            arSize: parseInt(val('s-arrow-size')),
            arType: val('s-arrow-type')
        };
        this.totalWidth = parseInt(val('s-canvas-width'));
        this.padding = parseInt(val('s-canvas-padding'));

        // Layout dimensions for Página de Inicio (from settings)
        this.hitoBoxW = parseInt(val('s-diag-hito-w')) || 340;
        this.subtitleBoxW = parseInt(val('s-diag-sub-w')) || 300;
        this.rowGap = parseInt(val('s-diag-row-gap')) || 25;

        this.canvas = document.createElement('canvas');
        this.ctx = this.canvas.getContext('2d');
    }

    static MEASURE_SAFETY = 1.08;

    measureText(text, fontSize, fontWeight = 'normal') {
        this.ctx.font = `${fontWeight} ${fontSize}px ${this.font}`;
        return this.ctx.measureText(text).width * PaginaInicioRenderer.MEASURE_SAFETY;
    }

    wrapText(text, fontSize, fontWeight, maxWidth) {
        if (!text) return [];
        this.ctx.font = `${fontWeight} ${fontSize}px ${this.font}`;
        const safeMax = maxWidth / PaginaInicioRenderer.MEASURE_SAFETY;
        const words = text.toString().split(' ');
        let lines = [];
        let currentLine = words[0] || '';
        for (let i = 1; i < words.length; i++) {
            const test = currentLine + ' ' + words[i];
            if (this.ctx.measureText(test).width < safeMax) {
                currentLine = test;
            } else {
                lines.push(currentLine);
                currentLine = words[i];
            }
        }
        lines.push(currentLine);
        return lines;
    }

    render() {
        const d = this.data.diagram;
        const LH = 1.35;
        let svgContent = '';
        let pathsStr = '';
        let nodesStr = '';

        const startX = this.padding;
        let currentY = this.padding;

        // ── 1. Draw course title header (orange box) ──
        const headerTitle = d.customTitle || d.title;
        const headerFontSize = this.fonts.hito;
        const headerW = this.measureText(headerTitle, headerFontSize, '800') * 1.05 + this.dims.pad * 2;
        const headerH = headerFontSize * LH + this.dims.pad * 2;
        const headerCenterY = currentY + headerH / 2;

        nodesStr += `<g class="interactive-node" style="cursor:pointer" data-node-id="course_title">`;
        nodesStr += `<rect x="${startX}" y="${currentY}" width="${headerW}" height="${headerH}" rx="${this.dims.br}" fill="${this.colors.orange}" />`;
        nodesStr += `<text x="${startX + this.dims.pad}" y="${headerCenterY}" font-family="${this.font}" font-size="${headerFontSize}" fill="${this.colors.text}" font-weight="800" dominant-baseline="central">${headerTitle}</text>`;
        nodesStr += `</g>`;

        const headerBox = { x: startX, y: currentY, w: headerW, h: headerH, b: currentY + headerH };
        currentY += headerH + 35;

        // Trunk X position (vertical line comes down from header)
        const trunkX = startX + 25;
        const r = this.dims.cr;

        // Hito box X (indented from trunk)
        const hitoX = trunkX + r + 15;

        // Auto-fit: measure all Hito titles/subtitles to determine maximum box width
        const minHitoW = this.hitoBoxW; // user setting as minimum
        let autoHitoW = minHitoW;
        d.nodes.forEach(node => {
            const hitoNum = node.customTitle !== undefined ? node.customTitle : `Hito ${node.text.hitoNum}`;
            const hitoType = node.customSubtitle !== undefined ? node.customSubtitle : node.text.hitoType;
            const titleW = this.measureText(hitoNum, this.fonts.title, '800') + (this.dims.pad * 2);
            if (titleW > autoHitoW) autoHitoW = titleW;
            if (hitoType) {
                const typeW = this.measureText(hitoType, this.fonts.subtitle, '400') + (this.dims.pad * 2);
                if (typeW > autoHitoW) autoHitoW = typeW;
            }
        });
        this.hitoBoxW = Math.ceil(autoHitoW);

        // Subtitle box X (after hito box + gap)
        const subGap = 75;
        const subX = hitoX + this.hitoBoxW + subGap;
        // Auto-calc subtitle box width to fill remaining canvas
        this.subtitleBoxW = Math.max(200, this.totalWidth - subX - this.padding);

        // Store boxes for routing
        const hitoBoxes = [];
        const subBoxes = [];

        // ── 2. Draw each Hito row ──
        d.nodes.forEach((node, idx) => {
            const hitoNum = node.customTitle !== undefined ? node.customTitle : `Hito ${node.text.hitoNum}`;
            const hitoType = node.customSubtitle !== undefined ? node.customSubtitle : node.text.hitoType;
            const subtitle = node.customDescription !== undefined ? node.customDescription : node.text.subtitle;

            const showTitle = !node.hiddenTitle;
            const showSub = !node.hiddenSubtitle;
            const showDesc = !node.hiddenDescription;

            const bgColor = node.customBgColor || this.colors.main;
            const borderColor = node.customBorderColor || this.colors.border;
            const nodeFont = node.customFont ? node.customFont.replace(/['"]/g, '') : this.font;
            const nodeTitleSize = node.customTitleSize || this.fonts.title;
            const nodeSubSize = node.customSubtitleSize || this.fonts.subtitle;
            const nodeDescSize = node.customDescriptionSize || this.fonts.subtitle;

            // ── Hito blue box ──
            let hitoTextSvg = '';
            let hitoCurY = currentY + this.dims.pad;

            if (showTitle) {
                if (node.richTitle) {
                    const segments = parseHtmlToSegments(node.richTitle);
                    const result = renderRichSvgText(segments, hitoX + this.dims.pad, hitoCurY, this.hitoBoxW - this.dims.pad * 2, { font: nodeFont, size: nodeTitleSize, weight: '800', color: this.colors.text }, this.ctx);
                    hitoTextSvg += result.svg;
                    hitoCurY += result.height;
                } else {
                    const titleSlot = nodeTitleSize * LH;
                    const titleCY = hitoCurY + titleSlot / 2;
                    hitoTextSvg += `<text x="${hitoX + this.dims.pad}" y="${titleCY}" font-family="${nodeFont}" font-size="${nodeTitleSize}" fill="${this.colors.text}" font-weight="800" dominant-baseline="central">${hitoNum}</text>`;
                    hitoCurY += titleSlot;
                }
            }

            if (showSub && hitoType) {
                if (node.richSubtitle) {
                    const segments = parseHtmlToSegments(node.richSubtitle);
                    const result = renderRichSvgText(segments, hitoX + this.dims.pad, hitoCurY, this.hitoBoxW - this.dims.pad * 2, { font: nodeFont, size: nodeSubSize, weight: '400', color: this.colors.text }, this.ctx);
                    hitoTextSvg += result.svg;
                    hitoCurY += result.height;
                } else {
                    const typeLines = this.wrapText(hitoType, nodeSubSize, '400', this.hitoBoxW - this.dims.pad * 2);
                    typeLines.forEach(line => {
                        const slot = nodeSubSize * LH;
                        const cy = hitoCurY + slot / 2;
                        hitoTextSvg += `<text x="${hitoX + this.dims.pad}" y="${cy}" font-family="${nodeFont}" font-size="${nodeSubSize}" fill="${this.colors.text}" font-weight="400" dominant-baseline="central">${italicizeSvgText(line)}</text>`;
                        hitoCurY += slot;
                    });
                }
            }

            const hitoH = hitoCurY - currentY + this.dims.pad;

            // ── Subtitle dashed box ──
            let subTextSvg = '';
            let subCurY = currentY + this.dims.pad;

            if (showDesc && subtitle) {
                if (node.richDescription) {
                    const segments = parseHtmlToSegments(node.richDescription);
                    const result = renderRichSvgText(segments, subX + this.dims.pad, subCurY, this.subtitleBoxW - this.dims.pad * 2, { font: nodeFont, size: nodeDescSize, weight: '400', color: '#333' }, this.ctx);
                    subTextSvg += result.svg;
                    subCurY += result.height;
                } else {
                    const subLines = this.wrapText(subtitle, nodeDescSize, '400', this.subtitleBoxW - this.dims.pad * 2);
                    subLines.forEach(line => {
                        const slot = nodeDescSize * LH;
                        const cy = subCurY + slot / 2;
                        subTextSvg += `<text x="${subX + this.dims.pad}" y="${cy}" font-family="${nodeFont}" font-size="${nodeDescSize}" fill="#333" font-weight="400" dominant-baseline="central">${italicizeSvgText(line)}</text>`;
                        subCurY += slot;
                    });
                }
            }

            const subH = Math.max(hitoH, subCurY - currentY + this.dims.pad);
            const rowH = Math.max(hitoH, subH);

            // Center smaller box vertically
            const hitoYOffset = currentY + (rowH - hitoH) / 2;
            const subYOffset = currentY + (rowH - subH) / 2;

            // Draw hito blue box
            nodesStr += `<g class="interactive-node" style="cursor:pointer" data-node-id="${node.id}">`;
            nodesStr += `<rect x="${hitoX}" y="${hitoYOffset}" width="${this.hitoBoxW}" height="${hitoH}" rx="${this.dims.br}" fill="${bgColor}" stroke="${borderColor}" stroke-width="${this.dims.bw}" />`;
            // Re-render text at correct offset
            nodesStr += hitoTextSvg.replace(new RegExp(`y="${currentY + this.dims.pad}`, 'g'), `y="${hitoYOffset + this.dims.pad}`);
            nodesStr += `</g>`;

            // Draw subtitle dashed box
            nodesStr += `<g class="interactive-node" style="cursor:pointer" data-node-id="${node.id}_sub">`;
            nodesStr += `<rect x="${subX}" y="${subYOffset}" width="${this.subtitleBoxW}" height="${subH}" rx="${this.dims.br}" fill="${this.colors.subBoxBg}" stroke="${this.colors.subBox}" stroke-width="${this.dims.bw}" />`;
            nodesStr += subTextSvg;
            nodesStr += `</g>`;


            hitoBoxes.push({ x: hitoX, y: hitoYOffset, w: this.hitoBoxW, h: hitoH, cx: hitoX + this.hitoBoxW / 2, cy: hitoYOffset + hitoH / 2, r: hitoX + this.hitoBoxW, b: hitoYOffset + hitoH });
            subBoxes.push({ x: subX, y: subYOffset, w: this.subtitleBoxW, h: subH, cx: subX + this.subtitleBoxW / 2, cy: subYOffset + subH / 2 });

            currentY += rowH + this.rowGap;
        });

        // ── 3. Draw connectors ──
        const as = this.dims.arSize;

        // Vertical trunk from header down
        if (hitoBoxes.length > 0) {
            const firstHito = hitoBoxes[0];
            const lastHito = hitoBoxes[hitoBoxes.length - 1];

            // Header → first hito (L-shape)
            const cr1 = Math.max(1, Math.min(r, Math.abs(firstHito.cy - headerBox.b) / 2, Math.abs(firstHito.x - trunkX) - as - 1));
            pathsStr += `<path d="M${trunkX},${headerBox.b} V${firstHito.cy - cr1} A${cr1},${cr1} 0 0 0 ${trunkX + cr1},${firstHito.cy} H${firstHito.x - as}" fill="none" stroke="${this.colors.conn}" stroke-width="${this.dims.cw}" stroke-linecap="round" />`;
            pathsStr += renderArrowhead(this.dims.arType || 'triangle', firstHito.x, firstHito.cy, as, this.colors.conn, this.dims.cw, 'right');

            // Subsequent hitos: trunk continues down
            for (let i = 1; i < hitoBoxes.length; i++) {
                const hb = hitoBoxes[i];
                const prevB = hitoBoxes[i - 1].b + this.rowGap * 0.1;
                const crN = Math.max(1, Math.min(r, Math.abs(hb.cy - headerBox.b) / 2, Math.abs(hb.x - trunkX) - as - 1));
                pathsStr += `<path d="M${trunkX},${headerBox.b} V${hb.cy - crN} A${crN},${crN} 0 0 0 ${trunkX + crN},${hb.cy} H${hb.x - as}" fill="none" stroke="${this.colors.conn}" stroke-width="${this.dims.cw}" stroke-linecap="round" />`;
                pathsStr += renderArrowhead(this.dims.arType || 'triangle', hb.x, hb.cy, as, this.colors.conn, this.dims.cw, 'right');
            }

            // Hito → Subtitle (horizontal arrow)
            for (let i = 0; i < hitoBoxes.length; i++) {
                const hb = hitoBoxes[i];
                const sb = subBoxes[i];
                pathsStr += `<path d="M${hb.r},${hb.cy} H${sb.x - as}" fill="none" stroke="${this.colors.conn}" stroke-width="${this.dims.cw}" stroke-linecap="round" />`;
                pathsStr += renderArrowhead(this.dims.arType || 'triangle', sb.x, hb.cy, as, this.colors.conn, this.dims.cw, 'right');
            }
        }

        const totalH = currentY + this.padding;
        const isTransparent = document.getElementById('s-canvas-transparent').checked;
        const bgFill = isTransparent ? 'none' : document.getElementById('s-canvas-bg').value;

        this.container.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="${this.totalWidth}" height="${totalH}" viewBox="0 0 ${this.totalWidth} ${totalH}">
            <rect width="100%" height="100%" fill="${bgFill}" />
            ${pathsStr}
            ${nodesStr}
        </svg>`;

        // Click listener
        this.container.addEventListener('click', (e) => {
            const g = e.target.closest('[data-node-id]');
            if (g) {
                const nodeId = g.getAttribute('data-node-id');
                window.openNodeEditor(nodeId);
            }
        });
    }
}

/* ========================================================================
   RENDERER — Tarjeta de Hito (Card SVG) — Fixed 376×872
   ======================================================================== */
class TarjetaRenderer {
    constructor(nodeData, container) {
        this.node = nodeData;
        this.container = container;
        this.font = "'Suprema', Arial, sans-serif";

        // Fixed canvas size
        this.W = 376;
        this.H = 872;

        // Shape zones (from settings)
        this.headerH = parseInt(val('s-card-header-h')) || 92;
        this.shadowH = parseInt(val('s-card-shadow-h')) || 12;
        this.footerH = parseInt(val('s-card-footer-h')) || 194;
        this.cornerR = parseInt(val('s-card-corner-r')) || 35;

        // Body fills remaining space
        this.bodyH = this.H - this.headerH - this.shadowH - this.footerH;

        // Colors — read from settings panel
        this.colors = {
            orange: val('s-card-header') || '#F16522',
            orangeDark: this.darkenColor(val('s-card-header') || '#F16522', 30),
            blue: val('s-card-footer') || '#516BED',
            blueLt: this.lightenColor(val('s-card-footer') || '#516BED', 40),
            gray: val('s-card-body') || '#DDD',
            white: '#FFF',
            black: '#000'
        };

        this.canvas = document.createElement('canvas');
        this.ctx = this.canvas.getContext('2d');
    }

    wrapText(text, fontSize, maxWidth) {
        if (!text) return [];
        this.ctx.font = `${fontSize}px Arial`;
        // Safety factor: canvas measures with Arial but SVG renders with Suprema (wider)
        const safeMax = maxWidth * 0.85;
        const words = text.toString().split(' ');
        let lines = [];
        let current = words[0] || '';
        for (let i = 1; i < words.length; i++) {
            const test = current + ' ' + words[i];
            if (this.ctx.measureText(test).width < safeMax) {
                current = test;
            } else {
                lines.push(current);
                current = words[i];
            }
        }
        lines.push(current);
        return lines;
    }

    render() {
        const n = this.node;
        const hitoTitle = n.card.customTitle !== undefined ? n.card.customTitle : `Hito ${n.text.hitoNum}`;
        const bodyText = n.card.customBody !== undefined ? n.card.customBody : (n.card.body || '');
        const questionText = n.card.customQuestion !== undefined ? n.card.customQuestion : (n.card.question || '');

        const W = this.W;
        const H = this.H;
        const R = this.cornerR;
        const shInset = this.shadowH + 2; // inset for body rect and shadows

        const headerBottom = this.headerH;            // bottom of orange header shape
        const bodyTop = headerBottom + this.shadowH;   // top of gray body
        const bodyH = this.bodyH;
        const footerY = bodyTop + bodyH;               // top of blue footer
        const footerH = this.footerH;

        // ── Read font sizes from settings ──
        const textPadX = 28;
        const bodyMaxW = W - textPadX * 2;
        const bodyAlign = val('s-card-body-align') || 'left';
        const footerAlign = val('s-card-footer-align') || 'center';
        let bodyFontSize = parseInt(val('s-card-body-font')) || 20;
        let questionFontSize = parseInt(val('s-card-q-font')) || 22;

        let svg = '';

        // ── Shapes (fixed positions) ──

        // Orange header (rounded top)
        svg += `<path d="M${R},0 L${W - R},0 Q${W},0 ${W},${R} L${W},${headerBottom} L0,${headerBottom} L0,${R} Q0,0 ${R},0 Z" fill="${this.colors.orange}"/>`;

        // Header shadow strip
        svg += `<polygon points="0,${headerBottom} ${shInset},${bodyTop} ${W - shInset},${bodyTop} ${W},${headerBottom}" fill="${this.colors.orangeDark}"/>`;

        // Gray body rect
        svg += `<rect x="${shInset}" y="${bodyTop}" width="${W - shInset * 2}" height="${bodyH}" fill="${this.colors.gray}"/>`;

        // Footer shadow strip
        svg += `<polygon points="${W},${footerY} ${W - shInset},${footerY - this.shadowH} ${shInset},${footerY - this.shadowH} 0,${footerY}" fill="${this.colors.blueLt}"/>`;

        // Blue footer (rounded bottom)
        svg += `<path d="M0,${footerY} L${W},${footerY} L${W},${H - R} Q${W},${H} ${W - R},${H} L${R},${H} Q0,${H} 0,${H - R} Z" fill="${this.colors.blue}"/>`;

        // ── Header text: "Hito N" ──
        const titleFontSize = parseInt(val('s-card-title-font')) || 60;
        svg += `<text x="${W / 2}" y="${headerBottom / 2}" font-family="${this.font}" font-size="${titleFontSize}" fill="${this.colors.white}" font-weight="700" text-anchor="middle" dominant-baseline="central">${this.escapeXml(hitoTitle)}</text>`;

        // ── Body text (pure SVG text for export compatibility) ──
        const bodyLH = parseFloat(val('s-card-body-lh')) || 1.35;
        const bodyPadTop = 18;
        const bodyLeftX = textPadX;
        const bodyRightX = W - textPadX;
        const bodyLineH = bodyFontSize * bodyLH;
        const bodyLines = this.wrapText(bodyText, bodyFontSize, bodyMaxW);
        const bodyStartY = bodyTop + bodyPadTop + bodyFontSize;
        svg += this.renderAlignedLines(bodyLines, bodyAlign, bodyLeftX, bodyRightX, bodyStartY, bodyLineH, bodyFontSize, this.colors.black, '400');

        // ── Footer text (pure SVG text for export compatibility) ──
        const footerLH = parseFloat(val('s-card-footer-lh')) || 1.3;
        const footerPadX = textPadX;
        const footerLeftX = footerPadX;
        const footerRightX = W - footerPadX;
        const footerMaxW = W - footerPadX * 2;
        const footerLineH = questionFontSize * footerLH;
        const footerLines = this.wrapText(questionText, questionFontSize, footerMaxW);
        // Vertically center footer text within the footer area
        const footerTextBlockH = footerLines.length * footerLineH;
        const footerStartY = footerY + (footerH - footerTextBlockH) / 2 + questionFontSize * 0.8;
        svg += this.renderAlignedLines(footerLines, footerAlign, footerLeftX, footerRightX, footerStartY, footerLineH, questionFontSize, this.colors.white, '600');

        const fullSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="${this.escapeXml(hitoTitle)}">
            ${svg}
        </svg>`;

        this.container.innerHTML = fullSvg;
    }

    escapeXml(s) {
        return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    /**
     * Render lines with alignment: left, center, right, justify
     * For justify: positions each word individually with calculated spacing
     */
    renderAlignedLines(lines, align, leftX, rightX, startY, lineH, fontSize, fill, weight) {
        let out = '';
        const totalW = rightX - leftX;
        const centerX = (leftX + rightX) / 2;

        lines.forEach((line, idx) => {
            const y = startY + idx * lineH;
            const escaped = this.escapeXml(line);

            if (align === 'justify' && idx < lines.length - 1) {
                // Justify: distribute words across the full width (except last line)
                const words = line.split(' ');
                if (words.length <= 1) {
                    out += `<text x="${leftX}" y="${y}" font-family="${this.font}" font-size="${fontSize}" fill="${fill}" font-weight="${weight}">${escaped}</text>`;
                } else {
                    // Measure each word width
                    this.ctx.font = `${weight === '700' || weight === '600' ? 'bold' : 'normal'} ${fontSize}px Arial`;
                    const wordWidths = words.map(w => this.ctx.measureText(w).width * 1.15);
                    const totalWordsW = wordWidths.reduce((a, b) => a + b, 0);
                    const totalGap = totalW - totalWordsW;
                    const gapPerSpace = totalGap / (words.length - 1);

                    let wx = leftX;
                    words.forEach((word, wi) => {
                        const italicAttr = isEnglishWord(word) ? ' font-style="italic"' : '';
                        out += `<text x="${wx.toFixed(1)}" y="${y}" font-family="${this.font}" font-size="${fontSize}" fill="${fill}" font-weight="${weight}"${italicAttr}>${this.escapeXml(word)}</text>`;
                        wx += wordWidths[wi] + gapPerSpace;
                    });
                }
            } else {
                // left, center, right (or last line of justify → left)
                let x, anchor;
                if (align === 'center') { x = centerX; anchor = 'middle'; }
                else if (align === 'right') { x = rightX; anchor = 'end'; }
                else { x = leftX; anchor = 'start'; } // left or justify-last-line
                out += `<text x="${x}" y="${y}" font-family="${this.font}" font-size="${fontSize}" fill="${fill}" font-weight="${weight}" text-anchor="${anchor}">${italicizeSvgText(escaped)}</text>`;
            }
        });
        return out;
    }

    darkenColor(hex, amount) {
        hex = hex.replace('#', '');
        let r = Math.max(0, parseInt(hex.substring(0, 2), 16) - amount);
        let g = Math.max(0, parseInt(hex.substring(2, 4), 16) - amount);
        let b = Math.max(0, parseInt(hex.substring(4, 6), 16) - amount);
        return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
    }

    lightenColor(hex, amount) {
        hex = hex.replace('#', '');
        let r = Math.min(255, parseInt(hex.substring(0, 2), 16) + amount);
        let g = Math.min(255, parseInt(hex.substring(2, 4), 16) + amount);
        let b = Math.min(255, parseInt(hex.substring(4, 6), 16) + amount);
        return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
    }
}

/* ========================================================================
   TAB NAVIGATION (Diagrama + Tarjetas)
   ======================================================================== */
let currentActiveTab = 'diagram'; // 'diagram' | 'tarjeta_0' ... 'tarjeta_4'

function renderTabs() {
    const tabBar = document.getElementById('canvas-tabs');
    if (!tabBar) return;

    // Only show tabs for Página de Inicio documents
    if (!currentDiagramData || !currentDiagramData.diagram || currentDiagramData.diagram.type !== 'paginaInicio') {
        tabBar.innerHTML = '';
        tabBar.style.display = 'none';
        currentActiveTab = 'diagram';
        return;
    }

    tabBar.style.display = 'flex';
    const nodes = currentDiagramData.diagram.nodes;

    let html = `<button class="tab-btn ${currentActiveTab === 'diagram' ? 'active' : ''}" data-tab="diagram">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
        Diagrama
    </button>`;

    nodes.forEach((node, idx) => {
        const tabId = `tarjeta_${idx}`;
        // Check if card text needs review (empty body or question)
        const body = (node.card?.customBody !== undefined ? node.card.customBody : node.card?.body) || '';
        const question = (node.card?.customQuestion !== undefined ? node.card.customQuestion : node.card?.question) || '';
        const needsReview = !body.trim() || !question.trim();
        const reviewDot = needsReview
            ? `<span class="tab-review-dot warn" title="Texto incompleto — revisar">!</span>`
            : `<span class="tab-review-dot ok" title="Texto completo">✓</span>`;
        html += `<button class="tab-btn ${currentActiveTab === tabId ? 'active' : ''}" data-tab="${tabId}">
            Hito ${node.text.hitoNum} ${reviewDot}
        </button>`;
    });

    tabBar.innerHTML = html;

    // Bind click
    tabBar.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const tab = btn.getAttribute('data-tab');
            switchTab(tab);
        });
    });
}

function switchTab(tab) {
    currentActiveTab = tab;
    renderTabs(); // Update active state

    const container = document.getElementById('diagram');
    container.innerHTML = '';

    if (tab === 'diagram') {
        renderDiagram(currentDiagramData, true);
    } else {
        // Render tarjeta
        const idx = parseInt(tab.replace('tarjeta_', ''));
        const node = currentDiagramData.diagram.nodes[idx];
        if (node) {
            const renderer = new TarjetaRenderer(node, container);
            renderer.render();
        }
        renderTarjetaSidebar(idx);
    }
}

function renderTarjetaSidebar(hitoIdx) {
    const list = document.getElementById('structure-content-list');
    if (!list || !currentDiagramData) return;

    const node = currentDiagramData.diagram.nodes[hitoIdx];
    if (!node) return;

    const hitoTitle = (node.card.customTitle !== undefined ? node.card.customTitle : `Hito ${node.text.hitoNum}`).replace(/"/g, '&quot;');
    const bodyText = (node.card.customBody !== undefined ? node.card.customBody : node.card.body).replace(/"/g, '&quot;').replace(/</g, '&lt;');
    const questionText = (node.card.customQuestion !== undefined ? node.card.customQuestion : node.card.question).replace(/"/g, '&quot;').replace(/</g, '&lt;');

    let html = `
    <div class="structure-group">
        <div class="structure-card" data-id="${node.id}">
            <div class="structure-group-title" style="margin-bottom:0.3rem;">Tarjeta — Hito ${node.text.hitoNum}</div>
            <div class="card-collapsible-body">
                <div class="field-row">
                    <label style="font-size:0.7rem; color:var(--edtech-text-muted); font-weight:600; margin-bottom:0.1rem;">Título (Header)</label>
                </div>
                <div class="field-row">
                    <input type="text" class="live-edit-card-title" data-idx="${hitoIdx}" value="${hitoTitle}" placeholder="Hito N...">
                </div>
                <div class="field-row" style="margin-top:0.4rem;">
                    <label style="font-size:0.7rem; color:var(--edtech-text-muted); font-weight:600; margin-bottom:0.1rem;">Texto de competencia (Body)</label>
                </div>
                <div class="field-row">
                    <textarea class="live-edit-card-body" rows="4" data-idx="${hitoIdx}" placeholder="Texto de competencia..." spellcheck="true" lang="es">${bodyText}</textarea>
                </div>
                <div class="field-row" style="margin-top:0.4rem;">
                    <label style="font-size:0.7rem; color:var(--edtech-text-muted); font-weight:600; margin-bottom:0.1rem;">Pregunta Autoevalúate (Footer)</label>
                </div>
                <div class="field-row">
                    <textarea class="live-edit-card-question" rows="3" data-idx="${hitoIdx}" placeholder="Pregunta retadora..." spellcheck="true" lang="es">${questionText}</textarea>
                </div>
            </div>
        </div>
    </div>`;

    list.innerHTML = html;

    // Bind events
    const titleInput = list.querySelector('.live-edit-card-title');
    if (titleInput) {
        titleInput.addEventListener('input', (e) => {
            node.card.customTitle = e.target.value;
            renderActiveCard(hitoIdx);
        });
    }

    const bodyTextarea = list.querySelector('.live-edit-card-body');
    if (bodyTextarea) {
        setTimeout(() => { bodyTextarea.style.height = 'auto'; bodyTextarea.style.height = bodyTextarea.scrollHeight + 'px'; }, 10);
        bodyTextarea.addEventListener('input', (e) => {
            e.target.style.height = 'auto';
            e.target.style.height = e.target.scrollHeight + 'px';
            node.card.customBody = e.target.value;
            renderActiveCard(hitoIdx);
        });
    }

    const qTextarea = list.querySelector('.live-edit-card-question');
    if (qTextarea) {
        setTimeout(() => { qTextarea.style.height = 'auto'; qTextarea.style.height = qTextarea.scrollHeight + 'px'; }, 10);
        qTextarea.addEventListener('input', (e) => {
            e.target.style.height = 'auto';
            e.target.style.height = e.target.scrollHeight + 'px';
            node.card.customQuestion = e.target.value;
            renderActiveCard(hitoIdx);
        });
    }
}

function renderActiveCard(hitoIdx) {
    const container = document.getElementById('diagram');
    const node = currentDiagramData.diagram.nodes[hitoIdx];
    if (node && container) {
        const renderer = new TarjetaRenderer(node, container);
        renderer.render();
    }
}

function renderDiagram(data, updateSidebar = true) {
    const container = document.getElementById('diagram');
    container.innerHTML = '';
    if (!data) return;

    if (data.diagram && data.diagram.type === 'paginaInicio') {
        // Render tabs for Página de Inicio
        renderTabs();

        if (currentActiveTab === 'diagram') {
            const renderer = new PaginaInicioRenderer(data, container);
            renderer.render();
        } else {
            const idx = parseInt(currentActiveTab.replace('tarjeta_', ''));
            const node = data.diagram.nodes[idx];
            if (node) {
                const renderer = new TarjetaRenderer(node, container);
                renderer.render();
            }
            if (updateSidebar) {
                renderTarjetaSidebar(idx);
                return; // Don't run renderStructurePanel
            }
        }
    } else {
        // Hide tabs for regular Hito documents
        const tabBar = document.getElementById('canvas-tabs');
        if (tabBar) { tabBar.innerHTML = ''; tabBar.style.display = 'none'; }

        const renderer = new SVGRenderer(data, container);
        renderer.render();
    }

    if (updateSidebar) renderStructurePanel();
}

/* ========================================================================
   SETTINGS PANEL
   ======================================================================== */

// Toggle panel visibility
document.getElementById('toggle-panel').addEventListener('click', () => {
    const panel = document.getElementById('settings-panel');
    const btn = document.getElementById('toggle-panel');
    panel.classList.toggle('hidden');
    btn.classList.toggle('active');
});

// Apply settings → update CSS vars and re-render
document.getElementById('apply-settings').addEventListener('click', applySettings);

// Real-time: listen for any input change inside settings panel
document.querySelectorAll('#settings-panel input, #settings-panel select').forEach(el => {
    el.addEventListener('input', applySettings);
    el.addEventListener('change', applySettings);
});

function applySettings() {
    const root = document.documentElement;
    const dc = document.getElementById('diagram');

    // Lienzo
    dc.style.width = val('s-canvas-width') + 'px';
    const transpEl = document.getElementById('s-canvas-transparent');
    const isTransparent = transpEl ? transpEl.checked : true;
    dc.style.background = isTransparent ? 'transparent' : val('s-canvas-bg');

    // Tipografía
    root.style.setProperty('--font', val('s-font-family'));
    root.style.setProperty('--hito-font', val('s-hito-font') + 'px');
    root.style.setProperty('--title-font', val('s-title-font') + 'px');
    root.style.setProperty('--subtitle-font', val('s-subtitle-font') + 'px');
    root.style.setProperty('--act-title-font', val('s-act-title-font') + 'px');
    root.style.setProperty('--act-sub-font', val('s-act-sub-font') + 'px');
    root.style.setProperty('--code-font', val('s-code-font') + 'px');

    // Colores
    root.style.setProperty('--orange', val('s-color-orange'));
    root.style.setProperty('--main-blue', val('s-color-main'));
    root.style.setProperty('--week-blue', val('s-color-week'));
    root.style.setProperty('--content-blue', val('s-color-content'));
    root.style.setProperty('--activity-blue', val('s-color-activity'));
    root.style.setProperty('--text-color', val('s-color-text'));

    // Conectores
    root.style.setProperty('--cw', val('s-conn-width'));
    root.style.setProperty('--conn', val('s-conn-color'));
    root.style.setProperty('--corner-radius', val('s-corner-radius'));
    root.style.setProperty('--arrow-size', val('s-arrow-size'));

    // Bordes
    root.style.setProperty('--border-radius', val('s-border-radius') + 'px');
    root.style.setProperty('--border-width', val('s-border-width') + 'px');
    root.style.setProperty('--border-color', val('s-border-color'));

    // Dimensiones
    root.style.setProperty('--main-width', val('s-main-width') + 'px');
    root.style.setProperty('--week-width', val('s-week-width') + 'px');
    const aw = parseInt(val('s-act-width'));
    root.style.setProperty('--act-width', aw > 0 ? aw + 'px' : '0');
    root.style.setProperty('--node-padding', val('s-node-padding') + 'px');

    // Espaciado
    root.style.setProperty('--week-gap', val('s-week-gap') + 'px');
    root.style.setProperty('--act-gap', val('s-act-gap') + 'px');
    root.style.setProperty('--week-to-act', val('s-week-to-act') + 'px');

    // Re-render if we have data
    if (currentDiagramData) {
        renderDiagram(currentDiagramData);
    }
}

function val(id) {
    const el = document.getElementById(id);
    if (el) return el.value;
    return DEFAULTS[id] !== undefined ? String(DEFAULTS[id]) : '';
}

// Reset settings to defaults
const DEFAULTS = {
    's-canvas-width': 1000, 's-canvas-bg': '#ffffff', 's-canvas-transparent': true, 's-canvas-padding': 10,
    's-font-family': "'Montserrat', sans-serif",
    's-hito-font': 30, 's-title-font': 16, 's-subtitle-font': 18,
    's-act-title-font': 16, 's-act-sub-font': 18, 's-code-font': 10,
    's-color-orange': '#f57c20',
    's-color-main': '#2e5bcc',
    's-color-week': '#6f90cd',
    's-color-content': '#214fc7',
    's-color-activity': '#6f90cd',
    's-color-text': '#ffffff',
    's-conn-width': 2,
    's-conn-color': '#f57c20',
    's-corner-radius': 10, 's-arrow-size': 10, 's-arrow-type': 'triangle',
    's-border-radius': 13, 's-border-width': 2, 's-border-color': '#f57c20',
    's-main-width': 340, 's-week-width': 280, 's-act-width': 0, 's-node-padding': 16,
    's-week-gap': 25, 's-act-gap': 10, 's-week-to-act': 75,
    's-card-header': '#F16522', 's-card-body': '#DDDDDD', 's-card-footer': '#516BED',
    's-diag-hito-w': 360, 's-diag-sub-w': 300, 's-diag-sub-color': '#F57C20', 's-diag-row-gap': 25, 's-diag-sub-bg': '#DDDDDD',
    's-card-body-font': 20, 's-card-q-font': 20, 's-card-body-align': 'justify', 's-card-footer-align': 'center',
    's-card-body-lh': 1.35, 's-card-footer-lh': 1.3,
    's-card-title-font': 60, 's-card-header-h': 92, 's-card-footer-h': 194,
    's-card-corner-r': 35, 's-card-shadow-h': 12,
};

document.getElementById('reset-settings').addEventListener('click', () => {
    for (const [id, value] of Object.entries(DEFAULTS)) {
        const el = document.getElementById(id);
        if (el) {
            if (el.type === 'checkbox') el.checked = value;
            else el.value = value;
            
            // Sync initial HEX string to text input if it exists
            const hexInput = document.getElementById(`hex-${id.substring(2)}`);
            if (hexInput) hexInput.value = value;
        }
    }
    applySettings();
});

// ── Export Config ──
document.getElementById('export-config-btn').addEventListener('click', () => {
    const config = {};
    for (const id of Object.keys(DEFAULTS)) {
        const el = document.getElementById(id);
        if (el) {
            config[id] = el.type === 'checkbox' ? el.checked : el.value;
        }
    }
    const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'tarjetas_config.json';
    a.click();
    URL.revokeObjectURL(a.href);
});

// ── Import Config ──
document.getElementById('import-config-btn').addEventListener('click', () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                const config = JSON.parse(ev.target.result);
                for (const [id, value] of Object.entries(config)) {
                    const el = document.getElementById(id);
                    if (el) {
                        if (el.type === 'checkbox') el.checked = value;
                        else el.value = value;
                        const hexInput = document.getElementById(`hex-${id.substring(2)}`);
                        if (hexInput) hexInput.value = value;
                    }
                }
                applySettings();
            } catch (err) {
                alert('Error al importar configuración: ' + err.message);
            }
        };
        reader.readAsText(file);
    });
    input.click();
});

// Bind dual-color input syncing globally
document.querySelectorAll('.color-hex-group').forEach(group => {
    const colorInput = group.querySelector('input[type="color"]');
    const textInput = group.querySelector('input[type="text"]');
    if (colorInput && textInput) {
        colorInput.addEventListener('input', (e) => {
            textInput.value = e.target.value.toUpperCase();
        });
        textInput.addEventListener('input', (e) => {
            let val = e.target.value;
            if (!val.startsWith('#')) val = '#' + val;
            if (/^#[0-9A-F]{6}$/i.test(val)) {
                colorInput.value = val;
            }
        });
        textInput.addEventListener('blur', (e) => {
            let val = e.target.value;
            if (!val.startsWith('#')) val = '#' + val;
            if (!/^#[0-9A-F]{6}$/i.test(val)) {
                // Revert to valid color input state if invalid hex pasted
                e.target.value = colorInput.value.toUpperCase();
            } else {
                e.target.value = val.toUpperCase();
            }
        });
    }
});

/* ========================================================================
   INTERACTIVE NODE EDITOR
   ======================================================================== */
// Function to recursively find a node by ID in currentDiagramData
function findNodeById(node, id) {
    if (node.id === id) return node;
    if (node.children) {
        for (let child of node.children) {
            const found = findNodeById(child, id);
            if (found) return found;
        }
    }
    return null;
}

// Function to find a node and its parent context for structural tree mutations
function findParentAndNode(node, id, parent = null, index = -1) {
    if (node.id === id) return { parent, index, node };
    if (node.children) {
        for (let i = 0; i < node.children.length; i++) {
            const found = findParentAndNode(node.children[i], id, node, i);
            if (found) return found;
        }
    }
    return null;
}

// Function to mutate tree hierarchy
function mutateHierarchy(targetNodeInfo, newHierarchy, rootNode) {
    const { parent, index, node: targetNode } = targetNodeInfo;
    
    // Slice it out of its current parent array
    parent.children.splice(index, 1);
    targetNode.type = newHierarchy;
    
    if (newHierarchy === 'finalWeekNode') {
        const grandParentInfo = findParentAndNode(rootNode, parent.id);
        if (grandParentInfo && grandParentInfo.parent) {
            grandParentInfo.parent.children.splice(grandParentInfo.index + 1, 0, targetNode);
        }
        if (!targetNode.children) targetNode.children = [];
    } else if (newHierarchy === 'finalActivityNode') {
        if (index > 0) {
            const prevWeek = parent.children[index - 1];
            if (!prevWeek.children) prevWeek.children = [];
            prevWeek.children.push(targetNode);
        } else if (parent.children.length > 0) {
            const nextWeek = parent.children[0];
            if (!nextWeek.children) nextWeek.children = [];
            nextWeek.children.unshift(targetNode);
        } else {
            parent.children.push(targetNode);
        }
        if (targetNode.children) targetNode.children = [];
    }
}

/* ========================================================================
   STRUCTURE PANEL (RIGHT SIDEBAR)
   ======================================================================== */
window.openNodeEditor = function(id) { 
    const el = document.querySelector(`.structure-card[data-id="${id}"]`);
    if (el) {
        // Remove previous highlights
        document.querySelectorAll('.structure-card.highlight-card').forEach(n => n.classList.remove('highlight-card'));
        
        // Scroll into view safely
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        
        // Apply temporal highlight
        el.classList.add('highlight-card');
        setTimeout(() => el.classList.remove('highlight-card'), 1500);

        // Optional: Focus the first rich editor
        const editor = el.querySelector('.rich-editor');
        if (editor) editor.focus();
    }
};

function renderStructurePanel() {
    const list = document.getElementById('structure-content-list');
    if (!list) return;

    if (!currentDiagramData) {
        list.innerHTML = `<div class="placeholder" style="margin-top: 2rem;">
            <span style="color:var(--edtech-text-muted); font-size:0.85rem; text-align:center; display:block;">Carga un archivo DOCX para visualizar y editar todo el contenido estructurado aquí.</span>
        </div>`;
        return;
    }

    // ── Check if this is a Página de Inicio diagram ──
    if (currentDiagramData.diagram.type === 'paginaInicio') {
        renderStructurePanelPaginaInicio();
        return;
    }

    const root = currentDiagramData.diagram.nodes[0];
    if (!root) return;

    let html = '';

    // Function to build individual card HTML
    const buildCard = (node, label) => {
        const title = node.richTitle || (node.customTitle !== undefined ? node.customTitle : (node.text.title || ''));
        const subtitle = node.richSubtitle || (node.customSubtitle !== undefined ? node.customSubtitle : (node.text.subtitle || ''));
        const isActivityOrWeek = node.type !== 'hitoBox' && node.type !== 'mainNode';
        const isHideable = node.type !== 'hitoBox' && node.type !== 'mainNode';
        
        const hiddenClass = node.hidden ? ' is-hidden' : '';
        const titleFieldId = `rt-${node.id}-title`;
        const subFieldId = `rt-${node.id}-subtitle`;

        // Get current per-container settings
        const nodeFont = node.customFont || val('s-font-family');
        const nodeTitleSize = node.customTitleSize || (node.type === 'hitoBox' ? val('s-hito-font') : (isActivityOrWeek ? val('s-act-title-font') : val('s-title-font')));
        const nodeSubSize = node.customSubtitleSize || (isActivityOrWeek ? val('s-act-sub-font') : val('s-subtitle-font'));

        // Build mini swatches
        let swatchesHtml = '';
        if (isActivityOrWeek || isHideable) {
            swatchesHtml += `<div style="display:flex; gap: 0.2rem; align-items: center;">`;
            const hitoColors = ['#214fc7', '#6f90cd'];
            hitoColors.forEach(c => {
                const isActive = (node.customBgColor && node.customBgColor.toLowerCase() === c.toLowerCase()) ? 'border: 2px solid #fff;' : 'border: 1px solid rgba(255,255,255,0.2);';
                swatchesHtml += `<div class="node-swatch" data-id="${node.id}" data-color="${c}" title="Usar color ${c}" style="width:12px;height:12px;border-radius:50%;background-color:${c};cursor:pointer;${isActive}"></div>`;
            });
            if (node.customBgColor) {
               swatchesHtml += `<button class="btn-icon reset-swatch" data-id="${node.id}" title="Revertir a Default" style="margin-left:2px; padding:0;"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/></svg></button>`;
            }
            swatchesHtml += `</div>`;
        }

        let headerActions = `<div style="display:flex; align-items:center; gap:0.4rem; margin-left:auto;">`;
        if (isActivityOrWeek) {
            headerActions += `
            <select class="live-edit-hierarchy" data-id="${node.id}" style="width:auto; font-size:0.65rem; position:relative; z-index:2;">
                <option value="finalWeekNode" ${node.type === 'finalWeekNode' ? 'selected' : ''}>Semana</option>
                <option value="finalActivityNode" ${node.type === 'finalActivityNode' ? 'selected' : ''}>Actividad</option>
            </select>`;
        }
        headerActions += swatchesHtml;
        if (isHideable) {
            const iconSvg = node.hidden 
                ? `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>`
                : `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><line x1="2" y1="2" x2="22" y2="22"/></svg>`;
            headerActions += `
            <button class="btn-icon danger hide-node-btn" data-id="${node.id}" style="padding:0;">
                ${iconSvg}
            </button>`;
        }
        headerActions += `</div>`;

        const chevronSvg = `<button class="btn-icon card-collapse-btn" data-id="${node.id}" style="padding:0; margin-right:0.2rem;" title="Colapsar/Expandir">
            <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        </button>`;

        const eyeOpenSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>`;
        const eyeClosedSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>`;

        // Container controls row (width, border color, alignment)
        const curWidth = node.customWidth || '';
        const curBorder = node.customBorderColor || val('s-border-color') || '#F57C20';
        const curAlign = node.customTitleAlign || 'left';
        const containerCtrlsHtml = `
        <div class="container-controls">
            <label>W<input type="number" class="ctrl-width" data-id="${node.id}" value="${curWidth}" placeholder="auto" min="50" max="800" step="10"></label>
            <label>Borde<input type="color" class="ctrl-border-color" data-id="${node.id}" value="${curBorder}"></label>
            <div class="align-group">
                <button class="align-btn${curAlign === 'left' ? ' active' : ''}" data-id="${node.id}" data-align="left" title="Izquierda">⫷</button>
                <button class="align-btn${curAlign === 'center' ? ' active' : ''}" data-id="${node.id}" data-align="center" title="Centro">⬌</button>
                <button class="align-btn${curAlign === 'right' ? ' active' : ''}" data-id="${node.id}" data-align="right" title="Derecha">⫸</button>
            </div>
        </div>`;

        let cardHtml = `
        <div class="structure-card${hiddenClass}" data-id="${node.id}">
            <div style="display:flex; align-items:center; margin-bottom:0.1rem;">
                ${chevronSvg}
                <div class="structure-group-title" style="margin-bottom:0;">${label}</div>
                ${headerActions}
            </div>
            <div class="card-collapsible-body">
                ${containerCtrlsHtml}
                <div class="field-label">Título</div>
                <div class="rich-field-row">
                    ${createRichFieldHTML(titleFieldId, title, { placeholder: 'Título...', font: nodeFont, size: nodeTitleSize })}
                    <button class="btn-icon toggle-field-btn" data-id="${node.id}" data-field="title" title="Ocultar/Mostrar título">
                        ${node.hiddenTitle ? eyeClosedSvg : eyeOpenSvg}
                    </button>
                </div>
        `;

        if (node.type !== 'hitoBox') {
            cardHtml += `
                <div class="field-label">Subtítulo</div>
                <div class="rich-field-row">
                    ${createRichFieldHTML(subFieldId, subtitle, { placeholder: 'Subtítulo...', font: nodeFont, size: nodeSubSize })}
                    <button class="btn-icon toggle-field-btn" data-id="${node.id}" data-field="subtitle" title="Ocultar/Mostrar subtítulo">
                        ${node.hiddenSubtitle ? eyeClosedSvg : eyeOpenSvg}
                    </button>
                </div>`;
        }

        cardHtml += `</div></div>`;
        return cardHtml;
    };


    // 1. Root (Hito)
    html += `<div class="structure-group">`;
    html += buildCard(root, 'Hito Principal');
    html += `</div>`;

    // 2. Avance
    const avance = root.children[0];
    if (avance) {
        html += `<div class="structure-group">`;
        html += buildCard(avance, 'Información General');
        html += `</div>`;

        // 3. Weeks & Activities
        avance.children.forEach((week, wIndex) => {
            html += `<div class="structure-group">`;
            html += buildCard(week, `Bloque ${wIndex + 1}: Semana`);
            
            if (week.children && week.children.length > 0) {
                html += `<div class="structure-children">`;
                week.children.forEach((act) => {
                    html += buildCard(act, act.type === 'finalActivityNode' ? 'Actividad' : 'Sub-Bloque (Error)');
                });
                html += `</div>`;
            }
            html += `</div>`;
        });
    }

    list.innerHTML = html;
    bindStructurePanelEvents();
}

function bindStructurePanelEvents() {
    const list = document.getElementById('structure-content-list');
    if (!list || !currentDiagramData) return;
    const rootNode = currentDiagramData.diagram.nodes[0];

    // Collapse/Expand toggle for each card
    list.querySelectorAll('.card-collapse-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const card = e.currentTarget.closest('.structure-card');
            if (card) card.classList.toggle('is-collapsed');
        });
    });

    // Per-field visibility toggle (title / subtitle independently)
    list.querySelectorAll('.toggle-field-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const id = e.currentTarget.getAttribute('data-id');
            const field = e.currentTarget.getAttribute('data-field');
            const targetInfo = findParentAndNode(rootNode, id);
            if (targetInfo && targetInfo.node) {
                if (field === 'title') targetInfo.node.hiddenTitle = !targetInfo.node.hiddenTitle;
                else if (field === 'subtitle') targetInfo.node.hiddenSubtitle = !targetInfo.node.hiddenSubtitle;
                renderDiagram(currentDiagramData, true);
            }
        });
    });

    // ── Rich Text Toolbar Events ──
    bindRichTextEvents(list, (fieldId) => {
        // Parse fieldId: "rt-{nodeId}-{fieldKey}"
        const parts = fieldId.match(/^rt-(.+)-(title|subtitle)$/);
        if (!parts) return null;
        const nodeId = parts[1];
        const fieldKey = parts[2];
        const targetInfo = findParentAndNode(rootNode, nodeId);
        return targetInfo ? { node: targetInfo.node, fieldKey } : null;
    }, () => renderDiagram(currentDiagramData, false));

    // ── Container Controls ──
    // Width
    list.querySelectorAll('.ctrl-width').forEach(input => {
        input.addEventListener('change', (e) => {
            const id = e.target.getAttribute('data-id');
            const targetInfo = findParentAndNode(rootNode, id);
            if (targetInfo && targetInfo.node) {
                const v = parseInt(e.target.value);
                targetInfo.node.customWidth = v > 0 ? v : undefined;
                if (!v) delete targetInfo.node.customWidth;
                renderDiagram(currentDiagramData, false);
            }
        });
    });

    // Border Color
    list.querySelectorAll('.ctrl-border-color').forEach(input => {
        input.addEventListener('input', (e) => {
            const id = e.target.getAttribute('data-id');
            const targetInfo = findParentAndNode(rootNode, id);
            if (targetInfo && targetInfo.node) {
                targetInfo.node.customBorderColor = e.target.value;
                renderDiagram(currentDiagramData, false);
            }
        });
    });

    // Alignment
    list.querySelectorAll('.align-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            const align = e.currentTarget.getAttribute('data-align');
            const targetInfo = findParentAndNode(rootNode, id);
            if (targetInfo && targetInfo.node) {
                targetInfo.node.customTitleAlign = align;
                // Update button active states
                const group = e.currentTarget.closest('.align-group');
                if (group) group.querySelectorAll('.align-btn').forEach(b => b.classList.remove('active'));
                e.currentTarget.classList.add('active');
                renderDiagram(currentDiagramData, false);
            }
        });
    });

    // Hierarchy Select
    list.querySelectorAll('.live-edit-hierarchy').forEach(select => {
        select.addEventListener('change', (e) => {
            const id = e.target.getAttribute('data-id');
            const newHierarchy = e.target.value;
            const targetInfo = findParentAndNode(rootNode, id);
            if (targetInfo && targetInfo.node && targetInfo.node.type !== newHierarchy) {
                mutateHierarchy(targetInfo, newHierarchy, rootNode);
                renderDiagram(currentDiagramData, true); 
            }
        });
    });

    // Hide/Show Node Toggle
    list.querySelectorAll('.hide-node-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            const targetInfo = findParentAndNode(rootNode, id);
            if (targetInfo && targetInfo.node) {
                targetInfo.node.hidden = !targetInfo.node.hidden;
                renderDiagram(currentDiagramData, true); 
            }
        });
    });

    // Swatch Color Override
    list.querySelectorAll('.node-swatch').forEach(swatch => {
        swatch.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            const color = e.currentTarget.getAttribute('data-color');
            const targetInfo = findParentAndNode(rootNode, id);
            if (targetInfo && targetInfo.node) {
                targetInfo.node.customBgColor = color;
                renderDiagram(currentDiagramData, true); 
            }
        });
    });

    // Reset Swatch
    list.querySelectorAll('.reset-swatch').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            const targetInfo = findParentAndNode(rootNode, id);
            if (targetInfo && targetInfo.node) {
                delete targetInfo.node.customBgColor;
                renderDiagram(currentDiagramData, true); 
            }
        });
    });
}


/* ── Página de Inicio Structure Panel ── */
function renderStructurePanelPaginaInicio() {
    const list = document.getElementById('structure-content-list');
    if (!list || !currentDiagramData) return;

    const d = currentDiagramData.diagram;
    const eyeOpenSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>`;
    const eyeClosedSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>`;

    let html = '';

    // ── DIAGRAMA SECTION ──
    html += `<div class="sidebar-section-header" style="margin-bottom:0.6rem; padding:0.4rem 0.6rem; background: rgba(245,124,32,0.15); border-radius:4px; font-weight:700; font-size:0.85rem; color:var(--edtech-accent); text-transform:uppercase; letter-spacing:0.5px;">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="margin-right:4px; vertical-align:-2px"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
        Diagrama
    </div>`;

    // Course title card
    const courseTitle = d.customTitle || d.title || '';
    html += `<div class="structure-group">
        <div class="structure-card" data-id="course_title">
            <div class="structure-group-title">Título de Asignatura</div>
            <div class="card-collapsible-body">
                ${createRichFieldHTML('pi-course-title', d.richCourseTitle || courseTitle, { placeholder: 'Nombre de asignatura...', font: val('s-font-family'), size: val('s-hito-font') })}
            </div>
        </div>
    </div>`;

    // Each Hito (diagram section)
    d.nodes.forEach((node, idx) => {
        const hitoTitle = (node.customTitle !== undefined ? node.customTitle : `Hito ${node.text.hitoNum}`).replace(/"/g, '&quot;');
        const hitoType = (node.customSubtitle !== undefined ? node.customSubtitle : node.text.hitoType).replace(/"/g, '&quot;');
        const subtitle = (node.customDescription !== undefined ? node.customDescription : node.text.subtitle).replace(/"/g, '&quot;').replace(/</g, '&lt;');

        // Swatches — only for Hito 1
        let swatchesHtml = '';
        if (idx === 0) {
            swatchesHtml = `<div style="display:flex; gap:0.2rem; align-items:center;">`;
            const themeColors = [
                val('s-color-orange'),
                val('s-color-main'),
                val('s-color-week'),
                val('s-color-content'),
                val('s-color-activity')
            ];
            themeColors.forEach(c => {
                const isActive = (node.customBgColor && node.customBgColor.toLowerCase() === c.toLowerCase()) ? 'border:2px solid #fff;' : 'border:1px solid rgba(255,255,255,0.2);';
                swatchesHtml += `<div class="node-swatch" data-id="${node.id}" data-color="${c}" title="${c}" style="width:12px;height:12px;border-radius:50%;background:${c};cursor:pointer;${isActive}"></div>`;
            });
            if (node.customBgColor) {
                swatchesHtml += `<button class="btn-icon reset-swatch" data-id="${node.id}" title="Revertir" style="margin-left:2px;padding:0;"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/></svg></button>`;
            }
            swatchesHtml += `</div>`;
        }

        html += `<div class="structure-group">
            <div class="structure-card" data-id="${node.id}">
                <div style="display:flex; align-items:center; margin-bottom:0.1rem;">
                    <div class="structure-group-title" style="margin-bottom:0;">Hito ${idx + 1}</div>
                    ${swatchesHtml ? `<div style="display:flex; align-items:center; gap:0.4rem; margin-left:auto;">${swatchesHtml}</div>` : ''}
                </div>
                <div class="card-collapsible-body">
                    <div class="field-label">Título</div>
                    <div class="rich-field-row">
                        ${createRichFieldHTML('pi-' + node.id + '-title', node.richTitle || hitoTitle, { placeholder: 'Hito N...', font: node.customFont || val('s-font-family'), size: node.customTitleSize || val('s-hito-font') })}
                        <button class="btn-icon toggle-field-btn" data-id="${node.id}" data-field="title" title="Ocultar/Mostrar">
                            ${node.hiddenTitle ? eyeClosedSvg : eyeOpenSvg}
                        </button>
                    </div>
                    <div class="field-label">Tipo</div>
                    <div class="rich-field-row">
                        ${createRichFieldHTML('pi-' + node.id + '-subtitle', node.richSubtitle || hitoType, { placeholder: 'Tipo...', font: node.customFont || val('s-font-family'), size: node.customSubtitleSize || val('s-subtitle-font') })}
                        <button class="btn-icon toggle-field-btn" data-id="${node.id}" data-field="subtitle" title="Ocultar/Mostrar">
                            ${node.hiddenSubtitle ? eyeClosedSvg : eyeOpenSvg}
                        </button>
                    </div>
                    <div class="field-label">Descripción</div>
                    <div class="rich-field-row">
                        ${createRichFieldHTML('pi-' + node.id + '-description', node.richDescription || subtitle, { placeholder: 'Descripción...', font: node.customFont || val('s-font-family'), size: node.customDescriptionSize || val('s-subtitle-font') })}
                        <button class="btn-icon toggle-field-btn" data-id="${node.id}" data-field="description" title="Ocultar/Mostrar">
                            ${node.hiddenDescription ? eyeClosedSvg : eyeOpenSvg}
                        </button>
                    </div>
                </div>
            </div>
        </div>`;

    });

    // ── TARJETAS SECTION ──
    html += `<div class="sidebar-section-header" style="margin-top:1rem; margin-bottom:0.6rem; padding:0.4rem 0.6rem; background: rgba(81,107,237,0.15); border-radius:4px; font-weight:700; font-size:0.85rem; color:#516BED; text-transform:uppercase; letter-spacing:0.5px;">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="margin-right:4px; vertical-align:-2px"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>
        Tarjetas (Aprendizajes)
    </div>`;

    d.nodes.forEach((node, idx) => {
        const cardBody = (node.text.cardBody || '').replace(/"/g, '&quot;').replace(/</g, '&lt;');
        const cardQuestion = (node.text.cardQuestion || '').replace(/"/g, '&quot;').replace(/</g, '&lt;');

        html += `<div class="structure-group">
            <div class="structure-card" data-id="${node.id}_card">
                <div class="structure-group-title">Tarjeta Hito ${idx + 1}</div>
                <div class="card-collapsible-body">
                    <div class="field-row">
                        <textarea class="live-edit-pi-card-body" rows="2" data-idx="${idx}" placeholder="Competencia / aprendizaje..." spellcheck="true" lang="es">${cardBody}</textarea>
                    </div>
                    <div class="field-row">
                        <textarea class="live-edit-pi-card-question" rows="2" data-idx="${idx}" placeholder="Pregunta Autoevalúate..." spellcheck="true" lang="es">${cardQuestion}</textarea>
                    </div>
                </div>
            </div>
        </div>`;
    });

    list.innerHTML = html;
    bindPaginaInicioPanelEvents();
}

function bindPaginaInicioPanelEvents() {
    const list = document.getElementById('structure-content-list');
    if (!list || !currentDiagramData) return;
    const d = currentDiagramData.diagram;

    // Course title - rich text
    const courseTitleField = list.querySelector('[data-field-id="pi-course-title"]');
    if (courseTitleField) {
        bindRichTextEvents(courseTitleField.closest('.structure-card'), (fieldId) => {
            if (fieldId === 'pi-course-title') {
                return { node: d, fieldKey: 'title', 
                    // Override _onRichChange behavior for course title
                };
            }
            return null;
        }, () => renderDiagram(currentDiagramData, false));
        // Also bind the editor directly for course title special handling
        const editor = courseTitleField.querySelector('.rich-editor');
        if (editor) {
            editor.addEventListener('input', () => {
                d.richCourseTitle = editor.innerHTML;
                d.customTitle = editor.textContent;
                renderDiagram(currentDiagramData, false);
            });
        }
    }

    // ── Rich Text Events for PI hito fields ──
    bindRichTextEvents(list, (fieldId) => {
        // Parse fieldId: "pi-{nodeId}-{fieldKey}"
        const parts = fieldId.match(/^pi-(.+)-(title|subtitle|description)$/);
        if (!parts) return null;
        const nodeId = parts[1];
        const fieldKey = parts[2];
        const node = d.nodes.find(n => n.id === nodeId);
        return node ? { node, fieldKey } : null;
    }, () => renderDiagram(currentDiagramData, false));


    // Field visibility toggles
    list.querySelectorAll('.toggle-field-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            const field = e.currentTarget.getAttribute('data-field');
            const node = d.nodes.find(n => n.id === id);
            if (node) {
                if (field === 'title') node.hiddenTitle = !node.hiddenTitle;
                else if (field === 'subtitle') node.hiddenSubtitle = !node.hiddenSubtitle;
                else if (field === 'description') node.hiddenDescription = !node.hiddenDescription;
                renderDiagram(currentDiagramData, true);
            }
        });
    });

    // Color swatches
    list.querySelectorAll('.node-swatch').forEach(swatch => {
        swatch.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            const color = e.currentTarget.getAttribute('data-color');
            const node = d.nodes.find(n => n.id === id);
            if (node) { node.customBgColor = color; renderDiagram(currentDiagramData, true); }
        });
    });

    // Reset swatch
    list.querySelectorAll('.reset-swatch').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = e.currentTarget.getAttribute('data-id');
            const node = d.nodes.find(n => n.id === id);
            if (node) { delete node.customBgColor; renderDiagram(currentDiagramData, true); }
        });
    });

    // Card body (tarjetas section)
    list.querySelectorAll('.live-edit-pi-card-body').forEach(textarea => {
        setTimeout(() => { textarea.style.height = 'auto'; textarea.style.height = textarea.scrollHeight + 'px'; }, 10);
        textarea.addEventListener('input', (e) => {
            e.target.style.height = 'auto';
            e.target.style.height = e.target.scrollHeight + 'px';
            const idx = parseInt(e.target.getAttribute('data-idx'));
            if (d.nodes[idx]) {
                d.nodes[idx].text.cardBody = e.target.value;
                // Re-render if we're on that tarjeta tab
                if (currentActiveTab === `tarjeta_${idx}`) renderDiagram(currentDiagramData, false);
            }
        });
    });

    // Card question (tarjetas section)
    list.querySelectorAll('.live-edit-pi-card-question').forEach(textarea => {
        setTimeout(() => { textarea.style.height = 'auto'; textarea.style.height = textarea.scrollHeight + 'px'; }, 10);
        textarea.addEventListener('input', (e) => {
            e.target.style.height = 'auto';
            e.target.style.height = e.target.scrollHeight + 'px';
            const idx = parseInt(e.target.getAttribute('data-idx'));
            if (d.nodes[idx]) {
                d.nodes[idx].text.cardQuestion = e.target.value;
                if (currentActiveTab === `tarjeta_${idx}`) renderDiagram(currentDiagramData, false);
            }
        });
    });

    // openNodeEditor for Página de Inicio
    window.openNodeEditor = function(id) {
        // Strip _sub suffix for subtitle box clicks
        const cleanId = id.replace(/_sub$/, '');
        const el = list.querySelector(`.structure-card[data-id="${cleanId}"]`);
        if (el) {
            document.querySelectorAll('.structure-card.highlight-card').forEach(n => n.classList.remove('highlight-card'));
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
            el.classList.add('highlight-card');
            setTimeout(() => el.classList.remove('highlight-card'), 1500);
        }
    };
}

/* ── Export JSON Data (Structure & Customizations) ── */
document.getElementById('export-data-json').addEventListener('click', () => {
    if (!currentDiagramData) {
        alert('No hay datos cargados para exportar. Por favor carga un DOCX primero.');
        return;
    }
    const text = JSON.stringify(currentDiagramData, null, 2);
    const blob = new Blob([text], { type: 'application/json' });
    const a = document.createElement('a');
    a.download = generateSmartFilename() + '_datos.json';
    a.href = URL.createObjectURL(blob);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
});

/* ========================================================================
   ZOOM CONTROLS
   ======================================================================== */
let currentZoom = 1.0;
function applyZoom() {
    const wrapper = document.querySelector('.canvas-wrapper');
    if (wrapper) wrapper.style.transform = `scale(${currentZoom})`;
    const label = document.getElementById('zoom-value');
    if (label) label.textContent = `${Math.round(currentZoom * 100)}%`;
}
applyZoom(); // Apply default on load

document.getElementById('zoom-in').addEventListener('click', () => {
    currentZoom = Math.min(2, currentZoom + 0.1);
    applyZoom();
});
document.getElementById('zoom-out').addEventListener('click', () => {
    currentZoom = Math.max(0.2, currentZoom - 0.1);
    applyZoom();
});
document.getElementById('zoom-reset').addEventListener('click', () => {
    currentZoom = 1.0;
    applyZoom();
});

/* ========================================================================
   DOWNLOAD — PNG (Bulk for Página de Inicio) — Native SVG→PNG
   ======================================================================== */
function svgToPng(svgElem, scale) {
    return new Promise((resolve, reject) => {
        const svgData = new XMLSerializer().serializeToString(svgElem);
        // Use data: URI instead of Blob URL — required for foreignObject rendering
        const svgBase64 = btoa(unescape(encodeURIComponent(svgData)));
        const dataUrl = 'data:image/svg+xml;base64,' + svgBase64;
        const img = new Image();
        img.onload = () => {
            // Get dimensions — naturalWidth may be 0 for data URI SVGs, fallback to SVG attributes
            let w = img.naturalWidth || parseInt(svgElem.getAttribute('width')) || svgElem.viewBox?.baseVal?.width || 800;
            let h = img.naturalHeight || parseInt(svgElem.getAttribute('height')) || svgElem.viewBox?.baseVal?.height || 600;
            const canvas = document.createElement('canvas');
            canvas.width = w * scale;
            canvas.height = h * scale;
            const ctx = canvas.getContext('2d');
            ctx.scale(scale, scale);
            ctx.drawImage(img, 0, 0, w, h);
            resolve(canvas.toDataURL('image/png'));
        };
        img.onerror = (e) => { reject(e); };
        img.src = dataUrl;
    });
}

document.getElementById('download-btn').addEventListener('click', async () => {
    const isPaginaInicio = currentDiagramData && currentDiagramData.diagram && currentDiagramData.diagram.type === 'paginaInicio';
    
    if (isPaginaInicio) {
        if (typeof JSZip === 'undefined') {
            alert('Error: JSZip no está cargado.');
            return;
        }
        
        const nodes = currentDiagramData.diagram.nodes;
        const totalFiles = 1 + nodes.length;
        const ok = confirm(`¿Descargar ${totalFiles} archivos PNG en un ZIP?\n\n• ${generateSmartFilename()}.png (Diagrama)\n${nodes.map((n, i) => `• ${generateCardFilename(n.text.hitoNum)}.png (Tarjeta Hito ${n.text.hitoNum})`).join('\n')}`);
        if (!ok) return;
        
        try {
            const zip = new JSZip();
            const container = document.getElementById('diagram');
            const savedTab = currentActiveTab;
            
            // 1. Render and capture Diagram via native SVG→PNG
            currentActiveTab = 'diagram';
            container.innerHTML = '';
            const diagRenderer = new PaginaInicioRenderer(currentDiagramData, container);
            diagRenderer.render();
            const diagSvg = container.querySelector('svg');
            if (diagSvg) {
                try {
                    const dataUrl = await svgToPng(diagSvg, 2);
                    zip.file(generateSmartFilename() + '.png', dataUrl.split(',')[1], { base64: true });
                    console.log('[PNG] Diagram added');
                } catch(e) { console.error('Diagram PNG error:', e); }
            }
            
            // 2. Render and capture each Tarjeta via SVG→PNG
            for (let i = 0; i < nodes.length; i++) {
                container.innerHTML = '';
                try {
                    const cardRenderer = new TarjetaRenderer(nodes[i], container);
                    cardRenderer.render();
                    const cardSvg = container.querySelector('svg');
                    if (cardSvg) {
                        const dataUrl = await svgToPng(cardSvg, 2);
                        zip.file(generateCardFilename(nodes[i].text.hitoNum) + '.png', dataUrl.split(',')[1], { base64: true });
                        console.log(`[PNG] Tarjeta ${i + 1} added`);
                    }
                } catch(e) { console.error(`Card ${i} PNG error:`, e); }
            }
            
            // Restore original tab view
            currentActiveTab = savedTab;
            switchTab(savedTab);
            
            // Generate and download ZIP with all PNGs
            const zipBlob = await zip.generateAsync({ type: 'blob' });
            const zipName = generateSmartFilename() + '_png.zip';
            const url = URL.createObjectURL(zipBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = zipName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            console.log('[PNG] ZIP download triggered:', zipName);
        } catch (err) {
            console.error('[PNG] Fatal error:', err);
            alert('Error al generar los PNG: ' + err.message);
        }
        
    } else {
        // Single file — direct download
        const dc = document.getElementById('diagram');
        const svgElem = dc.querySelector('svg');
        if (svgElem) {
            try {
                const dataUrl = await svgToPng(svgElem, 2);
                downloadDataUrl(dataUrl, generateSmartFilename() + '.png');
            } catch(e) {
                console.error(e);
                alert('Error al generar la imagen.');
            }
        }
    }
});

function downloadDataUrl(dataUrl, filename) {
    const a = document.createElement('a');
    a.download = filename;
    a.href = dataUrl;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

/* ========================================================================
   EXPORT TO SVG (Bulk for Página de Inicio)
   ======================================================================== */
document.getElementById('download-svg-btn').addEventListener('click', async () => {
    const isPaginaInicio = currentDiagramData && currentDiagramData.diagram && currentDiagramData.diagram.type === 'paginaInicio';
    
    if (isPaginaInicio) {
        if (typeof JSZip === 'undefined') {
            alert('Error: JSZip no está cargado.');
            return;
        }
        
        const nodes = currentDiagramData.diagram.nodes;
        const totalFiles = 1 + nodes.length;
        const ok = confirm(`¿Descargar ${totalFiles} archivos SVG en un ZIP?\n\n• ${generateSmartFilename()}.svg (Diagrama)\n${nodes.map((n, i) => `• ${generateCardFilename(n.text.hitoNum)}.svg (Tarjeta Hito ${n.text.hitoNum})`).join('\n')}`);
        if (!ok) return;
        
        try {
            const zip = new JSZip();
            const container = document.getElementById('diagram');
            const savedTab = currentActiveTab;
            
            // 1. Diagram SVG
            container.innerHTML = '';
            const diagRenderer = new PaginaInicioRenderer(currentDiagramData, container);
            diagRenderer.render();
            const diagSvg = container.querySelector('svg');
            if (diagSvg) {
                const svgStr = new XMLSerializer().serializeToString(diagSvg);
                zip.file(generateSmartFilename() + '.svg', svgStr);
                console.log('[SVG] Diagram added');
            }
            
            // 2. Each tarjeta SVG
            for (let i = 0; i < nodes.length; i++) {
                container.innerHTML = '';
                try {
                    const cardRenderer = new TarjetaRenderer(nodes[i], container);
                    cardRenderer.render();
                    const cardSvg = container.querySelector('svg');
                    if (cardSvg) {
                        const svgStr = new XMLSerializer().serializeToString(cardSvg);
                        zip.file(generateCardFilename(nodes[i].text.hitoNum) + '.svg', svgStr);
                        console.log(`[SVG] Tarjeta ${i + 1} added`);
                    }
                } catch(e) { console.error(`[SVG] Error rendering tarjeta ${i}:`, e); }
            }
            
            // Restore
            currentActiveTab = savedTab;
            switchTab(savedTab);
            
            // Generate and download ZIP with all SVGs
            const zipBlob = await zip.generateAsync({ type: 'blob' });
            const zipName = generateSmartFilename() + '_svg.zip';
            const url = URL.createObjectURL(zipBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = zipName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            console.log('[SVG] ZIP download triggered:', zipName);
        } catch (err) {
            console.error('[SVG] Fatal error:', err);
            alert('Error al generar los SVG: ' + err.message);
        }
        
    } else {
        // Single SVG export
        const container = document.getElementById('diagram');
        const svgElem = container.querySelector('svg');
        if (!svgElem) { alert('Por favor carga un archivo primero.'); return; }
        downloadSvgBlob(svgElem, generateSmartFilename() + '.svg');
    }
});

function downloadSvgBlob(svgElem, filename) {
    const svgContent = new XMLSerializer().serializeToString(svgElem);
    const blob = new Blob([svgContent], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

/* ========================================================================
   DOWNLOAD — ZIP (PNG + SVG bundled, supports multi-document)
   ======================================================================== */
document.getElementById('download-zip-btn').addEventListener('click', async () => {
    console.log('[ZIP] Button clicked!');
    
    if (!currentDiagramData) {
        alert('No hay datos cargados. Carga un archivo primero.');
        return;
    }
    
    if (typeof JSZip === 'undefined') {
        alert('Error: JSZip no está cargado.');
        return;
    }

    // Determine if multi-doc mode
    const isMultiDoc = Object.keys(allDocuments).length > 1;

    try {
        console.log('[ZIP] Starting ZIP generation...');
        const zip = new JSZip();
        const container = document.getElementById('diagram');

        // Save current state
        const savedDiagramData = currentDiagramData;
        const savedFileName = currentSourceFileName;
        const savedHitoTab = currentHitoTab;
        const savedTab = currentActiveTab;

        if (isMultiDoc) {
            // ── Multi-document mode: iterate all documents ──
            const tabOrder = ['pagina_inicio', 'hito_1', 'hito_2', 'hito_3', 'hito_4', 'hito_5'];
            let totalFiles = 0;

            for (const key of tabOrder) {
                if (!allDocuments[key]) continue;
                const doc = allDocuments[key];
                currentDiagramData = doc.data;
                currentSourceFileName = doc.fileName;

                const isPaginaInicio = doc.data.diagram && doc.data.diagram.type === 'paginaInicio';
                const prefix = key === 'pagina_inicio' ? 'pag_inicio' : key.replace('_', '');

                // Add editable JSON data file
                const jsonFname = generateSmartFilename() + '_datos.json';
                zip.file(jsonFname, JSON.stringify(doc.data, null, 2));
                totalFiles++;
                console.log(`[ZIP] JSON added: ${jsonFname}`);

                if (isPaginaInicio) {
                    // Render Página de Inicio diagram
                    console.log(`[ZIP] Rendering ${prefix} diagram...`);
                    container.innerHTML = '';
                    const diagRenderer = new PaginaInicioRenderer(doc.data, container);
                    diagRenderer.render();
                    const diagSvg = container.querySelector('svg');
                    if (diagSvg) {
                        const fname = generateSmartFilename();
                        zip.file(`${fname}.svg`, new XMLSerializer().serializeToString(diagSvg));
                        try {
                            const pngData = await svgToPng(diagSvg, 2);
                            zip.file(`${fname}.png`, pngData.split(',')[1], { base64: true });
                            totalFiles += 2;
                        } catch(e) { console.error(`ZIP ${prefix} diagram PNG error:`, e); totalFiles += 1; }
                    }

                    // Render each tarjeta
                    const nodes = doc.data.diagram.nodes || [];
                    for (let i = 0; i < nodes.length; i++) {
                        console.log(`[ZIP] Rendering ${prefix} tarjeta ${i + 1}/${nodes.length}...`);
                        container.innerHTML = '';
                        try {
                            const cardRenderer = new TarjetaRenderer(nodes[i], container);
                            cardRenderer.render();
                            const cardSvg = container.querySelector('svg');
                            if (cardSvg) {
                                const fname = generateCardFilename(nodes[i].text.hitoNum);
                                zip.file(`${fname}.svg`, new XMLSerializer().serializeToString(cardSvg));
                                try {
                                    const pngData = await svgToPng(cardSvg, 2);
                                    zip.file(`${fname}.png`, pngData.split(',')[1], { base64: true });
                                    totalFiles += 2;
                                } catch(e) { console.error(`ZIP card ${i} PNG error:`, e); totalFiles += 1; }
                            }
                        } catch(e) { console.error(`[ZIP] Error rendering tarjeta ${i}:`, e); }
                    }
                } else {
                    // Render Hito diagram (SVGRenderer)
                    console.log(`[ZIP] Rendering ${prefix} diagram...`);
                    container.innerHTML = '';
                    currentActiveTab = 'diagram';
                    renderDiagram(doc.data, true);

                    // Wait for render
                    await new Promise(r => setTimeout(r, 100));

                    const diagSvg = container.querySelector('svg');
                    if (diagSvg) {
                        const fname = generateSmartFilename();
                        zip.file(`${fname}.svg`, new XMLSerializer().serializeToString(diagSvg));
                        try {
                            const pngData = await svgToPng(diagSvg, 2);
                            zip.file(`${fname}.png`, pngData.split(',')[1], { base64: true });
                            totalFiles += 2;
                        } catch(e) { console.error(`ZIP ${prefix} diagram PNG error:`, e); totalFiles += 1; }
                    }
                }
            }
            console.log(`[ZIP] Total files generated: ${totalFiles}`);

        } else {
            // ── Single-document mode (existing behavior) ──
            const isPaginaInicio = currentDiagramData && currentDiagramData.diagram && currentDiagramData.diagram.type === 'paginaInicio';

            // Add editable JSON data file
            zip.file(generateSmartFilename() + '_datos.json', JSON.stringify(currentDiagramData, null, 2));

            if (isPaginaInicio) {
                const nodes = currentDiagramData.diagram.nodes;

                // Diagram
                container.innerHTML = '';
                const diagRenderer = new PaginaInicioRenderer(currentDiagramData, container);
                diagRenderer.render();
                const diagSvg = container.querySelector('svg');
                if (diagSvg) {
                    zip.file(generateSmartFilename() + '.svg', new XMLSerializer().serializeToString(diagSvg));
                    try {
                        const pngData = await svgToPng(diagSvg, 2);
                        zip.file(generateSmartFilename() + '.png', pngData.split(',')[1], { base64: true });
                    } catch(e) { console.error('ZIP diagram PNG error:', e); }
                }

                // Tarjetas
                for (let i = 0; i < nodes.length; i++) {
                    container.innerHTML = '';
                    try {
                        const cardRenderer = new TarjetaRenderer(nodes[i], container);
                        cardRenderer.render();
                        const cardSvg = container.querySelector('svg');
                        if (cardSvg) {
                            const fname = generateCardFilename(nodes[i].text.hitoNum);
                            zip.file(fname + '.svg', new XMLSerializer().serializeToString(cardSvg));
                            try {
                                const pngData = await svgToPng(cardSvg, 2);
                                zip.file(fname + '.png', pngData.split(',')[1], { base64: true });
                            } catch(e) { console.error(`ZIP card ${i} PNG error:`, e); }
                        }
                    } catch(e) { console.error(`[ZIP] Error rendering tarjeta ${i}:`, e); }
                }
            } else {
                // Single non-paginaInicio doc
                const svgElem = container.querySelector('svg');
                if (svgElem) {
                    zip.file(generateSmartFilename() + '.svg', new XMLSerializer().serializeToString(svgElem));
                    try {
                        const pngData = await svgToPng(svgElem, 2);
                        zip.file(generateSmartFilename() + '.png', pngData.split(',')[1], { base64: true });
                    } catch(e) { console.error('ZIP PNG error:', e); }
                }
            }
        }

        // Restore original state
        currentDiagramData = savedDiagramData;
        currentSourceFileName = savedFileName;
        currentHitoTab = savedHitoTab;
        currentActiveTab = savedTab;
        if (isMultiDoc) {
            renderHitoTabs();
        }
        switchTab(savedTab);

        // Generate and download ZIP
        console.log('[ZIP] Generating ZIP file...');
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const zipName = (isMultiDoc ? currentSourceFileName.replace(/\\.docx$/i, '').replace(/[^a-z0-9]+/ig, '_') : generateSmartFilename()) + '_paquete.zip';
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = zipName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        console.log('[ZIP] Download triggered:', zipName);
    } catch (err) {
        console.error('[ZIP] Fatal error:', err);
        alert('Error al generar el ZIP: ' + err.message);
    }
});
