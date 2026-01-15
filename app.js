// PartsReg - Application Logic

let state = {
    screen: 'home',
    apiKey: localStorage.getItem('openai_api_key') || '',
    emailRecipients: localStorage.getItem('email_recipients') || '',
    partsList: JSON.parse(localStorage.getItem('parts_list') || '[]'),
    currentProject: { customer: '', project: '', date: new Date().toISOString().split('T')[0] },
    projectParts: [],
    partDescription: '',
    quantity: 1,
    suggestedPart: null,
    suggestedParts: [],
    isListening: false,
    isProcessing: false,
    mediaRecorder: null,
    audioChunks: []
};

function render() {
    const app = document.getElementById('app');
    app.innerHTML = `
        <div class="bg-white rounded-xl shadow-lg p-6 mb-4">
            <div class="flex items-center justify-between mb-6">
                <h1 class="text-3xl font-bold text-gray-900">Artikelregistrering</h1>
                <button onclick="showSettings()" class="p-2 hover:bg-gray-100 rounded-lg">
                    <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"></path></svg>
                </button>
            </div>
            <div id="content">${renderContent()}</div>
        </div>
        <div class="text-center text-sm text-gray-600"><p>PartsReg v0.1</p></div>
    `;
}

function renderContent() {
    if (state.screen === 'home') return renderHome();
    if (state.screen === 'parts') return renderParts();
    if (state.screen === 'settings') return renderSettings();
}

function renderHome() {
    return `
        <div class="space-y-4">
            <h2 class="text-2xl font-bold text-gray-800">Ny rapportering</h2>
            <div><label class="block text-sm font-medium text-gray-700 mb-1">Datum</label><input type="date" id="projectDate" value="${state.currentProject.date}" class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"></div>
            <div><label class="block text-sm font-medium text-gray-700 mb-1">Kund</label><input type="text" id="customerName" value="${state.currentProject.customer}" placeholder="Ange kundnamn" class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"></div>
            <div><label class="block text-sm font-medium text-gray-700 mb-1">Projekt</label><input type="text" id="projectName" value="${state.currentProject.project}" placeholder="Ange projektnummer/namn" class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"></div>
            <button onclick="startProject()" class="w-full bg-blue-600 text-white py-3 rounded-lg font-medium hover:bg-blue-700">Registrera artiklar</button>
        </div>
    `;
}

function renderParts() {
    const hasSuggestions = state.suggestedParts && state.suggestedParts.length > 0;

    let suggestionsHTML = '';
    if (hasSuggestions) {
        suggestionsHTML = `
            <div class="bg-yellow-50 p-3 rounded-lg border border-yellow-200 mb-3">
                <p class="text-sm font-medium text-yellow-800 mb-2">Flera möjliga matchningar hittades. Välj rätt artikel:</p>
                ${state.suggestedParts.map((part, index) => `
                    <label class="flex items-start p-2 hover:bg-yellow-100 rounded cursor-pointer mb-2 border border-transparent hover:border-yellow-300">
                        <input type="radio" name="partChoice" value="${index}" onclick="selectPart(${index})" class="mt-1 mr-3">
                        <div class="flex-1">
                            <p class="font-bold text-gray-900">${part.articleNumber}</p>
                            <p class="text-sm text-gray-700">${part.name}</p>
                            <p class="text-xs text-gray-600">Säkerhet: ${part.confidence}</p>
                        </div>
                    </label>
                `).join('')}
            </div>
        `;
    }

    return `
        <div class="space-y-4">
            <div class="flex items-center justify-between">
                <div><h2 class="text-xl font-bold text-gray-800">${state.currentProject.customer}</h2><p class="text-sm text-gray-600">${state.currentProject.project}</p><p class="text-xs text-gray-500">${new Date(state.currentProject.date).toLocaleDateString('sv-SE')}</p></div>
                <button onclick="resetProject()" class="bg-red-100 text-red-700 px-3 py-2 rounded-lg text-sm font-medium hover:bg-red-200">Nytt projekt</button>
            </div>
            <div class="bg-white p-4 rounded-lg border border-gray-200">
                <h3 class="font-medium text-gray-800 mb-3">Lägg till del</h3>
                <div class="flex gap-2 mb-3">
                    <input type="text" id="partDescription" value="${state.partDescription}" placeholder="Beskriv delen" class="flex-1 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500" oninput="state.partDescription=this.value">
                    <button onclick="toggleVoiceRecognition()" class="p-2 rounded-lg ${state.isListening ? 'bg-red-500' : 'bg-gray-200 hover:bg-gray-300'}"><svg class="w-5 h-5 ${state.isListening ? 'text-white animate-pulse' : ''}" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 11a7 7 0 01-7 7m0 0a7 7 0 01-7-7m7 7v4m0 0H8m4 0h4m-4-8a3 3 0 01-3-3V5a3 3 0 116 0v6a3 3 0 01-3 3z"></path></svg></button>
                </div>
                <button onclick="matchPartWithAI()" ${state.isProcessing ? 'disabled' : ''} class="w-full bg-green-600 text-white py-2 rounded-lg font-medium hover:bg-green-700 disabled:bg-gray-300 mb-3">${state.isProcessing ? 'Söker matchning...' : 'Hitta artikel'}</button>
                ${suggestionsHTML}
                ${state.suggestedPart && state.suggestedPart.articleNumber ? `
                    <div class="bg-green-50 p-3 rounded-lg border border-green-200 mb-3">
                        <p class="text-sm font-medium text-green-800">Föreslagen artikel:</p>
                        <p class="text-lg font-bold text-green-900">${state.suggestedPart.articleNumber}</p>
                        <p class="text-sm text-green-700">${state.suggestedPart.name}</p>
                        <p class="text-xs text-green-600 mt-1">Säkerhet: ${state.suggestedPart.confidence}</p>
                    </div>
                    <div class="flex gap-2 items-center">
                        <div class="flex-1"><label class="block text-sm font-medium text-gray-700 mb-1">Antal</label><input type="number" id="quantity" value="${state.quantity}" min="1" class="w-full px-3 py-2 border border-gray-300 rounded-lg" oninput="state.quantity=this.value"></div>
                        <button onclick="addPartToProject()" class="mt-6 bg-blue-600 text-white p-2 rounded-lg hover:bg-blue-700"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"></path></svg></button>
                    </div>
                ` : ''}
            </div>
            ${state.projectParts.length > 0 ? `
                <div class="bg-white rounded-lg border border-gray-200">
                    <div class="p-4 border-b"><h3 class="font-medium text-gray-800">Registrerade delar (${state.projectParts.length})</h3></div>
                    ${state.projectParts.map((p, i) => `
                        <div class="p-3 flex items-center justify-between border-b hover:bg-gray-50">
                            <div class="flex-1"><p class="font-medium text-gray-900">${p.articleNumber}</p><p class="text-sm text-gray-600">${p.name}</p><p class="text-xs text-gray-500">Antal: ${p.quantity}</p></div>
                            <button onclick="removePart(${i})" class="text-red-600 hover:text-red-700 p-2"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg></button>
                        </div>
                    `).join('')}
                </div>
                <button onclick="generateExcel()" class="w-full bg-blue-600 text-white py-3 rounded-lg font-medium hover:bg-blue-700 flex items-center justify-center gap-2"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>Generera Excel</button>
            ` : ''}
        </div>
    `;
}

function renderSettings() {
    return `
        <div class="space-y-4">
            <h2 class="text-2xl font-bold text-gray-800">Inställningar</h2>
            <div>
                <label class="block text-sm font-medium text-gray-700 mb-1">OpenAI API-nyckel</label>
                <input type="password" id="apiKey" value="${state.apiKey}" placeholder="sk-..." class="w-full px-3 py-2 border border-gray-300 rounded-lg">
                <div class="flex items-center justify-between mt-1">
                    <p class="text-xs text-gray-500">Hämta från platform.openai.com. Dela aldrig din nyckel med andra.</p>
                    <button onclick="clearApiKey()" class="text-xs text-red-600 hover:text-red-800 hover:underline">Radera API-nyckel</button>
                </div>
            </div>
            <div><label class="block text-sm font-medium text-gray-700 mb-1">E-postmottagare</label><input type="text" id="emailRecipients" value="${state.emailRecipients}" placeholder="namn@foretag.se" class="w-full px-3 py-2 border border-gray-300 rounded-lg"><p class="text-xs text-gray-500 mt-1">Separera flera adresser med komma</p></div>
            <div class="border-t border-gray-200 pt-4">
                <label class="block text-sm font-medium text-gray-700 mb-2">Artikellista</label>
                <label class="flex items-center justify-center w-full px-4 py-3 bg-blue-50 border-2 border-blue-300 border-dashed rounded-lg cursor-pointer hover:bg-blue-100">
                    <div class="text-center"><svg class="w-6 h-6 mx-auto mb-1 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg><span class="text-sm font-medium text-blue-600">Ladda upp Excel-fil (.xlsx, .xls)</span><p class="text-xs text-gray-500 mt-1">Kolumner: "Artikelnummer" och "Beskrivning"</p></div>
                    <input type="file" accept=".xlsx,.xls" onchange="handleFileUpload(event)" class="hidden">
                </label>
                ${state.partsList.length > 0 ? `
                <div class="mt-3 border border-gray-300 rounded-lg overflow-hidden">
                    <div class="max-h-48 overflow-y-auto">
                        <table class="w-full text-sm">
                            <thead class="bg-gray-100 sticky top-0">
                                <tr>
                                    <th class="text-left px-3 py-2 font-medium text-gray-700">Artikelnr</th>
                                    <th class="text-left px-3 py-2 font-medium text-gray-700">Beskrivning</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-gray-200">
                                ${state.partsList.map(p => `<tr class="hover:bg-gray-50"><td class="px-3 py-1.5 text-gray-600">${p.articleNumber}</td><td class="px-3 py-1.5 text-gray-800">${p.name}</td></tr>`).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
                <div class="flex items-center justify-between mt-1">
                    <p class="text-xs text-gray-500">${state.partsList.length} artiklar</p>
                    <button onclick="clearPartsList()" class="text-xs text-red-600 hover:text-red-800 hover:underline">Töm lista</button>
                </div>
                ` : `<p class="text-sm text-gray-500 mt-3">Ingen artikellista inläst. Ladda upp en Excel-fil för att komma igång.</p>`}
            </div>
            <div class="flex gap-2">
                <button onclick="saveSettings()" class="flex-1 bg-green-600 text-white py-3 rounded-lg font-medium hover:bg-green-700">Spara</button>
                <button onclick="cancelSettings()" class="flex-1 bg-gray-200 text-gray-800 py-3 rounded-lg font-medium hover:bg-gray-300">Avbryt</button>
            </div>
        </div>
    `;
}

function showSettings() {
    state.previousScreen = state.screen;
    state.screen = 'settings';
    render();
}

function cancelSettings() {
    state.screen = state.previousScreen || (state.currentProject.customer ? 'parts' : 'home');
    render();
}

function saveSettings() {
    state.apiKey = document.getElementById('apiKey').value;
    state.emailRecipients = document.getElementById('emailRecipients').value;
    localStorage.setItem('openai_api_key', state.apiKey);
    localStorage.setItem('email_recipients', state.emailRecipients);
    localStorage.setItem('parts_list', JSON.stringify(state.partsList));
    localStorage.setItem('previous_screen', state.previousScreen || 'home');
    alert('Inställningar sparade!');
    state.screen = state.previousScreen || (state.currentProject.customer ? 'parts' : 'home');
    render();
}

function clearPartsList() {
    if (!confirm('Är du säker på att du vill tömma artikellistan?')) return;
    state.partsList = [];
    render();
}

function clearApiKey() {
    if (!confirm('Är du säker på att du vill ta bort API-nyckeln?')) return;
    document.getElementById('apiKey').value = '';
}

function startProject() {
    const customer = document.getElementById('customerName').value;
    const project = document.getElementById('projectName').value;
    const date = document.getElementById('projectDate').value;
    if (!customer || !project) {
        alert('Fyll i både kund och projekt!');
        return;
    }
    state.currentProject = { customer, project, date };
    state.screen = 'parts';
    render();
}

function resetProject() {
    if (confirm('Börja om med nytt projekt?')) {
        state.currentProject = { customer: '', project: '', date: new Date().toISOString().split('T')[0] };
        state.projectParts = [];
        state.partDescription = '';
        state.quantity = 1;
        state.suggestedPart = null;
        state.suggestedParts = [];
        state.screen = 'home';
        render();
    }
}

async function toggleVoiceRecognition() {
    // Om vi redan spelar in, stoppa och skicka till Whisper
    if (state.isListening && state.mediaRecorder) {
        state.mediaRecorder.stop();
        return;
    }

    // Kontrollera att API-nyckel finns
    if (!state.apiKey) {
        alert('Ange API-nyckel i inställningarna för att använda röstinspelning.');
        showSettings();
        return;
    }

    // Starta inspelning
    try {
        const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
        state.audioChunks = [];
        state.mediaRecorder = new MediaRecorder(stream);

        state.mediaRecorder.ondataavailable = (e) => {
            if (e.data.size > 0) state.audioChunks.push(e.data);
        };

        state.mediaRecorder.onstop = async () => {
            // Stoppa mikrofonen
            stream.getTracks().forEach(track => track.stop());
            state.isListening = false;
            state.isProcessing = true;
            render();

            // Skicka till Whisper
            try {
                const audioBlob = new Blob(state.audioChunks, { type: 'audio/webm' });
                const formData = new FormData();
                formData.append('file', audioBlob, 'recording.webm');
                formData.append('model', 'whisper-1');
                formData.append('language', 'sv');

                const response = await fetch('https://api.openai.com/v1/audio/transcriptions', {
                    method: 'POST',
                    headers: { 'Authorization': `Bearer ${state.apiKey}` },
                    body: formData
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error?.message || 'Whisper API-fel');
                }

                const result = await response.json();
                state.partDescription = result.text;
            } catch (err) {
                alert('Kunde inte tolka röst: ' + err.message);
            }

            state.isProcessing = false;
            state.mediaRecorder = null;
            render();
        };

        state.mediaRecorder.start();
        state.isListening = true;
        render();
    } catch (err) {
        alert(err.name === 'NotAllowedError' ? 'Ge tillåtelse till mikrofonen' : 'Kunde inte starta mikrofon: ' + err.message);
    }
}

async function matchPartWithAI() {
    const desc = document.getElementById('partDescription').value;
    if (!desc.trim()) { alert('Beskriv delen först!'); return; }
    if (!state.apiKey) { alert('Ange API-nyckel i inställningarna!'); showSettings(); return; }
    if (state.partsList.length === 0) { alert('Lägg till artiklar i inställningarna!'); showSettings(); return; }
    state.partDescription = desc;
    state.isProcessing = true;
    render();
    try {
        const res = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${state.apiKey}` },
            body: JSON.stringify({
                model: 'gpt-4o-mini',
                messages: [
                    { role: 'system', content: 'Du är expert på installatörsmaterial. Matcha beskrivning mot artikellista. Om beskrivningen är specifik nog för EN tydlig matchning (t.ex. "90 vinkel 50mm"), returnera: {"matches":[{"articleNumber":"3","name":"90 vinkel 50 mm","confidence":"high"}]}. Om beskrivningen är bred och matchar flera artiklar (t.ex. bara "vinkel"), returnera ALLA matchande artiklar (2-3 st max) sorterade efter sannolikhet: {"matches":[{"articleNumber":"3","name":"90 vinkel 50 mm","confidence":"medium"},{"articleNumber":"5","name":"45 vinkel 50mm","confidence":"medium"}]}. Om ingen match alls: {"matches":[]}.' },
                    { role: 'user', content: `Artikellista:\n${JSON.stringify(state.partsList)}\n\nBeskrivning: "${desc}"\n\nVilken/vilka artiklar matchar?` }
                ],
                temperature: 0.3
            })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error.message);
        const content = data.choices[0].message.content;
        const match = content.match(/\{[\s\S]*\}/);
        if (match) {
            const result = JSON.parse(match[0]);
            if (result.matches && result.matches.length > 0) {
                if (result.matches.length === 1) {
                    state.suggestedPart = result.matches[0];
                    state.suggestedParts = [];
                } else {
                    state.suggestedParts = result.matches.slice(0, 3);
                    state.suggestedPart = null;
                }
            } else {
                alert('Ingen matchning hittades. Försök beskriv delen annorlunda.');
                state.suggestedPart = null;
                state.suggestedParts = [];
            }
        } else {
            alert('Kunde inte tolka AI-svar');
        }
    } catch (e) {
        alert('Fel vid AI-matchning: ' + e.message);
    }
    finally {
        state.isProcessing = false;
        render();
    }
}

function selectPart(index) {
    state.suggestedPart = state.suggestedParts[index];
    state.suggestedParts = [];
    render();
}

function addPartToProject() {
    if (!state.suggestedPart || !state.suggestedPart.articleNumber) {
        alert('Ingen artikel vald!');
        return;
    }
    const qty = parseInt(document.getElementById('quantity').value);
    state.projectParts.push({
        id: Date.now(),
        articleNumber: state.suggestedPart.articleNumber,
        name: state.suggestedPart.name,
        quantity: qty,
        description: state.partDescription
    });
    state.partDescription = '';
    state.quantity = 1;
    state.suggestedPart = null;
    state.suggestedParts = [];
    render();
}

function removePart(i) {
    state.projectParts.splice(i, 1);
    render();
}

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
        try {
            const data = new Uint8Array(ev.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(ws);
            // Hämta kolumnnamn från första raden för fallback
            const columns = Object.keys(json[0] || {});

            const parts = json.map(r => {
                // Sök artikelnummer - vanligaste först, sedan fallback till första kolumnen
                const num = r['Artikelnr'] || r['artikelnr'] || r['Artikelnummer'] || r['artikelnummer'] ||
                            r['ArtNr'] || r['artnr'] || r['Art.nr'] || r['art.nr'] ||
                            r['ArticleNumber'] || r['Artikel'] || r['Nr'] || r['nr'] ||
                            (columns[0] ? r[columns[0]] : null);

                // Sök beskrivning - vanligaste först, sedan fallback till andra kolumnen
                const name = r['Artikelbeskrivning'] || r['artikelbeskrivning'] || r['Beskrivning'] || r['beskrivning'] ||
                             r['Benämning'] || r['benämning'] || r['Namn'] || r['namn'] ||
                             r['Name'] || r['Description'] || r['description'] ||
                             (columns[1] ? r[columns[1]] : null);

                if (!num || !name) return null;
                return { articleNumber: String(num).trim(), name: String(name).trim() };
            }).filter(p => p !== null);
            if (parts.length === 0) {
                alert('Inga artiklar hittades. Kontrollera kolumnnamn.');
                return;
            }

            // Hitta vilka artikelnummer som redan finns
            const existingNums = new Set(state.partsList.map(p => p.articleNumber));
            const newParts = parts.filter(p => !existingNums.has(p.articleNumber));
            const skippedCount = parts.length - newParts.length;

            // Kontrollera dubbletter inom importfilen
            const nums = parts.map(p => p.articleNumber);
            const dupsInFile = [...new Set(nums.filter((n, i) => nums.indexOf(n) !== i))];

            // Slå ihop listorna och uppdatera state
            state.partsList = [...state.partsList, ...newParts];

            let msg = `✓ ${newParts.length} nya artiklar tillagda!`;
            if (skippedCount > 0) msg += `\n⏭️ ${skippedCount} redan befintliga hoppades över`;
            if (dupsInFile.length > 0) msg += `\n⚠️ ${dupsInFile.length} dubbletter i filen`;
            msg += `\n\nTotalt: ${state.partsList.length} artiklar`;
            alert(msg);
            render();
        } catch (err) {
            alert('Kunde inte läsa Excel-fil');
        }
    };
    reader.readAsArrayBuffer(file);
}

async function generateExcel() {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Delar');

    // Add header info (rows 1-3) with bold labels
    const row1 = ws.addRow(['Kund:', state.currentProject.customer]);
    const row2 = ws.addRow(['Projekt:', state.currentProject.project]);
    const row3 = ws.addRow(['Datum:', new Date(state.currentProject.date).toLocaleDateString('sv-SE')]);

    // Bold the labels in column A
    row1.getCell(1).font = { bold: true };
    row2.getCell(1).font = { bold: true };
    row3.getCell(1).font = { bold: true };

    // Empty row
    ws.addRow([]);

    // Table headers (row 5) - all bold
    const headerRow = ws.addRow(['Artikelnr', 'Beskrivning', 'Antal', 'Input']);
    headerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
        cell.border = {
            bottom: { style: 'thin', color: { argb: 'FF000000' } }
        };
    });

    // Data rows
    state.projectParts.forEach(p => {
        const row = ws.addRow([p.articleNumber, p.name, p.quantity, p.description]);
        // Make Input column (D) italic
        row.getCell(4).font = { italic: true };
    });

    // Auto-fit column widths
    ws.columns.forEach((column, i) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            const cellLength = cell.value ? String(cell.value).length : 0;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        });
        // Set width with some padding
        column.width = Math.max(maxLength + 3, i === 2 ? 8 : 12);
    });

    // Generate file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);

    // Create filename
    const filename = `delar_${state.currentProject.customer}_${state.currentProject.project}_${Date.now()}.xlsx`;

    // Download file
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();

    // Create email text with headers matching Excel
    let emailText = `Kund: ${state.currentProject.customer}\n`;
    emailText += `Projekt: ${state.currentProject.project}\n`;
    emailText += `Datum: ${new Date(state.currentProject.date).toLocaleDateString('sv-SE')}\n\n`;
    emailText += `DELAR:\n`;
    emailText += `${'─'.repeat(50)}\n\n`;

    // Use same headers as Excel: Artikelnr, Beskrivning, Antal, Input
    state.projectParts.forEach((p, i) => {
        emailText += `${i + 1}.\n`;
        emailText += `   Artikelnr: ${p.articleNumber}\n`;
        emailText += `   Beskrivning: ${p.name}\n`;
        emailText += `   Antal: ${p.quantity}\n`;
        emailText += `   Input: ${p.description}\n\n`;
    });

    // Open email client
    if (state.emailRecipients) {
        const subject = encodeURIComponent(`Artikelregistrering: ${state.currentProject.customer} - ${state.currentProject.project}`);
        const body = encodeURIComponent(`Hej,\n\nSe detaljer nedan. Excel-fil bifogas separat.\n\n${emailText}\nMed vänlig hälsning`);
        const mailtoLink = `mailto:${state.emailRecipients}?subject=${subject}&body=${body}`;
        window.location.href = mailtoLink;

        alert(`Excel-fil nedladdad!\n\nE-postprogrammet öppnas nu.\nBifoga filen manuellt: ${filename}`);
    } else {
        alert(`Excel-fil nedladdad!\n\nKonfigurera e-postmottagare i inställningarna för att kunna skicka direkt.`);
    }
}

// Initialize app
render();
