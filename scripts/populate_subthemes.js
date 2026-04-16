const fs = require('fs');
const XLSX = require('xlsx');

// 1. Read .env for API Key
const envContent = fs.readFileSync('../.env', 'utf8');
const apiKeyMatch = envContent.match(/VITE_GEMINI_API_KEY\s*=\s*([^\s]+)/);
const apiKey = apiKeyMatch ? apiKeyMatch[1] : null;

if (!apiKey) {
    console.error("ERRO: Chave da API (VITE_GEMINI_API_KEY) não encontrada no arquivo .env");
    process.exit(1);
}

// 2. Load Excel Files
console.log("Lendo arquivos Excel...");
let temaWB, subtemaWB;
try {
    temaWB = XLSX.readFile('../data/tema.xlsx');
    subtemaWB = XLSX.readFile('../data/subtema.xlsx');
} catch (e) {
    console.error("Erro ao ler arquivos Excel:", e.message);
    process.exit(1);
}

const temaData = XLSX.utils.sheet_to_json(temaWB.Sheets[temaWB.SheetNames[0]]);
const subtemaData = XLSX.utils.sheet_to_json(subtemaWB.Sheets[subtemaWB.SheetNames[0]]);

// Mapear descrições dos Temas para usar como contexto
const temaMap = {};
temaData.forEach(row => {
    temaMap[row['Tema']] = row['Descrição do Tema'];
});

// Selecionar alguns exemplos reais do tema.xlsx para o prompt (estilo Few-shot)
const exampleRows = temaData.filter((row, idx) => idx % 20 === 0).slice(0, 5);
const examplesText = exampleRows.map(row => `- Tema: ${row['Tema']}\n  Descrição: ${row['Descrição do Tema']}`).join('\n');

const promptBase = `Atue como um redator especialista em serviços públicos municipais. Sua tarefa é criar uma descrição extremamente objetiva para subtemas do portal de atendimento 1746 da prefeitura.

Regras estritas:
1. Use uma linguagem simples e direta, acessível a qualquer cidadão.
2. Escreva exata e unicamente UMA frase curta.
3. Não use verbos no imperativo ou ação (ex: evite "Peça", "Solicite", "Informe"). Inicie a frase dando foco ao "quê" usando substantivos (ex: "Canal para requerimentos de...", "Área destinada à resolução de...").
4. Apenas retorne o texto final da descrição. Nunca inclua aspas, introduções ("Aqui está:"), rótulos ou quebras de linha.

Exemplos de descrições ideais que já foram aprovadas para os Temas principais:
${examplesText}

Com base no estilo acima, escreva a descrição para o seguinte Subtema:`;

async function generateDescription(tema, subtema, exemplosNivel3) {
    const parentDesc = temaMap[tema] || 'Geral';
    const context = `Contexto do Tema Pai (${tema}): ${parentDesc}\nSubtema a descrever: ${subtema}\nExemplos de serviços deste subtema: ${exemplosNivel3}`;
    const promptText = `${promptBase}\n\n${context}`;

    try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{ parts: [{ text: promptText }] }]
            })
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.error?.message || `HTTP ${response.status}`);
        }

        const data = await response.json();
        let text = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || '';
        
        // Limpeza básica
        text = text.replace(/^"|"$/g, '').replace(/^'|'$/g, '').trim();
        
        return text;
    } catch (error) {
        console.error(`  [!] Erro na API para "${subtema}":`, error.message);
        return '--- Erro na geração ---';
    }
}

async function main() {
    const results = [];
    console.log(`Iniciando processamento de ${subtemaData.length} subtemas...`);

    for (let i = 0; i < subtemaData.length; i++) {
        const row = subtemaData[i];
        const tema = row['Proposta 2026 - Nível 1 (Categorias) - Consolidar'];
        const subtema = row['Nível 2 (Subcategorias) - Consolidar'];
        const exemplos = row['Nível 3 (a nível de exemplos)'] || '';

        console.log(`[${i + 1}/${subtemaData.length}] Processando: ${subtema}...`);
        
        const descricao = await generateDescription(tema, subtema, exemplos);
        
        results.push({
            'Tema': tema,
            'Subtema': subtema,
            'Descrição do Subtema': descricao,
            'Exemplos de Referência': exemplos
        });

        // Delay de 2 segundos (ajustado para ser um pouco mais rápido mas ainda seguro)
        await new Promise(r => setTimeout(r, 2000));
    }

    console.log("Gerando arquivo Excel final...");
    const newWB = XLSX.utils.book_new();
    const newWS = XLSX.utils.json_to_sheet(results);
    XLSX.utils.book_append_sheet(newWB, newWS, "Subtemas Descritos");
    XLSX.writeFile(newWB, '../data/saida.xlsx');

    console.log("\n====================================================");
    console.log("SUCESSO: O arquivo 'saida.xlsx' foi gerado!");
    console.log("====================================================\n");
}

main();
