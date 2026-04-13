import fs from 'fs';

const apiKey = process.env.VITE_GEMINI_API_KEY;
const dataRaw = fs.readFileSync('./public/initial_data.json', 'utf8');
const appData = JSON.parse(dataRaw);

const promptBase = `Atue como um redator especialista em serviços públicos municipais. Sua tarefa é criar uma descrição extremamente objetiva para os temas e subtemas fornecidos de um portal de atendimento 1746.
Regras estritas:
1. Use uma linguagem simples e direta, acessível a qualquer cidadão.
2. Escreva exata e unicamente UMA frase curta por item.
3. Inicie a frase dando foco ao "quê", usando apenas substantivos (nunca verbos no imperativo).
4. Retorne APENAS um array JSON válido contendo as descrições na mesma exata ordem que os itens enviados. Em caso de erro em um item específico, retorne string vazia "". Não use crases de markdown \`\`\`json no topo ou embaixo, apenas os colchetes. Exemplo perfeito: ["Descrição do primeiro item.", "Descrição do segundo item."]. NENHUM CARACTERE ALÉM DO JSON!!`;

async function batchProcess(items) {
    const listText = items.map((item, i) => `${i+1}. ${item.type === 'theme' ? 'Tema' : 'Subtema'}: ${item.name} ${item.parent ? '(Dentro do tema: ' + item.parent + ')' : ''}`).join('\n');
    const promptText = promptBase + `\n\nItens:\n` + listText;
    
    let text = '';
    try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ contents: [{ parts: [{ text: promptText }] }] })
        });
        
        if(!response.ok) throw new Error(await response.text());
        const json = await response.json();
        text = json.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || '[]';
        
        // Limpeza de Markdown
        text = text.replace(/```json\n?|```\n?|```/g, '').trim();
        
        // Tentar encontrar o array [ ... ] caso venha texto extra
        const start = text.indexOf('[');
        const end = text.lastIndexOf(']');
        if (start !== -1 && end !== -1) {
            text = text.substring(start, end + 1);
        }
        
        return JSON.parse(text);
    } catch(e) {
        console.error('Falha no processamento do lote. Erro:', e.message);
        return Array(items.length).fill('');
    }
}

async function run() {
    const todosProcessar = [];
    appData.forEach(t => {
        todosProcessar.push({ type: 'theme', name: t.name, parent: null, ref: t });
        (t.subthemes || []).forEach(s => {
            todosProcessar.push({ type: 'subtheme', name: s.name, parent: t.name, ref: s });
        });
    });

    console.log(`Total para processar: ${todosProcessar.length}`);
    const batchSize = 35; // Lotando max para processar rápido e diminuir RPM usage
    
    for (let i = 0; i < todosProcessar.length; i += batchSize) {
        const batch = todosProcessar.slice(i, i + batchSize);
        console.log(`Processando lote ${Math.floor(i/batchSize) + 1} de ${Math.ceil(todosProcessar.length/batchSize)}... (${batch.length} itens)`);
        
        const resultDesc = await batchProcess(batch);
        
        batch.forEach((item, index) => {
            if (resultDesc[index]) {
                item.ref.description = resultDesc[index];
            }
        });
        
        // Delay para evitar limite de requisição gratuita (15 RPM)
        if (i + batchSize < todosProcessar.length) {
            console.log('Aguardando 4 segundos de rate limit...');
            await new Promise(r => setTimeout(r, 4500));
        }
    }
    
    fs.writeFileSync('./public/initial_data.json', JSON.stringify(appData, null, 2));
    console.log("Arquivo initial_data.json atualizado e salvo com sucesso!");
}

run();
