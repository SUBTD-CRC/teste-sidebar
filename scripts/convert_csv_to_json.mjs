import fs from 'fs';
import { parse } from 'csv-parse/sync';

const csvPath = '../data/De X Para - Portal 1746 e SGRC - Trabalho CRC - Barbosa.csv';
const jsonPath = '../public/initial_data.json';

// Step 1: Read existing data for descriptions
let existingDescriptions = {
    themes: {},
    subthemes: {}
};

try {
    const oldData = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
    oldData.forEach(t => {
        existingDescriptions.themes[t.name] = t.description;
        (t.subthemes || []).forEach(s => {
            existingDescriptions.subthemes[`${t.name}|${s.name}`] = s.description;
        });
    });
} catch (e) {
    console.log("No existing data found or error reading it.");
}

// Step 2: Read and parse CSV
const csvContent = fs.readFileSync(csvPath, 'utf8');
const records = parse(csvContent, {
    columns: false,
    skip_empty_lines: true,
    from_line: 2 // Skip header
});

const hierarchy = {};

records.forEach(row => {
    let [themeName, subthemeName, servicesStr] = row;
    
    if (!themeName) return;
    themeName = themeName.trim();
    subthemeName = subthemeName ? subthemeName.trim() : '';
    servicesStr = servicesStr ? servicesStr.trim() : '';

    if (!hierarchy[themeName]) {
        hierarchy[themeName] = {
            id: themeName,
            name: themeName,
            description: existingDescriptions.themes[themeName] || "",
            subthemes: {}
        };
    }

    if (subthemeName) {
        if (!hierarchy[themeName].subthemes[subthemeName]) {
            hierarchy[themeName].subthemes[subthemeName] = {
                id: subthemeName,
                name: subthemeName,
                description: existingDescriptions.subthemes[`${themeName}|${subthemeName}`] || "",
                services: []
            };
        }

        if (servicesStr) {
            // Split by comma, but be careful with quotes (csv-parse already handled them if they were quoted in CSV, 
            // but column 3 is a comma separated string within a cell)
            const serviceNames = servicesStr.split(',').map(s => s.trim()).filter(s => s);
            serviceNames.forEach(name => {
                // Check if service already exists in this subtheme
                if (!hierarchy[themeName].subthemes[subthemeName].services.find(srv => srv.name === name)) {
                    hierarchy[themeName].subthemes[subthemeName].services.push({
                        id: name,
                        name: name
                    });
                }
            });
        }
    }
});

// Step 3: Convert hierarchy to array format
const finalData = Object.values(hierarchy).map(theme => {
    return {
        ...theme,
        subthemes: Object.values(theme.subthemes)
    };
});

// Step 4: Save
fs.writeFileSync(jsonPath, JSON.stringify(finalData, null, 2));
console.log(`Sucesso: Convertido CSV para ${jsonPath} (${finalData.length} temas)`);
