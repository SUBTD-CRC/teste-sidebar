import Sortable from 'sortablejs';
import * as XLSX from 'xlsx';

// Application State
let appData = [];
let selectedThemeId = null;
let selectedSubthemeId = null;
let selectedServiceId = null;
let activeEditorId = null; 
let activeEditorType = null; // 'theme', 'subtheme', 'service'

// DOM Elements
const listThemes = document.getElementById('list-themes');
const listSubthemes = document.getElementById('list-subthemes');
const listServices = document.getElementById('list-services');
const colSubthemes = document.getElementById('col-subthemes');
const colServices = document.getElementById('col-services');

const inspectorPanel = document.getElementById('inspector-panel');
const emptyState = document.getElementById('empty-state');
const editorContent = document.getElementById('editor-content');
const itemForm = document.getElementById('item-form');
const itemNameInput = document.getElementById('item-name');
const itemDescInput = document.getElementById('item-description');
const inspectorTypeBadge = document.getElementById('inspector-type');
const toastContainer = document.getElementById('toast-container');
const descriptionGroup = document.getElementById('description-group');
const excelFileInput = document.getElementById('excel-file-input');

// Context Actions
const deleteItemBtn = document.getElementById('delete-item-btn');

// Modals Setup
const addModal = document.getElementById('add-modal');
const addForm = document.getElementById('add-form');
const modalTitle = document.getElementById('modal-title');
const addTypeInput = document.getElementById('add-type');
const newItemDescGroup = document.getElementById('new-item-desc-group');

const LOCAL_STORAGE_KEY = '1746_local_data';

// Icons mapped from layout reference
const themeIcons = {
    'animais': 'fa-paw',
    'acessibilidade': 'fa-wheelchair',
    'assistência': 'fa-people-group',
    'anticorrupção': 'fa-bullhorn',
    'cidadania': 'fa-people-arrows',
    'conservação': 'fa-helmet-safety',
    'cultura': 'fa-masks-theater',
    'defesa civil': 'fa-shield-halved',
    'educação': 'fa-graduation-cap',
    'empresas': 'fa-store',
    'imóveis': 'fa-building',
    'iluminação': 'fa-lightbulb',
    'limpeza': 'fa-trash-can',
    'meio ambiente': 'fa-seedling',
    'mulher': 'fa-person-dress',
    'ordem pública': 'fa-person-military-pointing',
    'ouvidoria': 'fa-ear-listen',
    'processos': 'fa-file-contract',
    'procon': 'fa-cart-shopping',
    'proteção': 'fa-user-shield',
    'saúde': 'fa-notes-medical',
    'serviços urbanos': 'fa-city',
    'trabalho': 'fa-briefcase',
    'trânsito': 'fa-traffic-light',
    'transparência': 'fa-magnifying-glass',
    'transporte': 'fa-bus',
    'tributos': 'fa-file-invoice-dollar',
    'funerários': 'fa-cross',
    'turismo': 'fa-map-location-dot',
    'carnaval': 'fa-mask'
};

function getItemIcon(name, defaultIcon) {
    const lowerName = (name || '').toLowerCase();
    let iconClass = defaultIcon;
    for (const key in themeIcons) {
        if (lowerName.includes(key)) {
            iconClass = themeIcons[key];
            break;
        }
    }
    return `<i class="fa-solid ${iconClass}"></i>`;
}

// Utilities
const generateId = () => Math.random().toString(36).substr(2, 9);
const sanitizeHTML = (str) => {
    const temp = document.createElement('div');
    temp.textContent = str;
    return temp.innerHTML;
};

// Data Management
function saveDataLocally() {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(appData));
}

async function loadData() {
    const local = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (local) {
        try {
            appData = JSON.parse(local);
            renderThemes();
            return;
        } catch (e) {
            console.error("Local data corrupt, fetching initial.");
        }
    }

    try {
        const response = await fetch('/initial_data.json');
        if (!response.ok) throw new Error('Failed to fetch default JSON');
        appData = await response.json();
    } catch (e) {
        appData = [];
    }
    
    normalizeData(appData);
    saveDataLocally();
    renderThemes();
}

function normalizeData(data) {
    data.forEach(t => {
        if(!t.id) t.id = generateId();
        if(!t.description) t.description = "";
        if(!t.subthemes) t.subthemes = [];
        t.subthemes.forEach(s => {
            if(!s.id) s.id = generateId();
            if(!s.description) s.description = "";
            if(!s.services) s.services = [];
            s.services.forEach(srv => {
                if(!srv.id) srv.id = generateId();
                if(!srv.description) srv.description = "";
            });
        });
    });
}

// Rendering
function createItemHTML(item, type, isSelected) {
    const selectedClass = isSelected ? 'selected' : '';
    let iconHtml = '';
    if (type === 'theme') iconHtml = `<div class="item-icon">${getItemIcon(item.name, 'fa-layer-group')}</div>`;
    else if (type === 'subtheme') iconHtml = `<div class="item-icon">${getItemIcon(item.name, 'fa-folder-open')}</div>`;
    else iconHtml = `<div class="item-icon"><i class="fa-regular fa-circle-dot"></i></div>`;

    return `
        <div class="list-item ${selectedClass}" data-id="${item.id}" data-type="${type}">
            <div class="drag-handle"><i class="fa-solid fa-grip-vertical"></i></div>
            ${iconHtml}
            <span class="item-name">${sanitizeHTML(item.name || '(Sem nome)')}</span>
            <i class="fa-solid fa-chevron-right item-arrow"></i>
        </div>
    `;
}

function renderThemes() {
    listThemes.innerHTML = '';
    appData.forEach(item => {
        listThemes.insertAdjacentHTML('beforeend', createItemHTML(item, 'theme', item.id === selectedThemeId));
    });
    
    // Bind click events
    listThemes.querySelectorAll('.list-item').forEach(el => {
        el.addEventListener('click', (e) => {
            if (e.target.closest('.drag-handle')) return;
            selectedThemeId = el.dataset.id;
            selectedSubthemeId = null; // reset children
            selectedServiceId = null;
            openEditor(selectedThemeId, 'theme');
            renderThemes(); // Update active class
            renderSubthemes();
        });
    });
}

function renderSubthemes() {
    listSubthemes.innerHTML = '';
    colSubthemes.classList.add('hidden');
    colServices.classList.add('hidden');
    listServices.innerHTML = '';

    if (!selectedThemeId) return;

    const theme = appData.find(t => t.id === selectedThemeId);
    if (!theme) return;

    colSubthemes.classList.remove('hidden');
    theme.subthemes.forEach(item => {
        listSubthemes.insertAdjacentHTML('beforeend', createItemHTML(item, 'subtheme', item.id === selectedSubthemeId));
    });

    listSubthemes.querySelectorAll('.list-item').forEach(el => {
        el.addEventListener('click', (e) => {
            if (e.target.closest('.drag-handle')) return;
            selectedSubthemeId = el.dataset.id;
            selectedServiceId = null;
            openEditor(selectedSubthemeId, 'subtheme');
            renderSubthemes(); // Update active class
            renderServices();
        });
    });
}

function renderServices() {
    listServices.innerHTML = '';
    colServices.classList.add('hidden');

    if (!selectedThemeId || !selectedSubthemeId) return;

    const theme = appData.find(t => t.id === selectedThemeId);
    const subtheme = theme?.subthemes.find(s => s.id === selectedSubthemeId);
    if (!subtheme) return;

    colServices.classList.remove('hidden');
    subtheme.services.forEach(item => {
        listServices.insertAdjacentHTML('beforeend', createItemHTML(item, 'service', item.id === selectedServiceId));
    });

    listServices.querySelectorAll('.list-item').forEach(el => {
        el.addEventListener('click', (e) => {
            if (e.target.closest('.drag-handle')) return;
            selectedServiceId = el.dataset.id;
            openEditor(selectedServiceId, 'service');
            renderServices();
        });
    });
}

// Editor
function openEditor(id, type) {
    activeEditorId = id;
    activeEditorType = type;
    
    emptyState.classList.add('hidden');
    editorContent.classList.remove('hidden');

    let item = null;
    if (type === 'theme') {
        item = appData.find(t => t.id === id);
    } else if (type === 'subtheme') {
        const t = appData.find(t => t.id === selectedThemeId);
        item = t?.subthemes.find(s => s.id === id);
    } else {
        const t = appData.find(t => t.id === selectedThemeId);
        const s = t?.subthemes.find(s => s.id === selectedSubthemeId);
        item = s?.services.find(srv => srv.id === id);
    }

    if (!item) return;

    inspectorTypeBadge.textContent = type === 'theme' ? 'Tema' : type === 'subtheme' ? 'Subtema' : 'Serviço';
    inspectorTypeBadge.className = `badge type-${type}`;

    itemNameInput.value = item.name || '';
    itemDescInput.value = item.description || '';

    if (type === 'service') {
        descriptionGroup.style.display = 'none';
    } else {
        descriptionGroup.style.display = 'flex';
    }
}

itemForm.addEventListener('submit', (e) => {
    e.preventDefault();
    if (!activeEditorId) return;

    let item = null;
    if (activeEditorType === 'theme') {
        item = appData.find(t => t.id === activeEditorId);
    } else if (activeEditorType === 'subtheme') {
        const t = appData.find(t => t.id === selectedThemeId);
        item = t?.subthemes.find(s => s.id === activeEditorId);
    } else {
        const t = appData.find(t => t.id === selectedThemeId);
        const s = t?.subthemes.find(s => s.id === selectedSubthemeId);
        item = s?.services.find(srv => srv.id === activeEditorId);
    }

    item.name = itemNameInput.value.trim();
    if (activeEditorType !== 'service') {
        item.description = itemDescInput.value.trim();
    }

    saveDataLocally();
    showToast('Atualizado com sucesso!', 'success');
    
    // Refresh view
    if (activeEditorType === 'theme') renderThemes();
    if (activeEditorType === 'subtheme') renderSubthemes();
    if (activeEditorType === 'service') renderServices();
});

deleteItemBtn.addEventListener('click', () => {
    if (!activeEditorId || !confirm("Tem certeza que deseja excluir este item e todos os seus sub-itens?")) return;

    if (activeEditorType === 'theme') {
        appData = appData.filter(t => t.id !== activeEditorId);
        selectedThemeId = null;
        selectedSubthemeId = null;
        selectedServiceId = null;
    } else if (activeEditorType === 'subtheme') {
        const theme = appData.find(t => t.id === selectedThemeId);
        theme.subthemes = theme.subthemes.filter(s => s.id !== activeEditorId);
        selectedSubthemeId = null;
        selectedServiceId = null;
    } else {
        const theme = appData.find(t => t.id === selectedThemeId);
        const subtheme = theme.subthemes.find(s => s.id === selectedSubthemeId);
        subtheme.services = subtheme.services.filter(s => s.id !== activeEditorId);
        selectedServiceId = null;
    }

    saveDataLocally();
    showToast('Item excluído', 'info');
    
    emptyState.classList.remove('hidden');
    editorContent.classList.add('hidden');
    
    renderThemes();
    renderSubthemes();
    renderServices();
});

// Adding new elements
document.getElementById('add-theme-btn').addEventListener('click', () => {
    openAddModal('theme');
});
document.getElementById('add-subtheme-btn').addEventListener('click', () => {
    if(!selectedThemeId) return;
    openAddModal('subtheme');
});
document.getElementById('add-service-btn').addEventListener('click', () => {
    if(!selectedSubthemeId) return;
    openAddModal('service');
});

function openAddModal(type) {
    let titleStr = type === 'theme' ? 'Novo Tema' : type === 'subtheme' ? 'Novo Subtema' : 'Novo Serviço';
    modalTitle.textContent = titleStr;
    addTypeInput.value = type;
    addForm.reset();
    newItemDescGroup.style.display = type === 'service' ? 'none' : 'flex';
    addModal.showModal();
}

document.getElementById('close-modal-btn').addEventListener('click', () => addModal.close());
document.getElementById('cancel-modal-btn').addEventListener('click', () => addModal.close());

addForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const type = addTypeInput.value;
    
    const newItem = {
        id: generateId(),
        name: document.getElementById('new-item-name').value.trim(),
    };

    if (type !== 'service') {
        newItem.description = document.getElementById('new-item-description').value.trim();
    }
    
    if (type === 'theme') {
        newItem.subthemes = [];
        appData.push(newItem);
        selectedThemeId = newItem.id; 
    } else if (type === 'subtheme') {
        newItem.services = [];
        const t = appData.find(t => t.id === selectedThemeId);
        t.subthemes.push(newItem);
        selectedSubthemeId = newItem.id;
    } else if (type === 'service') {
        const t = appData.find(t => t.id === selectedThemeId);
        const s = t.subthemes.find(s => s.id === selectedSubthemeId);
        s.services.push(newItem);
        selectedServiceId = newItem.id;
    }

    saveDataLocally();
    addModal.close();
    showToast(`${type} adicionado!`, 'success');
    
    renderThemes();
    if(type === 'subtheme' || type === 'service') renderSubthemes();
    if(type === 'service') renderServices();
    
    // Automatically open editor for new item
    openEditor(newItem.id, type);
});

// Initialize Sortable Lists over columns
function initSortables() {
    new Sortable(listThemes, {
        group: 'shared-levels',
        handle: '.drag-handle', animation: 150, fallbackOnBody: true,
        onEnd: (evt) => {
            if (evt.to === listSubthemes) {
                if (!selectedThemeId) return;
                const movedItem = appData.splice(evt.oldIndex, 1)[0];
                
                // Prevent dragging a Theme into its own Subthemes list
                if (movedItem.id === selectedThemeId) {
                     appData.splice(evt.oldIndex, 0, movedItem);
                     showToast('Não é possível mover um tema para dentro de si mesmo.', 'error');
                     return;
                }

                const targetTheme = appData.find(t => t.id === selectedThemeId);
                
                const newSubtheme = {
                    id: movedItem.id,
                    name: movedItem.name,
                    description: movedItem.description,
                    services: []
                };
                
                const subthemesToAppend = movedItem.subthemes || [];
                targetTheme.subthemes.splice(evt.newIndex, 0, newSubtheme);
                targetTheme.subthemes.push(...subthemesToAppend);

                saveDataLocally();
                if (selectedThemeId === movedItem.id) {
                   selectedThemeId = targetTheme.id;
                   selectedSubthemeId = null;
                   selectedServiceId = null;
                }
                renderThemes();
                renderSubthemes();
                return;
            }

            if (evt.oldIndex === evt.newIndex) return;
            const itemObj = appData.splice(evt.oldIndex, 1)[0];
            appData.splice(evt.newIndex, 0, itemObj);
            saveDataLocally();
        }
    });

    new Sortable(listSubthemes, {
        group: 'shared-levels',
        handle: '.drag-handle', animation: 150, fallbackOnBody: true,
        onEnd: (evt) => {
            if (!selectedThemeId) return;
            const sourceTheme = appData.find(t => t.id === selectedThemeId);

            if (evt.to === listThemes) {
                const movedSubtheme = sourceTheme.subthemes.splice(evt.oldIndex, 1)[0];
                
                const newTheme = {
                    id: movedSubtheme.id,
                    name: movedSubtheme.name,
                    description: movedSubtheme.description,
                    subthemes: []
                };
                
                appData.splice(evt.newIndex, 0, newTheme);
                saveDataLocally();
                if (selectedSubthemeId === movedSubtheme.id) {
                    selectedSubthemeId = null;
                    selectedServiceId = null;
                }
                renderThemes();
                renderSubthemes();
                renderServices();
                return;
            }

            if (evt.oldIndex === evt.newIndex) return;
            const itemObj = sourceTheme.subthemes.splice(evt.oldIndex, 1)[0];
            sourceTheme.subthemes.splice(evt.newIndex, 0, itemObj);
            saveDataLocally();
        }
    });

    new Sortable(listServices, {
        handle: '.drag-handle', animation: 150, fallbackOnBody: true,
        onEnd: (evt) => {
            if (evt.oldIndex === evt.newIndex || !selectedSubthemeId) return;
            const theme = appData.find(t => t.id === selectedThemeId);
            const subtheme = theme.subthemes.find(s => s.id === selectedSubthemeId);
            const itemObj = subtheme.services.splice(evt.oldIndex, 1)[0];
            subtheme.services.splice(evt.newIndex, 0, itemObj);
            saveDataLocally();
        }
    });
}

// Notifications
function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    const icon = type === 'success' ? 'fa-check-circle' : type === 'error' ? 'fa-exclamation-circle' : 'fa-info-circle';
    toast.className = `toast ${type}`;
    toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span class="toast-message">${sanitizeHTML(message)}</span>`;
    toastContainer.appendChild(toast);
    setTimeout(() => {
        toast.style.animation = 'slideOut 0.3s forwards';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// Export Data (Excel)
document.getElementById('export-json-btn').addEventListener('click', () => {
    const flatData = [];
    appData.forEach(t => {
        if (!t.subthemes || t.subthemes.length === 0) {
            flatData.push({ 'Tema': t.name, 'Descrição do Tema': t.description, 'Subtema': '', 'Descrição do Subtema': '', 'Serviços': '' });
        } else {
            t.subthemes.forEach(s => {
                let servicesStr = (s.services || []).map(srv => srv.name).join(', ');
                flatData.push({ 'Tema': t.name, 'Descrição do Tema': t.description,  'Subtema': s.name, 'Descrição do Subtema': s.description, 'Serviços': servicesStr });
            });
        }
    });

    const worksheet = XLSX.utils.json_to_sheet(flatData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sidebar");
    XLSX.writeFile(workbook, "1746_sidebar_config.xlsx");
    showToast('Download do Excel iniciado', 'success');
});

// Import Data (Excel)
document.getElementById('import-excel-btn').addEventListener('click', () => {
    excelFileInput.click();
});

excelFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            parseFlatExcelToHierarchy(jsonData);
            e.target.value = ''; // reset file input
        } catch(err) {
            showToast('Erro ao ler arquivo Excel.', 'error');
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
});

function parseFlatExcelToHierarchy(flatJson) {
    const newHierarchy = [];
    const themeMap = new Map();

    flatJson.forEach(row => {
        const themeName = (row['Tema'] || '').trim();
        if (!themeName) return;

        if (!themeMap.has(themeName)) {
            const newTheme = {
                id: generateId(),
                name: themeName,
                description: row['Descrição do Tema'] || '',
                subthemes: []
            };
            themeMap.set(themeName, newTheme);
            newHierarchy.push(newTheme);
        }
        
        const themeRef = themeMap.get(themeName);
        const subName = (row['Subtema'] || '').trim();
        
        if (subName) {
            let subRef = themeRef.subthemes.find(s => s.name === subName);
            if (!subRef) {
                subRef = {
                    id: generateId(),
                    name: subName,
                    description: row['Descrição do Subtema'] || '',
                    services: []
                };
                themeRef.subthemes.push(subRef);
            }
            
            const servicesStr = row['Serviços'] || '';
            if (servicesStr) {
                const srvArr = servicesStr.split(',').map(s => s.trim()).filter(s => s);
                srvArr.forEach(srvName => {
                    if (!subRef.services.find(x => x.name === srvName)) {
                        subRef.services.push({ id: generateId(), name: srvName });
                    }
                });
            }
        }
    });

    appData = newHierarchy;
    saveDataLocally();
    
    // reset UI
    selectedThemeId = null;
    selectedSubthemeId = null;
    selectedServiceId = null;
    colSubthemes.classList.add('hidden');
    colServices.classList.add('hidden');
    emptyState.classList.remove('hidden');
    editorContent.classList.add('hidden');
    
    renderThemes();
    showToast('Backup Excel restaurado com sucesso!', 'success');
}

// Init
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    initSortables();
});
