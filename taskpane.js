import { WordService } from './word_service.js';
import { StorageService } from './storage_service.js';
import { MockData } from './mock_data.js';

/* global Office, lucide */

// --- GLOBAL STATE ---
let state = {
    currentTab: 'duAn',
    editingIndex: -1,
    duAn: {
        tenDuAn: "",
        goiThau: "",
        dvtc: "",
        daiDienCDT: "",
        tvgs: "",
        ngayKhoiCong: "",
        ngayHoanThanh: ""
    },
    soHDForExport: "",
    exportFolderHandle: null,
    exportFolderLabel: "",
    nhanSu: [],
    mayMoc: [],
    vatLieu: [],
    thiNghiem: [],
    outputMode: 'multiple'
};

// --- CONFIGURATION ---
const categories = {
    duAn: { title: "Dự án", fields: ["tenDuAn", "goiThau", "dvtc", "daiDienCDT", "tvgs", "ngayKhoiCong", "ngayHoanThanh"], labels: ["Tên dự án", "Tên gói thầu", "Đơn vị thi công", "Đại diện CDT", "Tư vấn giám sát", "Ngày khởi công", "Ngày hoàn thành"] },
    nhanSu: { title: "Nhân sự", fields: ["stt", "name", "role", "major", "phone"], labels: ["STT", "Họ và tên", "Chức danh", "Chuyên ngành", "Số điện thoại"] },
    mayMoc: { title: "Máy móc", fields: ["stt", "name", "unit", "qty", "owner", "status"], labels: ["STT", "Tên thiết bị", "Đơn vị tính", "Số lượng", "Chủ sở hữu", "Hình thức"] },
    vatLieu: { title: "Vật liệu", fields: ["stt", "name", "standard", "origin", "note"], labels: ["STT", "Tên vật tư", "Tiêu chuẩn", "Nguồn gốc", "Ghi chú"] },
    thiNghiem: { title: "Phòng TN", fields: ["stt", "dvtn", "address", "ptn", "func"], labels: ["STT", "Đơn vị TN", "Địa chỉ", "Tên phòng TN", "Chức năng"] }
};

// --- INITIALIZATION ---
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Word) {
        await initializeApp();
    }
});

async function initializeApp() {
    try {
        updateLog("Đang khởi tạo hệ thống...");
        await loadState();
        registerEvents();
        switchTab('duAn');
        updateLog("Hệ thống sẵn sàng");
    } catch (e) {
        console.error("Lỗi khởi tạo:", e);
        updateLog("Lỗi khởi tạo: " + e.message);
    }
}

// --- DATA & STATE ---
async function loadState() {
    // Ưu tiên load dữ liệu thực tế từ file Word đang mở (thông qua StorageService)
    const duAnSaved = await StorageService.getProjectData("duAn");
    if (duAnSaved && Object.keys(duAnSaved).length > 0) state.duAn = duAnSaved;
    
    const nhanSuSaved = await StorageService.getProjectData("nhanSu");
    if (nhanSuSaved && nhanSuSaved.length > 0) state.nhanSu = nhanSuSaved;
    
    const mayMocSaved = await StorageService.getProjectData("mayMoc");
    if (mayMocSaved && mayMocSaved.length > 0) state.mayMoc = mayMocSaved;
    
    const vatLieuSaved = await StorageService.getProjectData("vatLieu");
    if (vatLieuSaved && vatLieuSaved.length > 0) state.vatLieu = vatLieuSaved;
    
    const thiNghiemSaved = await StorageService.getProjectData("thiNghiem");
    if (thiNghiemSaved && thiNghiemSaved.length > 0) state.thiNghiem = thiNghiemSaved;

    // Load folder lưu trữ từ IndexedDB
    const folderHandle = await StorageService.getFolderHandle();
    if (folderHandle) {
        state.exportFolderHandle = folderHandle;
        state.exportFolderLabel = folderHandle.name;
    }

    // Load trạng thái đồng bộ (cờ Xuất/Tách, chế độ nén, số HĐ)
    const syncState = await StorageService.getProjectData("syncState");
    if (syncState) {
        state.hasExportedMaster = syncState.hasExportedMaster || false;
        state.hasSplitFiles = syncState.hasSplitFiles || false;
        state.outputMode = syncState.outputMode || 'multiple';
        state.soHDForExport = syncState.soHDForExport || "";
    }
}

async function saveState() {
    await StorageService.setProjectData("duAn", state.duAn);
    await StorageService.setProjectData("nhanSu", state.nhanSu);
    await StorageService.setProjectData("mayMoc", state.mayMoc);
    await StorageService.setProjectData("vatLieu", state.vatLieu);
    await StorageService.setProjectData("thiNghiem", state.thiNghiem);
    
    // Lưu trạng thái cờ và cấu hình xuất
    await StorageService.setProjectData("syncState", {
        hasExportedMaster: state.hasExportedMaster,
        hasSplitFiles: state.hasSplitFiles,
        outputMode: state.outputMode,
        soHDForExport: state.soHDForExport
    });
}

// --- UI CONTROLLER ---
function switchTab(tabId) {
    state.currentTab = tabId;
    document.querySelectorAll('[data-tab]').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });
    
    renderContent();
    lucide.createIcons();
}

function renderContent() {
    try {
        const container = document.getElementById('tabContent');
        if (!container) return;
        container.innerHTML = "";
        
        if (state.currentTab === 'duAn') {
            renderProjectForm(container);
        } else if (state.currentTab === 'xuatBan') {
            renderExportSettings(container);
        } else {
            renderList(container, state.currentTab);
        }
        
        // Tự động căn chỉnh lại độ cao cho các textarea khi load dữ liệu
        setTimeout(() => {
            document.querySelectorAll('textarea').forEach(ta => {
                ta.style.height = 'auto';
                const newHeight = Math.min(ta.scrollHeight, 150);
                ta.style.height = newHeight + 'px';
                ta.style.overflowY = ta.scrollHeight > 150 ? 'auto' : 'hidden';
            });
        }, 50);
    } catch (e) {
        console.error("Lỗi render:", e);
        const container = document.getElementById('tabContent');
        if (container) {
            container.innerHTML = `<div class="p-5 text-red-600 bg-red-50 rounded-xl border border-red-100 text-sm">
                <strong>Đã xảy ra lỗi khi hiển thị:</strong><br>${e.message}
            </div>`;
        }
    }
}

function renderProjectForm(container) {
    const config = categories.duAn;
    const form = document.createElement("div");
    form.className = "space-y-4 bg-white p-5 rounded-[1.5rem] border border-slate-100 shadow-sm";
    
    config.fields.forEach((field, i) => {
        const div = document.createElement("div");
        const isDate = field === 'ngayKhoiCong' || field === 'ngayHoanThanh';
        const extraClass = isDate ? 'date-picker-input' : '';
        
        // Safety check for MockData
        const mockVal = (MockData && MockData.duAn) ? MockData.duAn[field] : null;
        const placeholder = `VD: ${mockVal || config.labels[i]}`;
        
        if (isDate) {
            div.innerHTML = `
                <label class="text-[10px] font-black text-slate-400 uppercase mb-1 ml-1">${config.labels[i]}</label>
                <input type="text" spellcheck="false" data-field="${field}" value="${state.duAn[field] || ''}" class="input-field project-input ${extraClass}" placeholder="${placeholder}">
            `;
        } else {
            div.innerHTML = `
                <label class="text-[10px] font-black text-slate-400 uppercase mb-1 ml-1">${config.labels[i]}</label>
                <textarea spellcheck="false" data-field="${field}" class="input-field project-input resize-y py-2 w-full" style="height: auto; min-height: 2.25rem;" rows="1" placeholder="${placeholder}">${state.duAn[field] || ''}</textarea>
            `;
        }
        form.appendChild(div);
    });
    
    container.appendChild(form);
    
    // Khởi tạo Flatpickr cho tiện ích lịch cực đẹp và chuẩn DD/MM/YYYY
    setTimeout(() => {
        if (typeof flatpickr !== 'undefined') {
            flatpickr(".date-picker-input", {
                dateFormat: "d/m/Y",
                locale: "vn",
                allowInput: true,
                altInput: true,
                altFormat: "d/m/Y"
            });
        }
    }, 50);
    
    // Real-time Save for Project Info
    form.querySelectorAll('.project-input').forEach(input => {
        input.onchange = async () => {
            state.duAn[input.dataset.field] = input.value;
            await saveState();
        };
    });
}

function renderList(container, type) {
    const items = state[type] || [];
    const config = categories[type];

    const listCard = document.createElement('div');
    listCard.className = 'bg-white p-4 rounded-[1.5rem] border border-slate-100 shadow-sm';

    const header = document.createElement('div');
    header.className = 'flex items-center justify-between mb-3';
    header.innerHTML = `<h3 class="font-black text-sm text-slate-700">${config.title}</h3>`;

    const btnAdd = document.createElement('button');
    btnAdd.className = 'h-8 px-3 bg-indigo-600 text-white rounded-lg text-xs font-bold hover:bg-indigo-700';
    btnAdd.innerText = 'Thêm dòng mới';
    btnAdd.onclick = async () => {
        const newRow = config.fields.map((field, i) => i === 0 ? (items.length + 1).toString() : '');
        if (type === 'mayMoc') newRow[5] = 'Đi thuê';
        state[type].push(newRow);
        await saveState();
        renderContent();
        lucide.createIcons();
    };
    header.appendChild(btnAdd);
    listCard.appendChild(header);

    if (items.length === 0) {
        const emptyMessage = document.createElement('div');
        emptyMessage.className = 'p-6 text-center text-slate-400 italic';
        emptyMessage.innerText = 'Chưa có dữ liệu. Nhấn "Thêm dòng mới" để nhập.';
        listCard.appendChild(emptyMessage);
        container.appendChild(listCard);
        return;
    }

    const tableWrapper = document.createElement('div');
    tableWrapper.className = 'overflow-x-auto custom-scrollbar border rounded-xl';
    
    const table = document.createElement('table');
    table.className = 'w-full border-collapse text-[12px] break-words';
    table.style.minWidth = '600px'; // Đảm bảo bảng không bị bóp quá hẹp

    // Tính toán tỷ lệ % cho các cột để không bị vỡ khung
    const thead = document.createElement('thead');
    let wMap = { stt: '8%' };
    if (type === 'mayMoc') wMap = { name: '28%', unit: '12%', qty: '12%', owner: '28%', status: '12%' };
    else if (type === 'nhanSu') wMap = { name: '30%', role: '22%', major: '20%', phone: '20%' };
    else if (type === 'vatLieu') wMap = { name: '28%', standard: '20%', origin: '20%', note: '24%' };
    else if (type === 'thiNghiem') wMap = { dvtn: '23%', address: '24%', ptn: '19%', func: '26%' };

    let headersHtml = `
        <th class="border px-1 py-1" style="width: 8%;">
            <div class="flex flex-col items-center gap-1">
                <input type="checkbox" id="selectAll_${type}" class="w-3.5 h-3.5 text-indigo-600 border-slate-300 rounded focus:ring-indigo-500 cursor-pointer" title="Chọn tất cả">
                <span class="text-[9px]">STT</span>
            </div>
        </th>`;
    config.fields.forEach((field, idx) => {
        if (field === 'stt') return;
        headersHtml += `<th class="border px-1 py-2" style="width: ${wMap[field] || 'auto'};">${config.labels[idx]}</th>`;
    });

    thead.innerHTML = `<tr class="bg-slate-100 text-slate-600 uppercase text-[9px] font-bold">${headersHtml}</tr>`;
    table.appendChild(thead);

    // Lắng nghe sự kiện "Chọn tất cả"
    setTimeout(() => {
        const chkSelectAll = thead.querySelector(`#selectAll_${type}`);
        if (chkSelectAll) {
            chkSelectAll.onchange = (e) => {
                const checkboxes = tbody.querySelectorAll(`input[name="selectRow_${type}"]`);
                checkboxes.forEach(cb => { cb.checked = e.target.checked; });
            };
        }
    }, 0);

    const tbody = document.createElement('tbody');
    items.forEach((item, idx) => {
        const row = document.createElement('tr');
        row.className = idx % 2 === 0 ? 'bg-white' : 'bg-slate-50';

        row.innerHTML = `<td class="border px-1 py-1 text-center font-bold">
            <label class="cursor-pointer flex items-center justify-center gap-1.5 group" title="Click để chọn xóa">
               <input type="checkbox" name="selectRow_${type}" value="${idx}" class="w-3.5 h-3.5 text-indigo-600 border-slate-300 rounded focus:ring-indigo-500 cursor-pointer">
               <span class="group-hover:text-indigo-600 text-[10px]">${item[0]}</span>
            </label>
        </td>`;

        config.fields.forEach((field, i) => {
            if (field === 'stt') return;
            const val = item[i] || '';
            const mockArray = MockData[type] && MockData[type][0] ? MockData[type][0] : [];
            const placeholder = mockArray[i] ? `VD: ${mockArray[i]}` : config.labels[i];
            
            let inputElement;
            
            // Xử lý riêng cho trường "Hình thức" (chỉ đọc) của máy móc
            if (type === 'mayMoc' && field === 'status') {
                inputElement = document.createElement('input');
                inputElement.type = 'text';
                inputElement.spellcheck = false;
                inputElement.readOnly = true;
                inputElement.className = 'w-full border border-slate-200 rounded px-2 py-1 text-xs bg-slate-50 text-slate-600 cursor-not-allowed h-8';
            } else if (field === 'qty') {
                // Trường số lượng dạng number (có nút tăng giảm)
                inputElement = document.createElement('input');
                inputElement.type = 'number';
                inputElement.min = '0';
                inputElement.className = 'w-full border border-slate-200 rounded px-2 py-1 text-xs bg-white text-center h-8 cursor-pointer';
                inputElement.placeholder = placeholder;
            } else {
                // Sử dụng Textarea cho tất cả các trường nhập liệu để hỗ trợ multiline
                inputElement = document.createElement('textarea');
                inputElement.rows = 1; // Hiển thị sẵn 1 dòng
                inputElement.spellcheck = false;
                inputElement.className = 'w-full border border-slate-200 rounded px-1 py-1 text-xs bg-white resize-y custom-scrollbar';
                inputElement.style.minHeight = '1.75rem';
                inputElement.style.maxHeight = '150px';
                inputElement.placeholder = placeholder;
            }
            
            inputElement.value = val;
            
            inputElement.onblur = async (e) => {
                // Chỉ xử lý nếu không phải trường readonly
                if (!(type === 'mayMoc' && field === 'status')) {
                    let finalVal = e.target.value;
                    if (field === 'qty' && finalVal) finalVal = String(finalVal).padStart(2, '0');
                    if (type === 'nhanSu' && field === 'phone' && finalVal) {
                        finalVal = finalVal.replace(/(?:0|\+84)[\d\.\-\s]{9,13}/g, match => {
                            let d = match.replace(/\D/g, '');
                            if (d.startsWith('84')) d = '0' + d.substring(2);
                            if (d.length >= 10 && d.startsWith('0')) return d.substring(0, 10).replace(/(\d{4})(\d{3})(\d{3})/, '$1.$2.$3');
                            return match;
                        });
                    }
                    e.target.value = finalVal; // Cập nhật lại UI thực tế
                    state[type][idx][i] = finalVal;
                }

                if (type === 'mayMoc' && field === 'owner') {
                    const ownerName = normalize(state[type][idx][4]);
                    const companyName = normalize(state.duAn.dvtc);
                    state[type][idx][5] = (ownerName.includes(companyName) || companyName.includes(ownerName)) ? 'Sở hữu' : 'Đi thuê';
                }

                if (type === 'mayMoc' && field !== 'owner' && field !== 'status') {
                    const ownerName = normalize(state[type][idx][4]);
                    const companyName = normalize(state.duAn.dvtc);
                    state[type][idx][5] = (ownerName.includes(companyName) || companyName.includes(ownerName)) ? 'Sở hữu' : 'Đi thuê';
                }

                await saveState();
                // Cần render lại để cập nhật trường Hình thức
                if (type === 'mayMoc' && (field === 'owner' || field === 'status')) {
                    renderContent();
                    lucide.createIcons();
                }
            };

            const cell = document.createElement('td');
            cell.className = 'border p-1';
            cell.appendChild(inputElement);
            row.appendChild(cell);
        });

        // Đã xóa cột Thao tác từng dòng

        tbody.appendChild(row);
    });

    table.appendChild(tbody);
    tableWrapper.appendChild(table);
    listCard.appendChild(tableWrapper);

    // Thêm các nút thao tác chung ở dưới cùng bảng
    const actionContainer = document.createElement('div');
    actionContainer.className = 'flex gap-2 mt-4';
    
    const btnAddBottom = document.createElement('button');
    btnAddBottom.className = 'flex-1 h-10 border-2 border-dashed border-slate-300 text-slate-500 rounded-xl hover:bg-slate-50 hover:border-indigo-400 hover:text-indigo-600 transition-all font-bold text-xs flex items-center justify-center gap-2';
    btnAddBottom.innerHTML = '<i data-lucide="plus" size="16"></i> THÊM DÒNG MỚI';
    btnAddBottom.onclick = btnAdd.onclick;
    
    const btnDelBottom = document.createElement('button');
    btnDelBottom.className = 'w-24 h-10 border-2 border-dashed border-red-200 text-red-500 rounded-xl hover:bg-red-50 hover:border-red-400 hover:text-red-700 transition-all font-bold text-xs flex items-center justify-center gap-2';
    btnDelBottom.innerHTML = '<i data-lucide="trash-2" size="14"></i> XÓA';
    btnDelBottom.onclick = async () => {
        const checkedBoxes = container.querySelectorAll(`input[name="selectRow_${type}"]:checked`);
        if (checkedBoxes.length === 0) {
            showToast('Kiểm tra bảng: Vui lòng tích chọn các dòng ở cột STT để xóa!', 'warning');
            return;
        }

        // Lấy danh sách index cần xóa và sắp xếp giảm dần (để splice không bị lệch index)
        const indicesToDelete = Array.from(checkedBoxes).map(cb => parseInt(cb.value, 10)).sort((a, b) => b - a);
        
        indicesToDelete.forEach(idx => {
            state[type].splice(idx, 1);
        });

        // Đánh lại STT cho toàn bảng
        state[type].forEach((itemRow, i2) => { itemRow[0] = (i2 + 1).toString(); });
        
        await saveState();
        renderContent();
        lucide.createIcons();
        showToast(`Đã xóa ${indicesToDelete.length} dòng dữ liệu!`, "success");
    };
    
    actionContainer.appendChild(btnAddBottom);
    actionContainer.appendChild(btnDelBottom);
    listCard.appendChild(actionContainer);

    container.appendChild(listCard);
}

function renderExportSettings(container) {
    container.innerHTML = `
        <div class="p-5 bg-white rounded-[1.5rem] border border-slate-100 shadow-sm space-y-4">
            <h4 class="text-[10px] font-black text-slate-400 uppercase tracking-widest">Cấu hình đóng gói</h4>
            <label class="flex items-center gap-3 p-4 bg-slate-50 rounded-2xl cursor-pointer border-2 border-transparent hover:border-indigo-100 transition-all">
                <input type="radio" name="exportMode" value="zip" ${state.outputMode === 'zip' ? 'checked' : ''} class="w-5 h-5 text-indigo-600">
                <span class="text-sm font-bold text-slate-700">Tạo tệp Nén ZIP (.zip)</span>
            </label>
            <label class="flex items-center gap-3 p-4 bg-slate-50 rounded-2xl cursor-pointer border-2 border-transparent hover:border-indigo-100 transition-all">
                <input type="radio" name="exportMode" value="multiple" ${state.outputMode === 'multiple' ? 'checked' : ''} class="w-5 h-5 text-indigo-600">
                <span class="text-sm font-bold text-slate-700">Tách từng tệp rời (.docx)</span>
            </label>
            <div class="space-y-2 mt-4">
                <button id="btnChooseFolder" class="w-full h-12 bg-slate-800 text-white font-bold rounded-xl hover:bg-slate-900 transition-all">Chọn thư mục lưu</button>
                <div id="exportFolderLabel" class="text-[12px] text-slate-500">${state.exportFolderLabel ? `Thư mục đã chọn: ${state.exportFolderLabel}` : 'Chưa chọn thư mục lưu. Nếu không chọn, hệ thống sẽ yêu cầu lưu tệp.'}</div>
                <div class="text-[11px] text-slate-400">Lưu ý: Chức năng chọn thư mục chỉ hoạt động nếu trình duyệt/hệ thống hỗ trợ API File System Access.</div>
            </div>
        </div>
    `;
    
    const soHDInput = document.createElement('div');
    soHDInput.className = 'mt-6 pt-4 border-t border-slate-200';
    soHDInput.innerHTML = `
        <label class="text-[10px] font-black text-slate-400 uppercase mb-2 ml-1">Số hợp đồng (cho đặt tên file)</label>
        <input type="text" id="soHDInput" spellcheck="false" value="${state.soHDForExport || ''}" placeholder="VD: HD-2024-001" class="input-field w-full">
    `;
    container.appendChild(soHDInput);

    const radioInputs = container.querySelectorAll('input[type="radio"]');
    radioInputs.forEach(input => {
        input.onchange = async () => { 
            state.outputMode = input.value; 
            await saveState();
        };
    });

    const soHDField = container.querySelector('#soHDInput');
    soHDField.onchange = async () => { 
        state.soHDForExport = soHDField.value; 
        await saveState();
    };

    const btnChooseFolder = container.querySelector('#btnChooseFolder');
    const exportFolderLabel = container.querySelector('#exportFolderLabel');
    if (btnChooseFolder) {
        btnChooseFolder.onclick = async () => {
            if (typeof window.showDirectoryPicker !== 'function') {
                showToast('Trình duyệt không hỗ trợ chức năng chọn thư mục trực tiếp.', 'error');
                return;
            }

            try {
                const folderHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
                state.exportFolderHandle = folderHandle;
                state.exportFolderLabel = folderHandle.name || 'Thư mục được chọn';
                exportFolderLabel.innerText = `Thư mục đã chọn: ${state.exportFolderLabel}`;
                showToast('Đã chọn thư mục lưu.', 'success');
            } catch (err) {
                if (err.name !== 'AbortError') {
                    showToast('Không thể chọn thư mục: ' + err.message, 'error');
                }
            }
        };
    }
}


// --- MODAL & CRUD ---
function openEditModal(index = -1) {
    state.editingIndex = index;
    const type = state.currentTab;
    const config = categories[type];
    const modalForm = document.getElementById('modalFormContent');
    modalForm.innerHTML = "";
    
    document.getElementById('modalTitle').innerText = (index === -1 ? "Thêm mới " : "Chỉnh sửa ") + config.title;
    
    const existingData = index === -1 ? [] : state[type][index];
    
    config.fields.forEach((field, i) => {
        if (field === 'stt') return; // Auto-generated
        if (field === 'status' && type === 'mayMoc') return; // Auto-calculated
        
        const div = document.createElement("div");
        const inputHtml = (field === 'qty') 
            ? `<input type="number" min="0" id="modalInput_${field}" value="${existingData[i] || ''}" class="input-field w-full">`
            : `<textarea spellcheck="false" id="modalInput_${field}" class="input-field resize-y py-2 w-full" style="height: auto; min-height: 2.5rem;" rows="1">${existingData[i] || ''}</textarea>`;

        div.innerHTML = `
            <label class="text-[10px] font-black text-slate-400 uppercase mb-1 ml-1">${config.labels[i]}</label>
            ${inputHtml}
        `;
        modalForm.appendChild(div);
    });
    
    document.getElementById('modalOverlay').classList.remove('hidden');
    
    // Tự động căn chỉnh độ cao các textarea trong bảng chi tiết
    setTimeout(() => {
        document.querySelectorAll('#modalFormContent textarea').forEach(ta => {
            ta.style.height = 'auto';
            ta.style.height = ta.scrollHeight + 'px';
        });
    }, 50);
}

async function saveModal() {
    const type = state.currentTab;
    const config = categories[type];
    const newEntry = [];
    
    // Auto-calculate STT
    newEntry[0] = state.editingIndex === -1 ? (state[type].length + 1).toString() : state[type][state.editingIndex][0];
    
    config.fields.forEach((field, i) => {
        if (field === 'stt') return;
        let val = document.getElementById(`modalInput_${field}`)?.value || "";
        if (field === 'qty' && val) val = String(val).padStart(2, '0');
        if (type === 'nhanSu' && field === 'phone' && val) {
            val = val.replace(/(?:0|\+84)[\d\.\-\s]{9,13}/g, match => {
                let d = match.replace(/\D/g, '');
                if (d.startsWith('84')) d = '0' + d.substring(2);
                if (d.length >= 10 && d.startsWith('0')) return d.substring(0, 10).replace(/(\d{4})(\d{3})(\d{3})/, '$1.$2.$3');
                return match;
            });
            document.getElementById(`modalInput_${field}`).value = val;
        }
        newEntry[i] = val;
    });
    
    // Specialized Logic: KiemTraSoHuu (VBA Port)
    if (type === 'mayMoc') {
        const ownerName = normalize(newEntry[4]); // Chu So Huu
        const companyName = normalize(state.duAn.dvtc); // Don Vi Thi Cong
        newEntry[5] = (ownerName.includes(companyName) || companyName.includes(ownerName)) ? "Sở hữu" : "Đi thuê";
    }
    
    if (state.editingIndex === -1) {
        state[type].push(newEntry);
    } else {
        state[type][state.editingIndex] = newEntry;
    }
    
    await saveState();
    closeModal();
    renderContent();
    lucide.createIcons();
}

function closeModal() {
    document.getElementById('modalOverlay').classList.add('hidden');
}

// --- CORE FUNCTIONS ---
async function syncDataToWord() {
    updateLog("Đang đồng bộ dữ liệu vào văn bản...");

    // 1. Cập nhật Biến văn bản (DocVariables) như VBA và tag cụ thể
    const docVars = {
        "DuAn": state.duAn.tenDuAn,
        "GoiThau": state.duAn.goiThau,
        "DVTC": state.duAn.dvtc,
        "DaiDienCDT": state.duAn.daiDienCDT,
        "TVGS": state.duAn.tvgs,
        "NgayKhoiCong": state.duAn.ngayKhoiCong,
        "NgayHoanThanh": state.duAn.ngayHoanThanh
    };
    
    try {
        // Thử cập nhật Content Controls trước
        await WordService.updateDocVariables(docVars);
    } catch (e) {
        updateLog("Content Controls không khả dụng: " + e.message);
    }

    // Luôn cập nhật Document Variables (DOCVARIABLE) để phù hợp template dùng DOCVARIABLE
    try {
        await WordService.updateDocumentVariables(docVars);
    } catch (e) {
        updateLog("Document Variables thất bại: " + e.message);
    }

    // Thay placeholder chung (ưu tiên) để không phụ thuộc hoàn toàn DOCVARIABLE.
    try {
        await WordService.replaceInDocument("<<DuAn>>", state.duAn.tenDuAn || "", "DuAn");
        await WordService.replaceInDocument("<<GoiThau>>", state.duAn.goiThau || "", "GoiThau");
        await WordService.replaceInDocument("<<DVTC>>", state.duAn.dvtc || "", "DVTC");
        await WordService.replaceInDocument("<<DaiDienCDT>>", state.duAn.daiDienCDT || "", "DaiDienCDT");
        await WordService.replaceInDocument("<<TVGS>>", state.duAn.tvgs || "", "TVGS");
        await WordService.replaceInDocument("<<NgayKhoiCong>>", state.duAn.ngayKhoiCong || "", "NgayKhoiCong");
        await WordService.replaceInDocument("<<NgayHoanThanh>>", state.duAn.ngayHoanThanh || "", "NgayHoanThanh");
    } catch (err) {
        updateLog("Không thể thay placeholder: " + err.message);
    }

    // Cập nhật các field trong document để DOCVARIABLE hiển thị giá trị mới.
    try {
        await WordService.updateAllFields();
    } catch (fieldError) {
        updateLog("Không thể update các field: " + fieldError.message);
    }
    
    await WordService.replaceTag("bmTenDuAn", state.duAn.tenDuAn || "---");
    
    // 2. Cập nhật Bảng (Table Syncs), ưu tiên Bookmark nếu có
    await WordService.xuatBang(state.nhanSu, "Họ và tên", "bmNhanSu", updateLog);
    await WordService.xuatBang(state.nhanSu, "Họ và tên", "bmNhanSu2", updateLog);
    await WordService.xuatBang(state.nhanSu, "Họ và tên", "bmNhanSu3", updateLog);
    await WordService.xuatBang(state.mayMoc, "Tên thiết bị|Xe máy|Máy móc|Thiết bị", "bmMayMoc", updateLog);
    await WordService.xuatBang(state.vatLieu, "Tên vật tư", "bmVatLieu", updateLog);
    await WordService.xuatBang(state.thiNghiem, "Đơn vị thí nghiệm", "bmThiNghiem", updateLog);
    
    // Format căn lề bảng
    await WordService.applyTableAlignment();
    
    await WordService.applyModernStyleToDocument();
}

async function onCapNhatClick() {
    try {
        updateLog("--- Bắt đầu Cập nhật dữ liệu ---");
        await syncDataToWord();
        updateLog("✓ Hoàn tất cập nhật dữ liệu vào văn bản.");
        
        // Tự động ghi đè lại file đã xuất nếu có
        if (state.exportFolderHandle && (state.hasExportedMaster || state.hasSplitFiles)) {
            updateLog("Bắt đầu cập nhật thay thế các file DOCX bên ngoài thư mục...");
            if (state.hasExportedMaster) {
                await WordService.processExport('master', state.duAn.tenDuAn, {
                    folderHandle: state.exportFolderHandle,
                    outputMode: state.outputMode
                });
                updateLog("✓ Đã cập nhật ghi đè file tổng.");
            }
            if (state.hasSplitFiles) {
                await WordService.processExport('split', state.duAn.tenDuAn, {
                    folderHandle: state.exportFolderHandle,
                    outputMode: state.outputMode
                });
                updateLog("✓ Đã cập nhật ghi đè file tách.");
            }
            msg = "Đã cập nhật Word & ghi đè toàn bộ file đã xuất!";
        }
        
        updateLog("✅ Đã CẬP NHẬT HOÀN TẤT!");
        showToast("Đã cập nhật dữ liệu thành công!", "success");
    } catch (e) {
        console.error("onCapNhatClick Error:", e);
        const errorDetail = e.message || "Lỗi không xác định";
        updateLog(`❌ LỖI CẬP NHẬT: ${errorDetail}`);
        if (e.stack) console.log(e.stack);
        showToast("Có lỗi xảy ra khi cập nhật (Xem Nhật ký bên dưới)", "error");
    }
}

async function onImportFromDocClick() {
    try {
        updateLog("Đang quét nội dung văn bản...");
        const data = await WordService.importDataFromDoc();
        
        // --- LOG CHẨN ĐOÁN (DIAGNOSTIC) ---
        if (data.inventory) {
            console.log("Word Inventory:", data.inventory);
            let diagMsg = `--- QUÉT THẤY ${data.inventory.tags.length} TAGS & ${data.inventory.tables.length} BẢNG ---`;
            updateLog(diagMsg);
            
            // Log 3 tag đầu tiên để chẩn đoán nhanh nếu có lỗi
            data.inventory.tags.slice(0, 5).forEach(t => {
                updateLog(`[Thẻ] ${t.tag}: ${t.text.substring(0, 20)}...`);
            });

            if (data.inventory.variables && data.inventory.variables.length > 0) {
                updateLog(`--- QUÉT THẤY ${data.inventory.variables.length} BIẾN ẨN (VARS) ---`);
                data.inventory.variables.slice(0, 3).forEach(v => {
                    updateLog(`[Biến] ${v.name}: ${v.value.substring(0, 20)}...`);
                });
            }
            
            data.inventory.tables.forEach(t => {
                updateLog(`[Bảng] ${t.header.substring(0, 30)}...`);
            });
        }

        // Cập nhật state
        let updateCount = 0;
        if (Object.keys(data.duAn).length > 0) {
            state.duAn = { ...state.duAn, ...data.duAn };
            updateCount += Object.keys(data.duAn).length;
        }
        
        if (data.nhanSu.length > 0) { state.nhanSu = data.nhanSu; updateCount++; }
        if (data.mayMoc.length > 0) { state.mayMoc = data.mayMoc; updateCount++; }
        if (data.vatLieu.length > 0) { state.vatLieu = data.vatLieu; updateCount++; }
        if (data.thiNghiem.length > 0) { state.thiNghiem = data.thiNghiem; updateCount++; }

        await saveState();
        renderContent();
        
        updateLog(`✓ Đã nhập thành công ${updateCount} nhóm dữ liệu!`);
        showToast("Đã khôi phục dữ liệu từ văn bản!", "success");
    } catch (e) {
        updateLog("Lỗi nhập liệu: " + e.message);
        showToast("Không thể nhập dữ liệu từ văn bản", "error");
    }
}

async function requestExportFolder() {
    if (state.exportFolderHandle) return true;
    if (typeof window.showDirectoryPicker !== 'function') {
        updateLog('⚠️ Trình duyệt không hỗ trợ chọn thư mục trực tiếp. Hệ thống sẽ lưu file vào thư mục Download.');
        return false;
    }

    try {
        const folderHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
        if (folderHandle) {
            state.exportFolderHandle = folderHandle;
            state.exportFolderLabel = folderHandle.name || 'Thư mục được chọn';
            // Lưu handle vào IndexedDB để dùng cho lần sau
            await StorageService.saveFolderHandle(folderHandle);
            
            const labelElement = document.getElementById('exportFolderLabel');
            if (labelElement) labelElement.innerText = `Thư mục đã chọn: ${state.exportFolderLabel}`;
            updateLog('Đã chọn thư mục lưu.');
            return true;
        }
    } catch (err) {
        if (err.name !== 'AbortError') {
            updateLog('Không thể chọn thư mục: ' + err.message);
        }
        return false;
    }

    return false;
}

async function onTachClick() {
    try {
        await syncDataToWord();
        updateLog("⏳ Đang chuẩn bị tách hồ sơ theo bookmark...");
        
        // Capture console logs để debug
        const originalLog = console.log;
        const originalWarn = console.warn;
        const originalError = console.error;
        const logs = [];
        
        console.log = (...args) => {
            const msg = args.join(" ");
            logs.push("[LOG] " + msg);
            originalLog(...args);
        };
        console.warn = (...args) => {
            const msg = args.join(" ");
            logs.push("[WARN] " + msg);
            originalWarn(...args);
        };
        console.error = (...args) => {
            const msg = args.join(" ");
            logs.push("[ERROR] " + msg);
            originalError(...args);
        };
        
        await requestExportFolder();
        await WordService.processExport('split', state.duAn.tenDuAn, {
            folderHandle: state.exportFolderHandle,
            outputMode: state.outputMode
        });
        state.hasSplitFiles = true;
        await saveState();
        
        console.log = originalLog;
        console.warn = originalWarn;
        console.error = originalError;
        
        // Phân tích logs để detect vấn đề
        const allLogs = logs.join("\n");
        
        // Check for critical issues
        if (allLogs.includes("ERROR") && 
            !allLogs.includes("getSlice") && 
            !allLogs.includes("Split thành công")) {
            updateLog("❌ LỖI CHI TIẾT:\n" + allLogs);
            showToast("Có lỗi xảy ra trong quá trình tách hồ sơ", "error");
            return;
        }
        
        if (allLogs.includes("ZIP file created: 0 bytes")) {
            updateLog("❌ ZIP file trống (0 bytes)!\n\n" + allLogs); 
            showToast("ZIP file trống - kiểm tra logs", "error");
            return;
        }
        
        // Check if "Found bookmark" appears (new logging format)
        const hasFoundBookmarks = allLogs.includes("✓ Found bookmark");
        const hasSplitSuccess = allLogs.includes("✅ Split thành công");
        
        if (!hasFoundBookmarks && !hasSplitSuccess) {
            updateLog("⚠️ CẢNH BÁO: Không thể truy cập bookmark!\n\n" +
                     "Lỗi có thể là:\n" +
                     "- Word API không hỗ trợ phương thức truy cập bookmark\n" +
                     "- Phiên bản Word quá cũ\n" +
                     "- Add-in không có quyền truy cập bookmarks\n\n" +
                     "Chi tiết:\n" + allLogs);
            showToast("Không thể truy cập bookmark - kiểm tra logs", "warning");
            return;
        }
        
        // Hiển thị logs chi tiết
        updateLog("📊 Chi tiết quá trình tách:\n" + allLogs);
        updateLog("\n✅ Đã TÁCH HỒ SƠ thành công!");
        showToast("Đã tách hồ sơ thành công!", "success");
    } catch (e) {
        updateLog("❌ Lỗi exception: " + e.message + "\n" + (e.stack || ""));
        showToast("Có lỗi xảy ra: " + e.message, "error");
    }
}

async function onXuatClick() {
    try {
        await syncDataToWord();
        updateLog("Đang xuất bộ hồ sơ tổng...");
        await requestExportFolder();
        await WordService.processExport('master', state.duAn.tenDuAn, {
            folderHandle: state.exportFolderHandle,
            outputMode: state.outputMode
        });
        state.hasExportedMaster = true;
        await saveState();
        updateLog("Đã XUẤT HỒ SƠ TỔNG (.docx) thành công!");
        showToast("Đã xuất hồ sơ tổng thành công!", "success");
    } catch (e) {
        updateLog("Lỗi: " + e.message + (e.debugInfo ? "\nDebug: " + JSON.stringify(e.debugInfo) : ""));
        showToast("Có lỗi xảy ra khi xuất hồ sơ", "error");
    }
}

// --- UTILS ---
function registerEvents() {
    document.querySelectorAll('[data-tab]').forEach(btn => {
        btn.onclick = () => switchTab(btn.dataset.tab);
    });
    
    document.getElementById('btnModalCancel').onclick = closeModal;
    document.getElementById('btnModalSave').onclick = saveModal;
    
    // Nút chức năng Footer (IDs đã được cập nhật bản 1.0.0.20)
    document.getElementById('btnCapNhat').onclick = onCapNhatClick;
    document.getElementById('btnSplit').onclick = onTachClick;
    document.getElementById('btnExport').onclick = onXuatClick;
    document.getElementById('btnImportDoc').onclick = onImportFromDocClick;
    
    document.getElementById('btnResetFooter').onclick = async () => {
        const confirmed = await openResetDataModal();
        if (!confirmed) {
            updateLog("Đã hủy thao tác khởi tạo lại dữ liệu.");
            return;
        }

        updateLog("Đang xóa trắng và chuẩn bị các ví dụ mờ...");
        try {
            // Thay vì nạp MockData vào state, chúng ta reset state về rỗng
            // Điều này sẽ làm cho các Placeholder (chữ mờ) hiện lên
            state.duAn = {
                tenDuAn: "", goiThau: "", dvtc: "", daiDienCDT: "", tvgs: "", ngayKhoiCong: "", ngayHoanThanh: ""
            };
            state.nhanSu = [];
            state.mayMoc = [];
            state.vatLieu = [];
            state.thiNghiem = [];
            
            state.soHDForExport = "";
            state.hasExportedMaster = false;
            state.hasSplitFiles = false;
            
            await saveState();
            renderContent();
            lucide.createIcons();
            showToast("Đã làm mới dữ liệu và hiển thị ghi chú mờ!", "success");
            updateLog("✓ Đã hoàn thành khởi tạo lại giao diện mẫu.");
        } catch (err) {
            showToast("Lỗi khi xóa dữ liệu: " + err.message, "error");
        }
    };

    // Tự động thay đổi chiều cao textarea khi nhập liệu
    document.addEventListener('input', function(e) {
        if (e.target.tagName && e.target.tagName.toLowerCase() === 'textarea') {
            e.target.style.height = 'auto';
            const newHeight = Math.min(e.target.scrollHeight, 150);
            e.target.style.height = newHeight + 'px';
            e.target.style.overflowY = e.target.scrollHeight > 150 ? 'auto' : 'hidden';
        }
    });
}

function normalize(s) { 
    if (!s) return "";
    return s.toLowerCase()
        .replace(/[àáạảãâầấậẩẫăằắặẳẵ]/g, 'a')
        .replace(/[èéẹẻẽêềếệểễ]/g, 'e')
        .replace(/[ìíịỉĩ]/g, 'i')
        .replace(/[òóọỏõôồốộổỗơờớợởỡ]/g, 'o')
        .replace(/[ùúụủũưừứựửữ]/g, 'u')
        .replace(/[ỳýỵỷỹ]/g, 'y')
        .replace(/đ/g, 'd')
        .replace(/[^a-z0-9\s]/g, ' ')
        .replace(/\bcp\b/g, 'co phan')
        .replace(/\btnhh\b/g, 'trach nhiem huu han')
        .replace(/\bcty\b/g, 'cong ty')
        .replace(/\btm\b/g, 'thuong mai')
        .replace(/\bdv\b/g, 'dich vu')
        .replace(/\bxd\b/g, 'xay dung')
        .replace(/\bjsc\b/g, 'co phan')
        .replace(/\s+/g, ' ')
        .trim(); 
}

function sanitizeFileName(input) {
    if (!input) return "";
    return input.replace(/[\\/:*?"<>|]/g, "").trim();
}

function generateProjectFolderName() {
    const duAn = state.duAn || {};
    const tenDuAn = sanitizeFileName(duAn.tenDuAn);
    const soHD = sanitizeFileName(state.soHDForExport);
    const dvtc = sanitizeFileName(duAn.dvtc);

    if (tenDuAn && soHD) {
        return `${tenDuAn}_${soHD}`;
    }

    if (tenDuAn) {
        return tenDuAn;
    }

    if (soHD) {
        return soHD;
    }

    if (dvtc) {
        return dvtc;
    }

    return "";
}

function openProjectNameModal(defaultName = "") {
    return new Promise((resolve) => {
        const modal = document.getElementById('projectNameModal');
        const input = document.getElementById('projectNameInput');
        const error = document.getElementById('projectNameError');
        const btnConfirm = document.getElementById('projectNameConfirm');
        const btnCancel = document.getElementById('projectNameCancel');

        input.value = defaultName;
        error.classList.add('hidden');

        const closeModal = () => {
            modal.classList.add('hidden');
            btnConfirm.onclick = null;
            btnCancel.onclick = null;
        };

        btnConfirm.onclick = () => {
            const value = input.value.trim();
            if (!value) {
                error.classList.remove('hidden');
                return;
            }
            closeModal();
            resolve(value);
        };

        btnCancel.onclick = () => {
            closeModal();
            resolve(null);
        };

        modal.classList.remove('hidden');
        input.focus();
    });
}

async function requestProjectName(defaultName = state.duAn.tenDuAn || "") {
    const autoName = generateProjectFolderName();
    if (autoName) {
        return autoName;
    }

    const projectName = await openProjectNameModal(defaultName);
    return projectName;
}

function openResetDataModal() {
    return new Promise((resolve) => {
        const modal = document.getElementById('resetDataModal');
        const btnConfirm = document.getElementById('resetDataConfirm');
        const btnCancel = document.getElementById('resetDataCancel');

        const closeModal = () => {
            modal.classList.add('hidden');
            btnConfirm.onclick = null;
            btnCancel.onclick = null;
        };

        btnConfirm.onclick = () => {
            closeModal();
            resolve(true);
        };

        btnCancel.onclick = () => {
            closeModal();
            resolve(false);
        };

        modal.classList.remove('hidden');
        lucide.createIcons();
    });
}

function updateLog(m) { document.getElementById('logMsg').innerText = m; }

function showToast(message, type = 'success') {
    const bgClass = type === 'error' ? 'bg-red-600' : 'bg-emerald-600';
    const toast = document.createElement('div');
    toast.className = `fixed bottom-4 left-1/2 -translate-x-1/2 ${bgClass} text-white px-4 py-2 rounded-xl shadow-xl shadow-slate-200/50 text-[12px] font-bold z-[100] transition-all duration-300`;
    toast.style.transform = 'translate(-50%, 20px)';
    toast.style.opacity = '0';
    toast.innerText = message;
    document.body.appendChild(toast);
    
    // Animate in
    setTimeout(() => {
        toast.style.transform = 'translate(-50%, 0)';
        toast.style.opacity = '1';
    }, 10);
    
    // Animate out
    setTimeout(() => {
        toast.style.transform = 'translate(-50%, 20px)';
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}
