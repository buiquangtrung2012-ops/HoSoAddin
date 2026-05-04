import { WordService } from './word_service.js?v=04052026.1228';
import { StorageService } from './storage_service.js?v=04052026.1228';
import { MockData } from './mock_data.js?v=04052026.1228';

/* global Office, lucide */

// --- GLOBAL STATE ---
let state = {
    currentTab: 'duAn',
    editingIndex: -1,
    duAn: {
        soHD: "",
        tenDuAn: "",
        goiThau: "",
        dvtc: "",
        daiDienCDT: "",
        tvgs: "",
        ngayKhoiCong: "",
        ngayHoanThanh: "",
        isLienDanh: false,
        dvtcMembers: [],
        sigFontSize: 11
    },
    soHDForExport: "",
    exportFolderHandle: null,
    exportFolderLabel: "",
    nhanSu: [],
    mayMoc: [],
    vatLieu: [],
    thiNghiem: [],
    outputMode: 'multiple',
    autoSplitOnUpdate: false,
    useProjectNameFolder: true
};

// --- CONFIGURATION ---
const categories = {
    duAn: { title: "Dự án", fields: ["tenDuAn", "goiThau", "dvtc", "daiDienCDT", "tvgs", "soHD", "ngayKhoiCong", "ngayHoanThanh"], labels: ["Tên dự án", "Tên gói thầu", "Đơn vị thi công", "Đại diện CDT", "Tư vấn giám sát", "Số hợp đồng", "Ngày khởi công", "Ngày hoàn thành"] },
    nhanSu: { title: "Nhân sự", fields: ["stt", "name", "role", "major", "phone"], labels: ["STT", "Họ và tên", "Chức danh", "Chuyên ngành", "Số điện thoại"] },
    mayMoc: { title: "Máy móc", fields: ["stt", "name", "unit", "qty", "owner", "status"], labels: ["STT", "Tên thiết bị", "Đơn vị tính", "Số lượng", "Chủ sở hữu", "Hình thức"] },
    vatLieu: { title: "Vật liệu", fields: ["stt", "name", "standard", "origin", "note"], labels: ["STT", "Tên vật tư", "Thông số/Tiêu chuẩn", "Nguồn gốc", "Đơn vị cung cấp"] },
    thiNghiem: { title: "Phòng TN", fields: ["stt", "dvtn", "address", "ptn", "func"], labels: ["STT", "Đơn vị TN", "Địa chỉ", "Tên phòng TN", "Chức năng"] },
    xuatBan: { title: "Cài đặt", fields: [], labels: [] },
    mauHoSo: { title: "Tạo mẫu", fields: [], labels: [] }
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
        state.autoSplitOnUpdate = syncState.autoSplitOnUpdate ?? false;
        state.useProjectNameFolder = syncState.useProjectNameFolder ?? true;
        state.soHDForExport = syncState.soHDForExport || "";

        // Migration: Tự động chuyển số hợp đồng từ cấu hình tập tin sang thông tin dự án nếu chưa có
        if (!state.duAn.soHD && state.soHDForExport) {
            state.duAn.soHD = state.soHDForExport;
        }
    }

    // Migration: Chuyển đổi dvtcMembers từ String (cũ) sang mảng Array
    if (typeof state.duAn.isLienDanh === 'undefined') state.duAn.isLienDanh = false;
    if (typeof state.duAn.dvtcMembers === 'string') {
        state.duAn.dvtcMembers = state.duAn.dvtcMembers.split('\n').map(m => m.trim()).filter(Boolean);
    }
    if (!Array.isArray(state.duAn.dvtcMembers)) state.duAn.dvtcMembers = [];
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
        autoSplitOnUpdate: state.autoSplitOnUpdate,
        useProjectNameFolder: state.useProjectNameFolder,
        soHDForExport: state.duAn.soHD // Sync back to the old field for compatibility
    });
}

// --- UI CONTROLLER ---
function switchTab(tabId) {
    state.currentTab = tabId;
    document.querySelectorAll('[data-tab]').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });

    // Update Page Title
    const pageTitle = document.getElementById('pageTitle');
    if (pageTitle) {
        pageTitle.innerText = categories[tabId]?.title?.toUpperCase() || tabId.toUpperCase();
    }

    renderContent();
    lucide.createIcons();
}

function renderContent() {
    try {
        const container = document.getElementById('tabContent');
        if (!container) return;
        container.innerHTML = "";

        if (state.currentTab === 'duAn') {
            renderProjectView(container);
        } else if (state.currentTab === 'xuatBan') {
            renderExportSettings(container);
        } else if (state.currentTab === 'mauHoSo') {
            renderTemplateCreator(container);
        } else {
            renderList(container, state.currentTab);
        }

        // Tự động căn chỉnh lại độ cao cho các textarea
        adjustAllTextareaHeights();
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

function renderProjectView(container) {
    const config = categories.duAn;

    const wrapper = document.createElement("div");
    wrapper.className = "max-w-3xl mx-auto space-y-4 pb-12";

    // --- NEW: General Signature Settings (Always Visible) ---
    const sigSettingsCard = document.createElement("div");
    sigSettingsCard.className = "px-5 py-4 bg-white rounded-2xl border border-slate-100 shadow-sm space-y-3";
    sigSettingsCard.innerHTML = `
        <div class="flex items-center justify-between">
            <div class="flex items-center gap-3">
                <div class="w-8 h-8 bg-indigo-50 text-indigo-600 rounded-lg flex items-center justify-center">
                    <i data-lucide="type" size="18"></i>
                </div>
                <div>
                    <h4 class="text-xs font-bold text-slate-700">Cấu hình Bảng ký tên</h4>
                    <p class="text-[10px] text-slate-500">Cỡ chữ hiển thị của tên các công ty</p>
                </div>
            </div>
            <div class="flex items-center gap-2">
                <span class="text-[10px] font-black text-slate-400 uppercase tracking-widest">Cỡ chữ:</span>
                <select id="selSigFontSize" class="bg-slate-50 border border-slate-100 rounded-lg px-2 py-1 text-xs font-bold text-slate-700 outline-none focus:border-indigo-300 transition-all">
                    <option value="8" ${state.duAn.sigFontSize == 8 ? 'selected' : ''}>8</option>
                    <option value="9" ${state.duAn.sigFontSize == 9 ? 'selected' : ''}>9</option>
                    <option value="10" ${state.duAn.sigFontSize == 10 ? 'selected' : ''}>10</option>
                    <option value="11" ${state.duAn.sigFontSize == 11 ? 'selected' : ''}>11</option>
                    <option value="12" ${state.duAn.sigFontSize == 12 ? 'selected' : ''}>12</option>
                    <option value="13" ${state.duAn.sigFontSize == 13 ? 'selected' : ''}>13</option>
                    <option value="14" ${state.duAn.sigFontSize == 14 ? 'selected' : ''}>14</option>
                </select>
            </div>
        </div>
    `;
    wrapper.appendChild(sigSettingsCard);

    const selFontSize = sigSettingsCard.querySelector('#selSigFontSize');
    if (selFontSize) {
        selFontSize.onchange = async () => {
            state.duAn.sigFontSize = parseInt(selFontSize.value, 10);
            await saveState();
        };
    }

    const fieldIcons = {
        tenDuAn: 'file-text',
        goiThau: 'clipboard-list',
        dvtc: 'briefcase',
        daiDienCDT: 'user',
        tvgs: 'users',
        soHD: 'hash',
        ngayKhoiCong: 'calendar',
        ngayHoanThanh: 'calendar-check'
    };

    // Create a grid for dates if they are next to each other
    let dateGrid = null;

    config.fields.forEach((field, i) => {
        const isDate = field === 'ngayKhoiCong' || field === 'ngayHoanThanh';
        const value = state.duAn[field] || "";
        const mockVal = MockData.duAn[field] || config.labels[i];

        const card = document.createElement("div");
        card.className = "info-card group p-5 cursor-default";

        card.innerHTML = `
            <div class="flex-1 min-w-0">
                <p class="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1.5">${config.labels[i]}</p>
                ${isDate ?
                `<input type="text" data-field="${field}" spellcheck="false" 
                        class="project-input w-full bg-transparent border-none outline-none font-bold text-slate-700 text-[14px] p-0 m-0 date-picker-input" 
                        placeholder="VD: ${mockVal}" value="${value}">` :
                `<textarea data-field="${field}" spellcheck="false" 
                        class="project-input w-full bg-transparent border-none outline-none font-bold text-slate-700 text-[14px] resize-none overflow-hidden p-0 m-0" 
                        placeholder="VD: ${mockVal}" rows="1">${value}</textarea>`
            }
            </div>
        `;

        if (isDate) {
            if (!dateGrid) {
                dateGrid = document.createElement("div");
                dateGrid.className = "grid grid-cols-2 gap-4";
                wrapper.appendChild(dateGrid);
            }
            dateGrid.appendChild(card);
        } else {
            dateGrid = null;
            wrapper.appendChild(card);

            // Thêm tùy chọn Liên danh ngay sau ô Đơn vị thi công
            if (field === 'dvtc') {
                const jvToggleArea = document.createElement("div");
                jvToggleArea.className = "px-5 py-4 bg-white rounded-2xl border border-slate-100 shadow-sm space-y-4 transition-all";
                jvToggleArea.innerHTML = `
                    <div class="flex items-center justify-between">
                        <div class="flex items-center gap-3">
                            <div>
                                <h4 class="text-xs font-bold text-slate-700">Chế độ Liên danh</h4>
                                <p class="text-[10px] text-slate-500">Sử dụng bảng ký tên 3 cột cho nhiều thành viên</p>
                            </div>
                        </div>
                        <label class="relative inline-flex items-center cursor-pointer">
                            <input type="checkbox" id="chkIsLienDanh" class="sr-only peer" ${state.duAn.isLienDanh ? 'checked' : ''}>
                            <div class="w-11 h-6 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-indigo-600"></div>
                        </label>
                    </div>
                    <div id="jvMembersArea" class="${state.duAn.isLienDanh ? '' : 'hidden'} pt-2 space-y-4 border-t border-slate-50">
                        <p class="text-[9px] font-black text-slate-400 uppercase tracking-widest">Danh sách thành viên ký tên</p>
                        <div id="membersListContainer" class="space-y-3"></div>
                        <button id="btnAddMember" class="flex items-center gap-2 px-3 py-2 bg-indigo-50 text-indigo-600 rounded-lg text-[11px] font-bold hover:bg-indigo-100 transition-all">
                            <i data-lucide="plus-circle" size="14"></i>
                            THÊM THÀNH VIÊN
                        </button>
                        <button id="btnUpdateSignatureTable" class="w-full flex items-center justify-center gap-2 px-3 py-2.5 bg-green-50 text-green-700 rounded-lg text-[11px] font-bold hover:bg-green-100 transition-all border border-green-200">
                            <i data-lucide="refresh-cw" size="14"></i>
                            CẬP NHẬT BẢNG KÝ TÊN
                        </button>
                    </div>
                `;
                wrapper.appendChild(jvToggleArea);

                const chk = jvToggleArea.querySelector('#chkIsLienDanh');
                const area = jvToggleArea.querySelector('#jvMembersArea');
                const listContainer = jvToggleArea.querySelector('#membersListContainer');
                const btnAdd = jvToggleArea.querySelector('#btnAddMember');

                const renderMembers = () => {
                    listContainer.innerHTML = "";
                    if (state.duAn.dvtcMembers.length === 0) {
                        listContainer.innerHTML = `<p class="text-[10px] text-slate-400 italic">Chưa có thành viên nào. Nhấn nút Thêm bên dưới.</p>`;
                    }
                    state.duAn.dvtcMembers.forEach((m, idx) => {
                        const row = document.createElement("div");
                        row.className = "flex items-center gap-2 group";
                        row.innerHTML = `
                            <div class="flex-1 relative">
                                <input type="text" value="${m}" placeholder="Tên thành viên ${idx + 1}"
                                    class="member-input w-full bg-slate-50 border border-slate-100 rounded-xl px-4 py-2.5 text-xs font-bold text-slate-700 outline-none focus:border-indigo-300 focus:bg-white transition-all">
                            </div>
                            <button class="btn-remove-member p-2 text-slate-300 hover:text-red-500 transition-colors">
                                <i data-lucide="x-circle" size="18"></i>
                            </button>
                        `;

                        const input = row.querySelector('.member-input');
                        input.oninput = async () => {
                            state.duAn.dvtcMembers[idx] = input.value.trim();
                            // Không saveState liên tục để tránh lag, chỉ cập nhật memory state
                        };
                        input.onblur = async () => {
                            await saveState();
                        };

                        row.querySelector('.btn-remove-member').onclick = async () => {
                            state.duAn.dvtcMembers.splice(idx, 1);
                            await saveState();
                            renderMembers();
                        };

                        listContainer.appendChild(row);
                    });
                    lucide.createIcons();
                };

                chk.onchange = async () => {
                    state.duAn.isLienDanh = chk.checked;
                    area.classList.toggle('hidden', !chk.checked);
                    await saveState();
                    if (chk.checked && state.duAn.dvtcMembers.length === 0) {
                        state.duAn.dvtcMembers.push("");
                        renderMembers();
                    }
                };

                btnAdd.onclick = async () => {
                    state.duAn.dvtcMembers.push("");
                    await saveState();
                    renderMembers();
                    const inputs = listContainer.querySelectorAll('.member-input');
                    if (inputs.length > 0) inputs[inputs.length - 1].focus();
                };

                // Nút cập nhật bảng ký tên ngay lập tức
                const btnUpdateSig = jvToggleArea.querySelector('#btnUpdateSignatureTable');
                if (btnUpdateSig) {
                    btnUpdateSig.onclick = async () => {
                        try {
                            btnUpdateSig.disabled = true;
                            btnUpdateSig.innerHTML = `<i data-lucide="loader" size="14"></i> ĐANG CẬP NHẬT...`;
                            lucide.createIcons();

                            const membersList = state.duAn.isLienDanh
                                ? (Array.isArray(state.duAn.dvtcMembers) ? state.duAn.dvtcMembers : [])
                                : [];

                            await WordService.updateSignatureTable(
                                state.duAn.isLienDanh,
                                membersList,
                                state.duAn.dvtc,
                                "bmKyLienDanh",
                                state.duAn.sigFontSize || 11,
                                (msg, percent) => updateLog(msg, percent)
                            );

                            showToast("✓ Đã cập nhật xong bảng ký tên!", "success");
                            btnUpdateSig.innerHTML = `<i data-lucide="check-circle" size="14"></i> CẬP NHẬT BẢNG KÝ TÊN`;
                            btnUpdateSig.disabled = false;
                            lucide.createIcons();
                        } catch (e) {
                            console.error("Update signature table error:", e);
                            showToast(`❌ Lỗi: ${e.message}`, "error");
                            btnUpdateSig.innerHTML = `<i data-lucide="alert-circle" size="14"></i> CẬP NHẬT BẢNG KÝ TÊN`;
                            btnUpdateSig.disabled = false;
                            lucide.createIcons();
                        }
                    };

                    if (state.duAn.isLienDanh) renderMembers();
                }
            }
        }
    });

    container.appendChild(wrapper);

    // Initialization
    setTimeout(() => {
        container.querySelectorAll('.project-input').forEach(input => {
            if (input.tagName.toLowerCase() === 'textarea') adjustTextareaHeight(input);
            input.addEventListener('input', () => {
                if (input.tagName.toLowerCase() === 'textarea') adjustTextareaHeight(input);
            });
            input.onchange = async () => {
                state.duAn[input.dataset.field] = input.value;
                await saveState();
            };
        });

        if (typeof flatpickr !== 'undefined') {
            flatpickr(".date-picker-input", { dateFormat: "d/m/Y", locale: "vn", allowInput: true });
        }
        lucide.createIcons();
    }, 50);
}

function openProjectEditModal(focusField) {
    const config = categories.duAn;
    const modalForm = document.getElementById('modalFormContent');
    modalForm.innerHTML = "";

    document.getElementById('modalTitle').innerText = "Chỉnh sửa thông tin dự án";

    config.fields.forEach((field, i) => {
        const div = document.createElement("div");
        const isDate = field === 'ngayKhoiCong' || field === 'ngayHoanThanh';
        const mockVal = (MockData && MockData.duAn) ? MockData.duAn[field] : null;

        div.innerHTML = `
            <label class="text-[10px] font-black text-slate-400 uppercase mb-1 ml-1">${config.labels[i]}</label>
            <textarea id="modalInput_${field}" spellcheck="false" class="input-field project-input resize-y py-3 w-full border border-slate-100 rounded-xl text-sm focus:border-indigo-400 transition-all ${isDate ? 'date-picker-input' : ''}" 
                style="height: auto; min-height: 3rem;" rows="1" placeholder="VD: ${mockVal || ''}">${state.duAn[field] || ''}</textarea>
        `;
        modalForm.appendChild(div);
    });

    document.getElementById('modalOverlay').classList.remove('hidden');

    // Initialize tooltips and datepickers
    setTimeout(() => {
        if (typeof flatpickr !== 'undefined') {
            flatpickr(".date-picker-input", { dateFormat: "d/m/Y", locale: "vn", allowInput: true });
        }
        adjustAllTextareaHeights();
        const firstField = document.getElementById(`modalInput_${focusField}`);
        if (firstField) firstField.focus();
    }, 50);

    state.currentEditType = 'duAn';
}

function adjustTextareaHeight(ta) {
    if (!ta) return;
    ta.style.height = 'auto';
    // Nếu không có nội dung, trả về min-height để tránh placeholder làm giãn khung
    if (!ta.value) {
        ta.style.height = '';
        ta.style.overflowY = 'hidden';
        return;
    }
    const newHeight = Math.min(ta.scrollHeight, 150);
    ta.style.height = newHeight + 'px';
    ta.style.overflowY = ta.scrollHeight > 150 ? 'auto' : 'hidden';
}

function adjustAllTextareaHeights() {
    setTimeout(() => {
        document.querySelectorAll('textarea').forEach(ta => {
            adjustTextareaHeight(ta);
        });
    }, 50);
}

// Theo dõi thay đổi độ rộng cửa sổ để tính lại độ cao textarea (Responsive)
window.addEventListener('resize', () => {
    adjustAllTextareaHeights();
});

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
    if (type === 'mayMoc') wMap = { name: '24%', unit: '10%', qty: '10%', owner: '36%', status: '12%' };
    else if (type === 'nhanSu') wMap = { name: '24%', role: '23%', major: '25%', phone: '20%' };
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

            // Tự động căn chỉnh độ cao khi nhập liệu cho bảng
            if (inputElement.tagName.toLowerCase() === 'textarea') {
                inputElement.addEventListener('input', () => adjustTextareaHeight(inputElement));
            }

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

            <div class="pt-4 border-t border-slate-100">
                <h4 class="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4">Tùy chọn cập nhật</h4>
                <label class="flex items-center gap-3 p-4 bg-indigo-50/50 border border-indigo-100/50 rounded-2xl cursor-pointer hover:bg-indigo-50 transition-all">
                    <input type="checkbox" id="chkAutoSplit" ${state.autoSplitOnUpdate ? 'checked' : ''} class="w-5 h-5 text-indigo-600 rounded">
                    <div>
                        <p class="text-sm font-bold text-slate-700">Tự động tách hồ sơ</p>
                        <p class="text-[10px] text-slate-500">Tự động xuất lại các file đã tách khi nhấn "CẬP NHẬT HỒ SƠ"</p>
                    </div>
                </label>
                <label class="flex items-center gap-3 p-4 bg-slate-50 border border-slate-100 rounded-2xl cursor-pointer hover:bg-slate-100 transition-all mt-3">
                    <input type="checkbox" id="chkUseProjectName" ${state.useProjectNameFolder ? 'checked' : ''} class="w-5 h-5 text-indigo-600 rounded">
                    <div>
                        <p class="text-sm font-bold text-slate-700">Tạo thư mục "1. Ho so dau vao"</p>
                        <p class="text-[10px] text-slate-500">Đặt tên thư mục là cố định thay vì trích xuất từ tên dự án</p>
                    </div>
                </label>
            </div>
        </div>
    `;

    const chkAutoSplit = container.querySelector('#chkAutoSplit');
    if (chkAutoSplit) {
        chkAutoSplit.onchange = async () => {
            state.autoSplitOnUpdate = chkAutoSplit.checked;
            await saveState();
        };
    }

    const chkUseProjectName = container.querySelector('#chkUseProjectName');
    if (chkUseProjectName) {
        chkUseProjectName.onchange = async () => {
            state.useProjectNameFolder = chkUseProjectName.checked;
            await saveState();
        };
    }

    const radioInputs = container.querySelectorAll('input[type="radio"]');
    radioInputs.forEach(input => {
        input.onchange = async () => {
            state.outputMode = input.value;
            await saveState();
        };
    });

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

                // Cập nhật: Lưu handle vào IndexedDB để dùng cho lần sau (Bền vững)
                await StorageService.saveFolderHandle(folderHandle);

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

function renderTemplateCreator(container) {
    const wrapper = document.createElement("div");
    wrapper.className = "max-w-4xl mx-auto space-y-8 pb-12";

    const sections = [
        {
            title: "TRƯỜNG DỮ LIỆU (CONTENT CONTROLS)",
            description: "Chèn các ô chứa thông tin văn bản. Add-in sẽ tự động điền giá trị vào các ô này.",
            items: [
                { label: "Tên Dự án", tag: "DuAn", icon: "file-text" },
                { label: "Số Hợp đồng", tag: "SoHD", icon: "hash" },
                { label: "Tên Gói thầu", tag: "GoiThau", icon: "clipboard-list" },
                { label: "Đơn vị Thi công", tag: "DVTC", icon: "briefcase" },
                { label: "Đại diện CDT", tag: "DaiDienCDT", icon: "user" },
                { label: "Tư vấn Giám sát", tag: "TVGS", icon: "users" },
                { label: "Ngày Khởi công", tag: "NgayKhoiCong", icon: "calendar" },
                { label: "Ngày Hoàn thành", tag: "NgayHoanThanh", icon: "calendar-check" }
            ]
        },
        {
            title: "VÙNG DỮ LIỆU BẢNG (BOOKMARKS)",
            description: "Chèn các điểm đánh dấu cho bảng. Add-in sẽ chèn danh sách dữ liệu vào vị trí này.",
            items: [
                { label: "Bảng Nhân sự", tag: "bmNhanSu", icon: "users" },
                { label: "Bảng Máy móc", tag: "bmMayMoc", icon: "truck" },
                { label: "Bảng Vật liệu", tag: "bmVatLieu", icon: "package" },
                { label: "Bảng Phòng TN", tag: "bmThiNghiem", icon: "microscope" },
                { label: "Bảng Ký tên", tag: "bmKyLienDanh", icon: "pen-tool" }
            ]
        },
        {
            title: "DẤU MỐC TÁCH FILE (SPLIT MARKERS)",
            description: "Bôi đen vùng văn bản và nhấn nút để đánh dấu. Dùng \"Tách Tờ trình tùy chỉnh\" để chỉ định thư mục và tên file cho mỗi vùng tách.",
            items: [
                { label: "QĐ TL Ban chỉ huy", tag: "TT_BCH", icon: "user-check" },
                { label: "Tờ trình Kế hoạch", tag: "TT_KeHoach", icon: "calendar" },
                { label: "Tờ trình Nhân sự", tag: "TT_NhanSu", icon: "users" },
                { label: "Tờ trình Máy móc", tag: "TT_MayMoc", icon: "truck" },
                { label: "Tờ trình Vật liệu", tag: "TT_VatLieu", icon: "package" },
                { label: "Tờ trình Phòng TN", tag: "TT_ThiNghiem", icon: "microscope" },
                { label: "Tách Tờ trình tùy chỉnh", tag: "dynamic_split", icon: "scissors", color: "text-rose-500", bg: "bg-rose-50" }
            ]
        }
    ];

    sections.forEach(section => {
        const sectionDiv = document.createElement("div");
        sectionDiv.className = "space-y-4";
        sectionDiv.innerHTML = `
            <div class="pl-4 border-l-4 border-indigo-500 rounded-l-sm bg-gradient-to-r from-indigo-50/80 to-transparent py-2">
                <h4 class="text-[11px] font-black text-indigo-900 uppercase tracking-widest mb-1">${section.title}</h4>
                <p class="text-[11px] text-slate-600">${section.description}</p>
            </div>
            <div class="grid grid-cols-3 gap-2">
                ${section.items.map(item => `
                    <button data-tag="${item.tag}" data-label="${item.label}" 
                        class="template-btn flex flex-col items-center gap-3 p-4 bg-white rounded-2xl border border-slate-100 shadow-sm hover:shadow-md hover:border-indigo-200 transition-all group">
                        <div class="w-10 h-10 ${item.bg || 'bg-slate-50'} ${item.color || 'text-indigo-500'} rounded-xl flex items-center justify-center group-hover:scale-110 transition-transform">
                            <i data-lucide="${item.icon}" size="20"></i>
                        </div>
                        <span class="text-[11px] font-bold text-slate-700">${item.label}</span>
                    </button>
                `).join('')}
            </div>
        `;
        wrapper.appendChild(sectionDiv);
    });

    container.appendChild(wrapper);

    // Event Delegation
    wrapper.querySelectorAll('.template-btn').forEach(btn => {
        btn.onclick = async () => {
            const tag = btn.dataset.tag;
            const label = btn.dataset.label;

            try {
                btn.disabled = true;
                if (tag === 'dynamic_split') {
                    const result = await openSplitModal();

                    if (result) {
                        const { folder, file } = result;
                        // Encode: TT_[thuMuc]__[tenFile] — dấu __ phân cách thư mục và tên file
                        const safeFolder = WordService.normalizeTextForSearch(folder).replace(/\s+/g, "_");
                        const safeFile = WordService.normalizeTextForSearch(file).replace(/\s+/g, "_");
                        const bookmarkName = `TT_${safeFolder}__${safeFile}`;
                        await WordService.insertBookmarkAtSelection(bookmarkName);
                        showToast(`✓ Đã đánh dấu: "${file}" → 📁 ${folder}`, "success");
                    }
                } else if (tag.startsWith('bm') || tag.startsWith('TT_')) {
                    let headers = [];
                    let colCount = 0;
                    let noBorder = false;
                    let finalBookmarkName = tag;

                    if (tag === 'bmNhanSu') {
                        const choice = await openChoiceModal("Cấu hình Bảng Nhân sự", [
                            { label: "Bảng đầy đủ (5 cột)", description: "STT, Họ và tên, Chức vụ, Chuyên môn, Số điện thoại", value: 5 },
                            { label: "Bảng rút gọn (4 cột)", description: "STT, Họ và tên, Chức vụ, Chuyên môn", value: 4 },
                            { label: "Bảng cơ bản (3 cột)", description: "STT, Họ và tên, Chức vụ", value: 3 }
                        ]);

                        if (!choice) { btn.disabled = false; return; }

                        colCount = choice;
                        if (colCount === 5) {
                            headers = ["STT", "Họ và tên", "Chức vụ", "Chuyên môn", "Số điện thoại"];
                            finalBookmarkName = "bmNhanSu3";
                        } else if (colCount === 4) {
                            headers = ["STT", "Họ và tên", "Chức vụ", "Chuyên môn"];
                            finalBookmarkName = "bmNhanSu2";
                        } else if (colCount === 3) {
                            headers = ["STT", "Họ và tên", "Chức vụ"];
                            finalBookmarkName = "bmNhanSu";
                        }
                    }
                    else if (tag === 'bmMayMoc') { headers = ["STT", "Tên máy móc, thiết bị", "Đơn vị tính", "Số lượng", "Chủ sở hữu", "Tình trạng"]; colCount = 6; }
                    else if (tag === 'bmVatLieu') { headers = ["STT", "Tên vật liệu", "Tiêu chuẩn", "Nguồn gốc xuất xứ", "Đơn vị cung cấp"]; colCount = 5; }
                    else if (tag === 'bmThiNghiem') { headers = ["STT", "Đơn vị thí nghiệm", "Địa chỉ", "Mã số LAS", "Nội dung thí nghiệm"]; colCount = 5; }
                    else if (tag === 'bmKyLienDanh') { headers = ["Nơi nhận:", "ĐƠN VỊ THI CÔNG"]; colCount = 2; noBorder = true; }

                    if (colCount > 0) {
                        await WordService.insertTableWithBookmark(finalBookmarkName, colCount, headers, noBorder);
                        showToast(`✓ Đã chèn Bảng mẫu: ${label}`, "success");
                    } else {
                        await WordService.insertBookmarkAtSelection(finalBookmarkName);
                        showToast(`✓ Đã chèn Bookmark: ${label}`, "success");
                    }
                } else {
                    await WordService.insertContentControlAtSelection(tag, label);
                    showToast(`✓ Đ chèn trường: ${label}`, "success");
                }
            } catch (e) {
                console.error("Template Creator Error:", e);
                // Hiển thị lỗi chi tiết hơn để debug
                const errMsg = e.message || e.toString();
                showToast(`❌ Lỗi: ${errMsg}`, "error");
            } finally {
                btn.disabled = false;
            }
        };
    });

    lucide.createIcons();
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

    // Xử lý riêng cho lưu thông tin dự án từ Modal mới
    if (state.currentEditType === 'duAn') {
        config.fields.forEach(field => {
            const val = document.getElementById(`modalInput_${field}`)?.value || "";
            state.duAn[field] = val;
        });
        state.currentEditType = null;
    } else {
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

        if (type === 'mayMoc') {
            const ownerName = normalize(newEntry[4]);
            const companyName = normalize(state.duAn.dvtc);
            newEntry[5] = (ownerName.includes(companyName) || companyName.includes(ownerName)) ? "Sở hữu" : "Đi thuê";
        }

        if (state.editingIndex === -1) {
            state[type].push(newEntry);
        } else {
            state[type][state.editingIndex] = newEntry;
        }
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
    updateLog("Đang đồng bộ dữ liệu vào văn bản...", 5);
    await new Promise(r => setTimeout(r, 20)); // Nhường luồng cho UI vẽ thanh tiến trình

    // 1. Cập nhật Biến văn bản (DocVariables) như VBA và tag cụ thể
    const docVars = {
        "SoHD": state.duAn.soHD,
        "DuAn": state.duAn.tenDuAn,
        "GoiThau": state.duAn.goiThau,
        "DVTC": state.duAn.dvtc,
        "DaiDienCDT": state.duAn.daiDienCDT,
        "TVGS": state.duAn.tvgs,
        "NgayKhoiCong": state.duAn.ngayKhoiCong,
        "NgayHoanThanh": state.duAn.ngayHoanThanh,
        "IsLienDanh": state.duAn.isLienDanh ? "True" : "False"
    };

    try {
        updateLog("Cập nhật thông tin dự án...", 10);
        await new Promise(r => setTimeout(r, 10));
        // Thử cập nhật Content Controls trước
        await WordService.updateDocVariables(docVars);
    } catch (e) {
        updateLog("Content Controls không khả dụng", 10);
    }

    // Luôn cập nhật Document Variables (DOCVARIABLE) để phù hợp template dùng DOCVARIABLE
    try {
        updateLog("Gắn dữ liệu vào tài liệu...", 15);
        await WordService.updateDocumentVariables(docVars);
    } catch (e) {
        updateLog("Gắn dữ liệu thất bại", 15);
    }

    // Các phương thức đồng bộ hiện đại (Content Controls và DocVariables) đã xử lý xong phần thông tin dự án

    // Cập nhật các field trong document để DOCVARIABLE hiển thị giá trị mới.
    try {
        updateLog("Làm mới định dạng text...", 20);
        await WordService.updateAllFields();
    } catch (fieldError) {
        updateLog("Không thể update các field", 20);
    }

    // Bookmark bmTenDuAn đã được thay thế hoàn toàn bằng Content Control (tag DuAn)
    // Tuy nhiên khôi phục lại cơ chế placeholder <<Tag>> để đảm bảo tương thích 100% với mẫu cũ
    try {
        await WordService.replaceInDocument("<<SoHD>>", state.duAn.soHD || "", "SoHD");
        await WordService.replaceInDocument("<<DuAn>>", state.duAn.tenDuAn || "", "DuAn");
        await WordService.replaceInDocument("<<GoiThau>>", state.duAn.goiThau || "", "GoiThau");
        await WordService.replaceInDocument("<<DVTC>>", state.duAn.dvtc || "", "DVTC");
        await WordService.replaceInDocument("<<DaiDienCDT>>", state.duAn.daiDienCDT || "", "DaiDienCDT");
        await WordService.replaceInDocument("<<TVGS>>", state.duAn.tvgs || "", "TVGS");
        await WordService.replaceInDocument("<<NgayKhoiCong>>", state.duAn.ngayKhoiCong || "", "NgayKhoiCong");
        await WordService.replaceInDocument("<<NgayHoanThanh>>", state.duAn.ngayHoanThanh || "", "NgayHoanThanh");
    } catch (e) {
        updateLog("Lỗi thay thế placeholder: " + e.message);
    }

    // 2. Cập nhật Bảng (Table Syncs), ưu tiên Bookmark nếu có
    updateLog(`📊 Đang chèn bảng Nhân sự 1...`, 30);
    await WordService.xuatBang(state.nhanSu, "Họ và tên", "bmNhanSu", updateLog);
    updateLog(`📊 Đang chèn bảng Nhân sự 2...`, 40);
    await WordService.xuatBang(state.nhanSu, "Họ và tên", "bmNhanSu2", updateLog);
    updateLog(`📊 Đang chèn bảng Nhân sự 3...`, 50);
    await WordService.xuatBang(state.nhanSu, "Họ và tên", "bmNhanSu3", updateLog);
    updateLog(`📊 Đang chèn bảng Máy móc...`, 60);
    await WordService.xuatBang(state.mayMoc, "Tên thiết bị|Xe máy|Máy móc|Thiết bị", "bmMayMoc", updateLog);
    updateLog(`📊 Đang chèn bảng Vật liệu...`, 70);
    await WordService.xuatBang(state.vatLieu, "Tên vật tư|Tên vật liệu", "bmVatLieu", updateLog);
    updateLog(`📊 Đang chèn bảng Thí nghiệm...`, 80);
    await WordService.xuatBang(state.thiNghiem, "Đơn vị thí nghiệm", "bmThiNghiem", updateLog);

    // Xử lý bảng ký tên Liên danh hoặc Thường
    try {
        updateLog("Cập nhật bảng ký tên...", 85);
        const membersList = state.duAn.isLienDanh
            ? (Array.isArray(state.duAn.dvtcMembers) ? state.duAn.dvtcMembers : [])
            : [];
        await WordService.updateSignatureTable(state.duAn.isLienDanh, membersList, state.duAn.dvtc, "bmKyLienDanh", state.duAn.sigFontSize || 11, updateLog);
    } catch (e) {
        updateLog("Không thể cập nhật bảng ký: " + e.message);
    }

    // Format căn lề bảng đã được tích hợp trực tiếp trong xuatBang
    updateLog("Áp dụng kiểu dáng cuối cùng...", 90);
    await WordService.applyModernStyleToDocument();
}

async function verifyPermission(fileHandle) {
    const options = { mode: 'readwrite' };
    if ((await fileHandle.queryPermission(options)) === 'granted') {
        return true;
    }
    if ((await fileHandle.requestPermission(options)) === 'granted') {
        return true;
    }
    return false;
}

async function onCapNhatClick() {
    try {
        updateLog("--- Bắt đầu Cập nhật dữ liệu ---", 2);
        await syncDataToWord();
        updateLog("✓ Hoàn tất cập nhật dữ liệu vào văn bản.", 95);

        // Tự động ghi đè lại file đã xuất nếu có
        if (state.exportFolderHandle && (state.hasExportedMaster || (state.hasSplitFiles && state.autoSplitOnUpdate))) {
            const hasPerm = await verifyPermission(state.exportFolderHandle);
            if (hasPerm) {
                updateLog("Cập nhật thay thế các file DOCX gốc...", 96);
                if (state.hasExportedMaster) {
                    await WordService.processExport('master', state.duAn.tenDuAn, {
                        folderHandle: state.exportFolderHandle,
                        outputMode: state.outputMode,
                        useProjectNameFolder: state.useProjectNameFolder
                    });
                    updateLog("✓ Đã cập nhật ghi đè file tổng.", 97);
                }
                if (state.hasSplitFiles && state.autoSplitOnUpdate) {
                    await WordService.processExport('split', state.duAn.tenDuAn, {
                        folderHandle: state.exportFolderHandle,
                        outputMode: state.outputMode,
                        useProjectNameFolder: state.useProjectNameFolder
                    });
                    updateLog("✓ Đã cập nhật ghi đè file tách.", 98);
                }
                showToast("Đã cập nhật Word & xuất file thành công!", "success");
            } else {
                showToast("Không có quyền ghi vào thư mục để tự động xuất file!", "warning");
            }
                }
                updateLog("✓ Hoàn tất cập nhật hồ sơ.", 100);
            } catch (exportErr) {
                console.error("Post-update Export Error:", exportErr);
                updateLog(`⚠ Cập nhật Word xong, nhưng lỗi khi xuất file: ${exportErr.message}`, 100);
                showToast("Dữ liệu Word đã cập nhật, nhưng không thể ghi đè file xuất!", "warning");
            }
        } else {
            showToast("Đã cập nhật dữ liệu thành công!", "success");
            updateLog("✓ Hoàn tất cập nhật hồ sơ.", 100);
        }
    } catch (e) {
        console.error("onCapNhatClick Error:", e);
        const errorDetail = e.message || "Lỗi không xác định";
        updateLog(`❌ LỖI NGHIÊM TRỌNG: ${errorDetail}`, 100);
        if (e.stack) console.log(e.stack);
        showToast("Có lỗi xảy ra khi cập nhật (Xem Nhật ký bên dưới)", "error");
    }
}

async function onImportFromDocClick() {
    try {
        updateLog("Đang quét nội dung văn bản...", 30);
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

        updateLog(`✓ Đã nhập thành công ${updateCount} nhóm dữ liệu!`, 100);
        showToast("Đã khôi phục dữ liệu từ văn bản!", "success");
    } catch (e) {
        updateLog("Lỗi nhập liệu: " + e.message);
        showToast("Không thể nhập dữ liệu từ văn bản", "error");
    }
}

async function requestExportFolder() {
    if (state.exportFolderHandle) {
        const hasPerm = await verifyPermission(state.exportFolderHandle);
        if (hasPerm) return true;
        // Nếu không được cấp quyền, reset handle và yêu cầu chọn lại
        state.exportFolderHandle = null;
        state.exportFolderLabel = "";
    }

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
        updateLog("⏳ Đang chuẩn bị tách hồ sơ theo bookmark...", 20);

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
            outputMode: state.outputMode,
            useProjectNameFolder: state.useProjectNameFolder,
            onProgress: (msg, percent) => updateLog(msg, percent)
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
        updateLog("\n✅ Đã TÁCH HỒ SƠ thành công!", 100);
        showToast("Đã tách hồ sơ thành công!", "success");
    } catch (e) {
        updateLog("❌ Lỗi exception: " + e.message + "\n" + (e.stack || ""));
        showToast("Có lỗi xảy ra: " + e.message, "error");
    }
}

async function onXuatClick() {
    try {
        await syncDataToWord();
        updateLog("Đang xuất bộ hồ sơ tổng...", 30);
        await requestExportFolder();
        await WordService.processExport('master', state.duAn.tenDuAn, {
            folderHandle: state.exportFolderHandle,
            outputMode: state.outputMode,
            useProjectNameFolder: state.useProjectNameFolder,
            onProgress: (msg, percent) => updateLog(msg, percent)
        });
        state.hasExportedMaster = true;
        await saveState();
        updateLog("Đã XUẤT HỒ SƠ TỔNG (.docx) thành công!", 100);
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
                tenDuAn: "", goiThau: "", dvtc: "", daiDienCDT: "", tvgs: "", ngayKhoiCong: "", ngayHoanThanh: "", isLienDanh: false, dvtcMembers: []
            };
            state.nhanSu = [];
            state.mayMoc = [];
            state.vatLieu = [];
            state.thiNghiem = [];

            state.soHDForExport = "";
            state.hasExportedMaster = false;
            state.hasSplitFiles = false;

            // Cập nhật: Xóa thông tin thư mục đã chọn
            state.exportFolderHandle = null;
            state.exportFolderLabel = "";
            await StorageService.clearFolderHandle();

            await saveState();
            renderContent();
            lucide.createIcons();

            // Xóa dữ liệu cũ trong Word bằng cách đồng bộ state rỗng
            updateLog("Đang xóa trắng dữ liệu trong file Word...");
            await syncDataToWord();

            showToast("Đã làm mới toàn bộ dữ liệu!", "success");
            updateLog("✓ Đã xóa dữ liệu trên Add-in và trong file Word.", 100);
        } catch (err) {
            showToast("Lỗi khi xóa dữ liệu: " + err.message, "error");
        }
    };

    // Tự động thay đổi chiều cao textarea khi nhập liệu
    document.addEventListener('input', function (e) {
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
    const soHD = sanitizeFileName(duAn.soHD);
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

function openPromptModal(title, message, defaultValue = "") {
    return new Promise((resolve) => {
        const modal = document.getElementById('promptModal');
        const input = document.getElementById('promptInput');
        const error = document.getElementById('promptError');
        const btnConfirm = document.getElementById('promptConfirm');
        const btnCancel = document.getElementById('promptCancel');
        const titleElem = document.getElementById('promptTitle');
        const msgElem = document.getElementById('promptMessage');

        titleElem.innerText = title;
        msgElem.innerText = message;
        input.value = defaultValue;
        error.classList.add('hidden');
        modal.classList.remove('hidden');

        setTimeout(() => input.focus(), 50);

        const close = () => {
            modal.classList.add('hidden');
            btnConfirm.onclick = null;
            btnCancel.onclick = null;
        };

        btnConfirm.onclick = () => {
            const val = input.value.trim();
            if (!val) {
                error.classList.remove('hidden');
                return;
            }
            close();
            resolve(val);
        };

        btnCancel.onclick = () => {
            close();
            resolve(null);
        };
    });
}

// --- Modal 2 field: Tách Tờ trình tùy chỉnh (Phương án A) ---
function openSplitModal() {
    return new Promise((resolve) => {
        const modal = document.getElementById('splitModal');
        const folderInput = document.getElementById('splitFolderInput');
        const fileInput = document.getElementById('splitFileInput');
        const error = document.getElementById('splitModalError');
        const btnConfirm = document.getElementById('splitModalConfirm');
        const btnCancel = document.getElementById('splitModalCancel');

        folderInput.value = '';
        fileInput.value = '';
        error.classList.add('hidden');
        modal.classList.remove('hidden');
        lucide.createIcons();

        setTimeout(() => folderInput.focus(), 50);

        const close = () => {
            modal.classList.add('hidden');
            btnConfirm.onclick = null;
            btnCancel.onclick = null;
        };

        btnConfirm.onclick = () => {
            const folder = folderInput.value.trim();
            const file = fileInput.value.trim();
            if (!folder || !file) {
                error.classList.remove('hidden');
                return;
            }
            close();
            resolve({ folder, file });
        };

        // Nhấn Enter trong ô file cũng submit
        fileInput.onkeydown = (e) => {
            if (e.key === 'Enter') btnConfirm.click();
        };
        folderInput.onkeydown = (e) => {
            if (e.key === 'Enter') fileInput.focus();
        };

        btnCancel.onclick = () => {
            close();
            resolve(null);
        };
    });
}

function openChoiceModal(title, options) {
    return new Promise((resolve) => {
        const modal = document.getElementById('choiceModal');
        const titleElem = document.getElementById('choiceTitle');
        const optionsElem = document.getElementById('choiceOptions');
        const btnCancel = document.getElementById('choiceCancel');

        titleElem.innerText = title;
        optionsElem.innerHTML = '';

        const close = () => {
            modal.classList.add('hidden');
            btnCancel.onclick = null;
        };

        options.forEach(opt => {
            const btn = document.createElement('button');
            btn.className = 'w-full p-4 bg-slate-50 hover:bg-indigo-50 border border-slate-100 hover:border-indigo-200 rounded-2xl text-left transition-all group flex flex-col gap-1';
            btn.innerHTML = `
                <span class="text-xs font-black text-slate-800 uppercase tracking-wide group-hover:text-indigo-700">${opt.label}</span>
                <span class="text-[10px] text-slate-500">${opt.description}</span>
            `;
            btn.onclick = () => {
                close();
                resolve(opt.value);
            };
            optionsElem.appendChild(btn);
        });

        btnCancel.onclick = () => {
            close();
            resolve(null);
        };

        modal.classList.remove('hidden');
    });
}

function updateLog(m, progress = undefined) {
    const logMsgElem = document.getElementById('logMsg');
    if (logMsgElem) logMsgElem.innerText = m;
    const progressContainer = document.getElementById('loadingProgressContainer');
    const progressBar = document.getElementById('loadingProgressBar');
    const progressText = document.getElementById('loadingProgressText');
    if (progressContainer && progressBar) {
        if (progress !== undefined) {
            // Hiển thị trực tiếp bằng inline CSS để bẻ gãy mọi rào cản css/tailwind nếu có
            progressContainer.classList.remove('hidden');
            progressContainer.style.display = 'flex';
            progressContainer.style.opacity = '1';

            progressBar.style.width = `${progress}%`;
            if (progressText) progressText.innerText = `${Math.floor(progress)}%`;

            if (progress >= 100) {
                setTimeout(() => {
                    progressContainer.style.opacity = '0';
                    setTimeout(() => {
                        progressContainer.style.display = 'none';
                        progressContainer.classList.add('hidden');
                        progressBar.style.width = '0%';
                        if (progressText) progressText.innerText = '0%';
                    }, 300);
                }, 2000);
            }
        }
    }
}

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
