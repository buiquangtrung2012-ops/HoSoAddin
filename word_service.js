/**
 * WORD SERVICE (Bản Hoàn Toàn Replice VBA - 1:1)
 * Hoạt động mạnh mẽ trên Word 2016-2021+
 */

/* global Word, Office, JSZip, saveAs */

export const WordService = {

    /**
     * Cập nhật thông tin dự án qua Content Controls (Chuẩn Word 2016 trở lên)
     */
    updateDocVariables: async (variablesObj) => {
        await Word.run(async (context) => {
            for (const [key, value] of Object.entries(variablesObj)) {
                const ctrls = context.document.contentControls.getByTag(key);
                ctrls.load("items");
                await context.sync();
                ctrls.items.forEach(cc => {
                    cc.insertText(value || " ", "Replace");
                });
            }
            await context.sync();
        });
    },


    /**
     * Cập nhật Document Variables
     */
    updateDocumentVariables: async (variablesObj) => {
        await Word.run(async (context) => {
            for (const [key, value] of Object.entries(variablesObj)) {
                try {
                    context.document.variables.getItem(key).set(value || " ");
                } catch (e) {
                    context.document.variables.add(key, value || " ");
                }
            }
            await context.sync();
        });
    },

    /**
     * Thay thế placeholder text trong document (ví dụ <<DuAn>>)
     * Kỹ thuật này giúp tương thích với các mẫu tài liệu cũ dùng nhãn văn bản.
     */
    replaceInDocument: async (searchText, replaceText, tag = null) => {
        await Word.run(async (context) => {
            const results = context.document.body.search(searchText, { matchCase: false, matchWholeWord: false });
            results.load("items");
            await context.sync();
            results.items.forEach(item => {
                const textToInsert = (replaceText && replaceText.length > 0) ? replaceText : " ";
                const textRange = item.insertText(textToInsert, "Replace");
                // Nếu muốn bọc Content Control sau khi thay thế
                if (tag) {
                    try {
                        const cc = textRange.insertContentControl();
                        cc.tag = tag;
                        cc.title = tag;
                        cc.appearance = "BoundingBox";
                    } catch (e) { /* CC might already exist or fail in some zones */ }
                }
            });
            await context.sync();
        });
    },

    /**
     * Cập nhật tất cả field trong document (DOCVARIABLE, DATE, v.v.)
     */
    updateAllFields: async () => {
        await Word.run(async (context) => {
            const fields = context.document.body.fields;
            fields.load("items");
            await context.sync();
            fields.items.forEach(field => {
                try { field.update(); } catch (e) { }
            });
            await context.sync();
        });
    },

    /**
     * Chuẩn hóa text để tìm kiếm (bỏ dấu tiếng Việt)
     */
    normalizeTextForSearch: (text) => {
        if (!text) return "";
        let s = text.toString().toLowerCase();
        const mapDia = {
            'à': 'a', 'á': 'a', 'ạ': 'a', 'ả': 'a', 'ã': 'a', 'â': 'a', 'ầ': 'a', 'ấ': 'a', 'ậ': 'a', 'ẩ': 'a', 'ẫ': 'a', 'ă': 'a', 'ằ': 'a', 'ắ': 'a', 'ặ': 'a', 'ẳ': 'a', 'ẵ': 'a',
            'è': 'e', 'é': 'e', 'ẹ': 'e', 'ẻ': 'e', 'ẽ': 'e', 'ê': 'e', 'ề': 'e', 'ế': 'e', 'ệ': 'e', 'ể': 'e', 'ễ': 'e',
            'ì': 'i', 'í': 'i', 'ị': 'i', 'ỉ': 'i', 'ĩ': 'i',
            'ò': 'o', 'ó': 'o', 'ọ': 'o', 'ỏ': 'o', 'õ': 'o', 'ô': 'o', 'ồ': 'o', 'ố': 'o', 'ộ': 'o', 'ổ': 'o', 'ỗ': 'o', 'ơ': 'o', 'ờ': 'o', 'ớ': 'o', 'ợ': 'o', 'ở': 'o', 'ỡ': 'o',
            'ù': 'u', 'ú': 'u', 'ụ': 'u', 'ủ': 'u', 'ũ': 'u', 'ư': 'u', 'ừ': 'u', 'ứ': 'u', 'ự': 'u', 'ử': 'u', 'ữ': 'u',
            'ỳ': 'y', 'ý': 'y', 'ỵ': 'y', 'ỷ': 'y', 'ỹ': 'y',
            'đ': 'd'
        };
        s = s.replace(/[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴ]/g, 'a')
            .replace(/[ÈÉẸẺẼÊỀẾỆỂỄ]/g, 'e')
            .replace(/[ÌÍỊỈĨ]/g, 'i')
            .replace(/[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]/g, 'o')
            .replace(/[ÙÚỤỦŨƯỪỨỰỬỮ]/g, 'u')
            .replace(/[ỲÝỴỶỸ]/g, 'y')
            .replace(/[Đ]/g, 'd');
        s = s.replace(/[\u0300\u0301\u0303\u0309\u0323]/g, "");
        s = s.replace(/[\u02C6\u0306\u031B]/g, "");
        s = s.replace(/["'«»""(){}\[\]]/g, "");
        s = s.replace(/\s+/g, " ").trim();
        Object.keys(mapDia).forEach(c => { s = s.replace(new RegExp(c, 'g'), mapDia[c]); });
        return s;
    },

    /**
     * Xuất dữ liệu vào bảng (Nuclear V5 - Heuristic & UI Logging)
     */
    xuatBang: async (data, keyword, bookmarkName = null, logCallback) => {
        const logger = (msg) => { if (logCallback) logCallback(msg); console.log(msg); };
        await Word.run(async (context) => {
            let targetTable = null;
            let targetColCount = 0;
            try {
                logger(`🔍 Đang chuẩn bị bảng ${bookmarkName || keyword}...`);

                // Bước 1: Tìm bảng (ưu tiên Bookmark)
                if (bookmarkName) {
                    try {
                        const bm = context.document.bookmarks.getItemOrNullObject(bookmarkName);
                        bm.load("isNullObject");
                        await context.sync();

                        if (!bm.isNullObject) {
                            const bmRange = bm.getRange();
                            const tablesInRange = bmRange.tables;
                            tablesInRange.load("items");
                            await context.sync();

                            if (tablesInRange.items.length > 0) {
                                targetTable = tablesInRange.items[0];
                                logger(`✓ Tìm thấy bảng trong vùng Bookmark ${bookmarkName}`);
                            } else {
                                // Bookmark nằm trước hoặc sau bảng – quét lân cận
                                const allTables = context.document.tables;
                                allTables.load("items");
                                await context.sync();

                                for (let t = 0; t < allTables.items.length; t++) {
                                    const table = allTables.items[t];
                                    const tableRange = table.getRange();
                                    const relation = tableRange.compareLocationWith(bmRange);
                                    await context.sync();

                                    if (relation.value === "After" || relation.value === "AdjacentAfter" || relation.value === "Overlapping") {
                                        targetTable = table;
                                        logger(`✓ Tìm thấy bảng lân cận Bookmark ${bookmarkName}`);
                                        break;
                                    }
                                }
                            }
                        }
                    } catch (err) {
                        console.warn(`xuatBang@bookmark(${bookmarkName})`, err.message);
                    }
                }

                if (!targetTable || targetTable.isNullObject) {
                    logger(`🔍 Bookmark thất bại, đang quét tìm bảng theo từ khóa "${keyword}"...`);
                    // Bước 2: Fallback – quét tất cả bảng, chọn bảng theo từ khóa tiêu đề
                    const tables = context.document.tables;
                    tables.load("items");
                    await context.sync();

                    const matchedTables = [];
                    for (let i = 0; i < tables.items.length; i++) {
                        const table = tables.items[i];
                        const fRow = table.rows.getFirst();
                        fRow.load("values");
                        await context.sync();

                        const rowText = fRow.values[0].join(" ");
                        const normRow = WordService.normalizeTextForSearch(rowText);
                        const keywords = keyword.split('|').map(k => WordService.normalizeTextForSearch(k));

                        if (keywords.some(k => normRow.includes(k))) {
                            matchedTables.push({ table, colCount: fRow.values[0].length });
                        }
                    }

                    if (matchedTables.length > 0) {
                        let matchIndex = 0;
                        if (bookmarkName) {
                            const match = bookmarkName.match(/(\d+)$/);
                            if (match) matchIndex = parseInt(match[1], 10) - 1;
                        }
                        if (matchIndex >= matchedTables.length) matchIndex = matchedTables.length - 1;
                        if (matchIndex < 0) matchIndex = 0;

                        targetTable = matchedTables[matchIndex].table;
                        targetColCount = matchedTables[matchIndex].colCount;
                        logger(`✓ Tìm thấy bảng qua từ khóa tại vị trí #${matchIndex + 1}`);
                    }
                }

                if (!targetTable || targetTable.isNullObject) {
                    logger(`✖ Không tìm thấy bảng đích cho ${bookmarkName || keyword}.`);
                    return;
                }

                const firstRow = targetTable.rows.getFirst();
                firstRow.load("values");
                targetTable.load("rowCount");
                await context.sync();

                targetColCount = firstRow.values[0].length;
                const rowCount = targetTable.rowCount;

                logger(`→ Xử lý bảng ${bookmarkName}: đang có ${rowCount} hàng, cần chèn ${data ? data.length : 0} dữ liệu.`);

                if (rowCount > 1) {
                    try {
                        logger(`🗑 Đang xóa ${rowCount - 1} hàng cũ...`);
                        targetTable.deleteRows(1, rowCount - 1);
                        await context.sync();
                    } catch (err) {
                        logger(`⚠ Lỗi xóa hàng nhanh, đang xóa thủ công...`);
                        targetTable.load("rows/items");
                        await context.sync();
                        for (let j = rowCount - 1; j >= 1; j--) {
                            try {
                                targetTable.rows.items[j].delete();
                                await context.sync();
                            } catch (e) { }
                        }
                    }
                }

                if (!data || data.length === 0) {
                    logger(`ℹ Không có dữ liệu để chèn cho ${bookmarkName}`);
                    return;
                }

                // Bước 2: Thêm dữ liệu mới
                const colCount = targetColCount || data[0].length;
                const newRowsValues = data
                    .filter(row => {
                        if (!row || row.length === 0) return false;
                        return row.slice(1).some(cell => cell && String(cell).trim() !== "");
                    })
                    .map(row => {
                        const normalizedRow = [];
                        for (let j = 0; j < colCount; j++) {
                            let val = row[j];
                            if (val === null || val === undefined) val = "";
                            normalizedRow.push(String(val));
                        }
                        return normalizedRow;
                    });

                if (newRowsValues.length > 0) {
                    logger(`📝 Đang chèn ${newRowsValues.length} hàng mới...`);
                    targetTable.addRows("End", newRowsValues.length, newRowsValues);

                    targetTable.headerRowCount = 1;
                    targetTable.load("rows/items");
                    await context.sync();

                    const justifiedKeywords = [
                        "ho va ten", "nhan su", "thiet bi", "xe may", "vat tu", "vat lieu",
                        "tieu chuan", "ghi chu", "don vi", "noi dung", "dia diem", "pham vi"
                    ];

                    // Định dạng chi tiết sẽ được xử lý trên từng đoạn văn (Paragraph) để tránh lỗi InvalidArgument

                    targetTable.rows.items.forEach((row, rIdx) => {
                        // Load paragraphs for formatting AND text for header identification
                        row.cells.load("items/body/paragraphs/items, items/body/text");
                    });
                    await context.sync();

                    const headerRow = targetTable.rows.items[0];
                    const headerTexts = headerRow.cells.items.map(cell => WordService.normalizeTextForSearch(cell.body.text || ""));

                    targetTable.rows.items.forEach((row, rIdx) => {
                        row.cells.items.forEach((cell, cIdx) => {
                            let cellAlignment = "Centered";
                            const headerText = headerTexts[cIdx] || "";
                            cell.verticalAlignment = "Center";

                            // Chỉ căn đều (Justified) các cột nội dung dài đặc thù
                            const isJustified = [
                                "ho va ten", "ten thiet bi", "ten vat tu", "tieu chuan", "don vi thi nghiem", "ghi chu", "noi dung"
                            ].some(kw => headerText.includes(kw));

                            // Nếu là cột STT (cột đầu tiên) thì luôn căn giữa
                            if (cIdx === 0) {
                                cellAlignment = "Centered";
                            } else {
                                if (rIdx === 0) {
                                    cellAlignment = "Centered";
                                } else {
                                    cellAlignment = isJustified ? "Justified" : "Centered";
                                }
                            }

                            // Sử dụng "#FFFFFF" (Màu trắng) để làm nền an toàn, không dùng "Clear" hay null để tránh InvalidArgument
                            cell.shadingColor = "#FFFFFF";

                            try {
                                cell.body.paragraphs.items.forEach(p => {
                                    p.alignment = cellAlignment;
                                    p.font.bold = (rIdx === 0);
                                });
                            } catch (e) { }
                        });
                    });

                    targetTable.headerRowCount = 1;
                    logger(`✓ Hoàn tất ${bookmarkName} với ${newRowsValues.length} hàng.`);
                }

                await context.sync();

            } catch (globalTableErr) {
                console.error(`xuatBang FATAL:`, globalTableErr);
                logger(`⚠️ Lỗi tại bảng ${bookmarkName || keyword}: ${globalTableErr.message}`);
            }
        });
    },
    
    /**
     * Cập nhật bảng ký tên Liên danh hoặc Thường (Bookmark: bmKyLienDanh)
     * Bản cập nhật v1130: Ultimate Safety - Chống sập do gộp ô/undefined body
     */
    updateSignatureTable: async (isLienDanh, membersList, dvtcName, bookmarkName, logCallback) => {
        const logger = (msg) => { if (logCallback) logCallback(msg); console.log(`[SignatureTable] ${msg}`); };
        
        // HELPER v1300: Fix căn lề, định dạng & Bỏ chữ nghiêng cho thành viên
        const safeFillCell = async (context, cell, text, isBold = true, alignment = "Centered") => {
            if (!cell || !text) return false;
            try {
                const range = cell.body.getRange();
                
                if (text === "Nơi nhận:") {
                    range.insertText("NƠI NHẬN:\n- Như trên;\n- Lưu VT.", "Replace");
                    await context.sync();
                    const ps = cell.body.paragraphs;
                    ps.load("items");
                    await context.sync();
                    for(let i=0; i<ps.items.length; i++) {
                        ps.items[i].font.set({ bold: i===0, italic: true, size: 10, name: "Times New Roman" });
                        try { ps.items[i].alignment = "Left"; } catch(e) {}
                    }
                } else {
                    range.insertText(text.toUpperCase(), "Replace");
                    // Ép italic: false để chữ đứng thẳng
                    range.font.set({ bold: isBold, italic: false, size: 11, name: "Times New Roman" });
                    await context.sync();
                    
                    const firstP = cell.body.paragraphs.getFirst();
                    try { firstP.alignment = alignment; } catch(e) {}
                }
                
                await context.sync();
                return true;
            } catch (e) {
                logger(`⚠️ Cell Lỗi: ${e.message}`);
                return false;
            }
        };

        const fillOneTable = async (context, tableData) => {
            const members = Array.isArray(membersList) ? membersList.filter(m => m && m.trim() !== "") : [];
            const itemsToFill = [];
            
            if (isLienDanh && members.length > 0) {
                itemsToFill.push("Nơi nhận:");
                members.forEach(m => itemsToFill.push(m));
            } else {
                itemsToFill.push("Nơi nhận:");
                itemsToFill.push((dvtcName || "").toUpperCase());
            }

            let itemIdx = 0;
            let rowIdx = 0;

            while (itemIdx < itemsToFill.length) {
                tableData.rows.load("items");
                await context.sync();
                
                if (rowIdx >= tableData.rows.items.length) {
                    tableData.addRows("End", 1);
                    await context.sync();
                    tableData.rows.load("items");
                    await context.sync();
                }

                const currentRow = tableData.rows.items[rowIdx];
                currentRow.cells.load("items");
                await context.sync();
                const cells = currentRow.cells.items;

                if (!cells || cells.length === 0) { rowIdx++; continue; }

                // 1. Điền nội dung vào tối đa 2 ô
                for (let colIdx = 0; colIdx < cells.length && colIdx < 2; colIdx++) {
                    if (itemIdx >= itemsToFill.length) break;
                    const text = itemsToFill[itemIdx];
                    const isNn = (text === "Nơi nhận:");
                    await safeFillCell(context, cells[colIdx], text, !isNn, isNn ? "Left" : "Centered");
                    itemIdx++;
                }

                // 2. Ép độ cao hàng sau khi đã có nội dung (Bản sửa lỗi v1300)
                try {
                    currentRow.set({
                        height: 107.71,
                        heightRule: "AtLeast"
                    });
                    await context.sync();
                } catch(e) { logger(`⚠️ Lỗi RowHeight: ${e.message}`); }

                rowIdx++;
            }
        };

        try {
            await Word.run(async (context) => {
                let targetTables = [];
                
                // 1. Kiểm tra Selection (V1300: Dùng try-catch để tránh crash ItemNotFound)
                try {
                    const sel = context.document.getSelection();
                    const selTable = sel.parentTable;
                    selTable.load("isNullObject");
                    await context.sync();
                    
                    if (!selTable.isNullObject) {
                        logger(`🎯 Cập nhật bảng đang chọn...`);
                        targetTables.push(selTable);
                    }
                } catch(e) {
                    logger(`ℹ️ Không có bảng nào được chọn (Hoặc con trỏ ngoài bảng).`);
                }

                // 2. Nếu không thấy Selection, Scan toàn bộ file
                if (targetTables.length === 0) {
                    logger(`🔍 Quét toàn bộ file tìm bảng ký tên...`);
                    const allTables = context.document.body.tables;
                    allTables.load("items");
                    await context.sync();
                    
                    for (let t of allTables.items) {
                        try {
                            const firstCell = t.getCell(0, 0);
                            const range = firstCell.body.getRange();
                            range.load("text");
                            await context.sync();
                            
                            if (range.text && range.text.toLowerCase().includes("nơi nhận")) {
                                targetTables.push(t);
                            }
                        } catch(e) {}
                    }
                }

                if (targetTables.length === 0) {
                    throw new Error("Không tìm thấy bảng. Hãy click vào bảng hoặc đảm bảo bảng có chữ 'Nơi nhận'.");
                }

                for (let i = 0; i < targetTables.length; i++) {
                    logger(`🖋️ [${i+1}/${targetTables.length}] Cập nhật bảng...`);
                    await fillOneTable(context, targetTables[i]);
                }

                await context.sync();
                logger(`✅ HOÀN TẤT: Đã cập nhật ${targetTables.length} bảng.`);
            });
        } catch (err) {
            logger(`❌ LỖI: ${err.message}`);
            console.error("[SignatureTable] Full error:", err);
            throw err;
        }
    },

    /**
     * Lấy nội dung file Word dưới dạng Blob
     */
    getFileContent: async () => {
        return new Promise((resolve, reject) => {
            Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const myFile = result.value;
                    const sliceCount = myFile.sliceCount;
                    let slicesReceived = 0;
                    const fileData = [];

                    const getSlice = (index) => {
                        myFile.getSliceAsync(index, (sliceResult) => {
                            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                fileData[index] = new Uint8Array(sliceResult.value.data);
                                slicesReceived++;
                                console.log(`File slice ${index + 1}/${sliceCount} received`);
                                if (slicesReceived === sliceCount) {
                                    myFile.closeAsync();
                                    const blob = new Blob(fileData, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
                                    console.log(`✓ File loaded: ${blob.size} bytes`);
                                    resolve(blob);
                                } else {
                                    // Get next slice with proper index
                                    getSlice(index + 1);
                                }
                            } else {
                                myFile.closeAsync();
                                reject(sliceResult.error);
                            }
                        });
                    };
                    getSlice(0);
                } else {
                    reject(result.error);
                }
            });
        });
    },

    /**
     * Xóa sạch dữ liệu khỏi Document (Reset Content Controls và Tables)
     */
    clearDocumentData: async () => {
        await Word.run(async (context) => {
            const docVarNames = ["DuAn", "GoiThau", "DVTC", "DaiDienCDT", "TVGS", "NgayKhoiCong", "NgayHoanThanh"];
            for (const key of docVarNames) {
                const ctrls = context.document.contentControls.getByTag(key);
                ctrls.load("items");
                await context.sync();
                ctrls.items.forEach(cc => { cc.insertText("", "Replace"); });
            }

            try {
                const variables = context.document.variables;
                variables.load("items");
                await context.sync();
                for (const key of docVarNames) {
                    try { variables.getItem(key).delete(); } catch (e) { }
                }
                await context.sync();
            } catch (e) { }

            const tables = context.document.tables;
            tables.load("items");
            await context.sync();

            const keywords = ["Họ và tên", "Tên Thiết Bị", "Tên vật tư", "Đơn vị thí nghiệm"];
            for (let i = 0; i < tables.items.length; i++) {
                const table = tables.items[i];
                const firstRow = table.rows.getFirst();
                firstRow.load("values");
                await context.sync();

                const rowText = firstRow.values[0].join(" ");
                const shouldClear = keywords.some(kw => rowText.toLowerCase().includes(kw.toLowerCase()));

                if (shouldClear) {
                    table.load("rowCount");
                    await context.sync();
                    const rowCount = table.rowCount;
                    if (rowCount > 1) {
                        try {
                            table.deleteRows(1, rowCount - 1);
                            await context.sync();
                        } catch (err) {
                            table.load("rows/items");
                            await context.sync();
                            for (let j = table.rows.items.length - 1; j >= 1; j--) {
                                try { table.rows.items[j].delete(); } catch (e) { }
                            }
                            await context.sync();
                        }
                    }
                }
            }

            await context.sync();
        });
    },

    /**
     * Bỏ dấu tiếng Việt (dùng cho tên file)
     */
    _removeDiacritics: (str) => {
        if (!str) return '';
        return str
            .replace(/[àáạảãâầấậẩẫăằắặẳẵ]/g, 'a').replace(/[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴ]/g, 'A')
            .replace(/[èéẹẻẽêềếệểễ]/g, 'e').replace(/[ÈÉẸẺẼÊỀẾỆỂỄ]/g, 'E')
            .replace(/[ìíịỉĩ]/g, 'i').replace(/[ÌÍỊỈĨ]/g, 'I')
            .replace(/[òóọỏõôồốộổỗơờớợởỡ]/g, 'o').replace(/[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]/g, 'O')
            .replace(/[ùúụủũưừứựửữ]/g, 'u').replace(/[ÙÚỤỦŨƯỪỨỰỬỮ]/g, 'U')
            .replace(/[ỳýỵỷỹ]/g, 'y').replace(/[ỲÝỴỶỸ]/g, 'Y')
            .replace(/đ/g, 'd').replace(/Đ/g, 'D')
            .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    },

    /**
     * Trích xuất địa danh hành chính từ tên dự án → định dạng VanHoa_BaVi
     * Tìm tất cả cấp: xã/phường/thị trấn, huyện/quận/thị xã, tỉnh/thành phố
     */
    xuLyLayTenDuAn: (tenDuAn) => {
        if (!tenDuAn) return 'Du_An';

        // Từ khóa sắp xếp DÀI TRƯỚC để tránh nhầm "thị" với "thị trấn"
        const adminKeywords = [
            'thị trấn', 'thị xã', 'thành phố',
            'xã', 'phường', 'huyện', 'quận', 'tỉnh'
        ];

        // Tách theo dấu phẩy/chấm phẩy → xử lý từng đoạn
        const segments = tenDuAn.split(/[,;]+/).map(s => s.trim()).filter(Boolean);
        const locationNames = [];

        for (const seg of segments) {
            const segLower = seg.toLowerCase();
            for (const kw of adminKeywords) {
                const idx = segLower.indexOf(kw);
                // Đảm bảo keyword là từ riêng (đứng sau khoảng trắng hoặc đầu chuỗi)
                if (idx !== -1 && (idx === 0 || segLower[idx - 1] === ' ')) {
                    const nameRaw = seg.slice(idx + kw.length).trim();
                    if (nameRaw) {
                        // Bỏ dấu + xóa khoảng trắng → "VanHoa", "BaVi"
                        const clean = WordService._removeDiacritics(nameRaw)
                            .replace(/\s+/g, '')
                            .replace(/[^A-Za-z0-9]/g, '');
                        if (clean) locationNames.push(clean);
                        break; // Mỗi segment chỉ lấy 1 tên
                    }
                }
            }
        }

        if (locationNames.length > 0) {
            return locationNames.join('_');
        }

        // Fallback: bỏ dấu tên gốc, giới hạn độ dài
        return WordService._removeDiacritics(tenDuAn)
            .replace(/[^A-Za-z0-9_]/g, '_')
            .replace(/_+/g, '_')
            .replace(/^_+|_+$/g, '')
            .slice(0, 80);
    },



    /**
     * Liệt kê tất cả bookmark có trong document
     */
    getAvailableBookmarks: async () => {
        return await Word.run(async (context) => {
            try {
                const bookmarks = context.document.bookmarks;
                bookmarks.load("items");
                await context.sync();

                const bookmarkNames = bookmarks.items.map(bm => bm.name);
                console.log(`📌 Available bookmarks in document: ${bookmarkNames.length} found`);
                bookmarkNames.forEach(name => {
                    console.log(`   - ${name}`);
                });

                return bookmarkNames;
            } catch (err) {
                console.error('Error getting bookmarks:', err);
                return [];
            }
        });
    },

    /**
     * Lấy OOXML của vùng bookmark chỉ định - Dùng search trên tất cả bookmarks để tránh API issue
     */
    getBookmarkOoxml: async (bookmarkName) => {
        return await Word.run(async (context) => {
            try {
                console.log(`🔍 Searching for bookmark: '${bookmarkName}'`);

                // Try method 1: getItemOrNullObject if available
                try {
                    const bmTest = context.document.bookmarks.getItemOrNullObject(bookmarkName);
                    bmTest.load("isNullObject");
                    await context.sync();

                    if (!bmTest.isNullObject) {
                        const range = bmTest.range;
                        range.load("text");
                        await context.sync();

                        console.log(`✓ Found bookmark '${bookmarkName}' via getItemOrNullObject, text length: ${range.text.length}`);
                        const ooxmlResult = range.getOoxml();
                        await context.sync();
                        return ooxmlResult.value;
                    }
                } catch (e1) {
                    console.log(`   Method 1 (getItemOrNullObject) failed, trying method 2...`);
                }

                // Try method 2: Iterate through all bookmarks
                const allBookmarks = context.document.bookmarks;
                allBookmarks.load("items");
                await context.sync();

                for (const bm of allBookmarks.items) {
                    if (bm.name === bookmarkName) {
                        const range = bm.range;
                        range.load("text");
                        await context.sync();

                        console.log(`✓ Found bookmark '${bookmarkName}' via iteration, text length: ${range.text.length}`);
                        const ooxmlResult = range.getOoxml();
                        await context.sync();
                        return ooxmlResult.value;
                    }
                }

                console.warn(`✗ Bookmark '${bookmarkName}' not found in iteration`);
                return null;

            } catch (err) {
                console.error(`getBookmarkOoxml '${bookmarkName}' lỗi:`, err.message || err);
                return null;
            }
        });
    },

    /**
     * Trích xuất nội dung <w:body> từ OOXML (bỏ sectPr – dùng sectPr của doc gốc)
     */
    _extractBodyContent: (ooxml) => {
        if (!ooxml) return '';

        // Nếu OOXML không chứa w:body (có thể là raw content), wrap nó
        if (!ooxml.includes('<w:body')) {
            // Content có thể là raw XML của paragraphs/tables, thử trực tiếp
            if (ooxml.includes('<w:p>') || ooxml.includes('<w:tbl>')) {
                const cleaned = ooxml.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '').trim();
                if (cleaned) {
                    console.log('_extractBodyContent: Raw content detected, using directly');
                    return cleaned;
                }
            }
            console.warn('_extractBodyContent: No <w:body> found and not raw content');
            return '';
        }

        // Dùng \b[^>]* để bắt cả trường hợp có thuộc tính: <w:body wsp:rsidR="...">
        const bodyMatch = ooxml.match(/<w:body\b[^>]*>([\s\S]*?)<\/w:body>/);
        if (!bodyMatch) {
            console.warn('_extractBodyContent: Cannot match <w:body>, OOXML preview:', ooxml.slice(0, 300));
            return '';
        }

        const extracted = bodyMatch[1]
            .replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '')
            .trim();

        if (extracted) {
            console.log('_extractBodyContent: Extracted successfully, length:', extracted.length);
        }
        return extracted;
    },

    /**
     * Tạo DOCX blob từ nhiều OOXML parts, dùng styles/rels của doc gốc
     */
    createSplitDocx: async (fullDocBlob, ooxmlParts) => {
        const validParts = ooxmlParts.filter(Boolean);
        if (validParts.length === 0) {
            console.warn('createSplitDocx: không có OOXML parts hợp lệ');
            return null;
        }

        try {
            const zip = await JSZip.loadAsync(fullDocBlob);
            const docEntry = zip.file("word/document.xml");
            if (!docEntry) {
                console.warn('createSplitDocx: không tìm thấy word/document.xml trong DOCX');
                return null;
            }
            const docXml = await docEntry.async("string");

            // Trích xuất body content từ mỗi OOXML part
            const bodyParts = validParts
                .map((ooxml, idx) => {
                    const extracted = WordService._extractBodyContent(ooxml);
                    if (!extracted) {
                        console.warn(`Part ${idx}: _extractBodyContent trả về rỗng`);
                    }
                    return extracted;
                })
                .filter(Boolean);

            if (bodyParts.length === 0) {
                console.warn('createSplitDocx: Tất cả parts đều rỗng sau khi trích xuất');
                return null;
            }

            // Kết hợp các body parts với paragraph breaks
            const combinedBody = bodyParts.join('\n');
            console.log('createSplitDocx: Kết hợp thành công, length:', combinedBody.length);

            // Giữ nguyên sectPr định dạng trang gốc
            const sectPrMatch = docXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/);
            const sectPr = sectPrMatch ? sectPrMatch[0] : '<w:sectPr><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

            // Thay thế body – dùng \b[^>]* để bắt cả trường hợp có thuộc tính
            const newDocXml = docXml.replace(
                /<w:body\b[^>]*>[\s\S]*?<\/w:body>/,
                `<w:body>${combinedBody}${sectPr}</w:body>`
            );

            const newZip = new JSZip();
            for (const [fileName, file] of Object.entries(zip.files)) {
                if (file.dir) continue;
                const content = fileName === 'word/document.xml'
                    ? newDocXml
                    : await file.async("uint8array");
                newZip.file(fileName, content);
            }

            const result = await newZip.generateAsync({ type: "blob" });
            console.log('createSplitDocx: Tạo blob thành công, size:', result.size);
            return result;
        } catch (err) {
            console.error('createSplitDocx lỗi:', err);
            return null;
        }
    },

    _saveBlobToFolder: async (dirHandle, relativePath, blob) => {
        if (!dirHandle || !relativePath || !blob) {
            return false;
        }

        const parts = relativePath.split('/').filter(Boolean);
        let current = dirHandle;

        for (let i = 0; i < parts.length - 1; i++) {
            current = await current.getDirectoryHandle(parts[i], { create: true });
        }

        const fileName = parts[parts.length - 1];
        const fileHandle = await current.getFileHandle(fileName, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        return true;
    },

    /**
     * Tách từng bookmark thành file DOCX riêng - Alternative approach
     */
    createDocxFromBookmarks: async (fullDocBlob, bookmarkNames) => {
        if (!bookmarkNames || bookmarkNames.length === 0) return null;

        try {
            const ooxmlParts = [];
            for (const bmName of bookmarkNames) {
                const ooxml = await WordService.getBookmarkOoxml(bmName);
                if (ooxml) {
                    ooxmlParts.push(ooxml);
                } else {
                    console.warn(`Bookmark '${bmName}' could not be retrieved`);
                }
            }

            if (ooxmlParts.length === 0) {
                console.warn('createDocxFromBookmarks: No valid bookmarks found');
                return null;
            }

            return await WordService.createSplitDocx(fullDocBlob, ooxmlParts);
        } catch (err) {
            console.error('createDocxFromBookmarks lỗi:', err);
            return null;
        }
    },

    /**
     * Đóng gói Hồ sơ (Xuất Tổng hoặc Tách Folder ZIP theo Bookmark)
     */
    processExport: async (mode, tenDuAn, options = {}) => {
        const { folderHandle = null, outputMode = 'zip', onProgress = null } = options || {};
        const useZip = outputMode === 'zip' || !folderHandle;

        // Trích xuất địa danh làm tên file/folder
        const baseName = WordService.xuLyLayTenDuAn(tenDuAn || "Du_An")
            .replace(/[\/\\:*?"<>|]/g, '_')
            .replace(/\s+/g, '_')
            .replace(/_+/g, '_')
            .replace(/^_+|_+$/, '');

        const fullBlob = await WordService.getFileContent();

        // Verify fullBlob is valid
        if (!fullBlob || fullBlob.size === 0) {
            console.error("❌ ERROR: fullBlob is invalid or empty!");
            throw new Error("Document content is empty - cannot export");
        }
        console.log(`✓ Document loaded: ${fullBlob.size} bytes, type: ${fullBlob.type}`);

        if (mode === 'master') {
            if (onProgress) onProgress("Đang tạo file tổng...", 50);
            const folderName = baseName || 'HoSo';
            const fileName = `HoSoTongHop_${baseName}.docx`;
            if (folderHandle) {
                await WordService._saveBlobToFolder(folderHandle, `${folderName}/${fileName}`, fullBlob);
                console.log(`✅ Lưu file tổng vào thư mục đã chọn: ${folderName}/${fileName}`);
                if (onProgress) onProgress("Đã lưu file tổng vào thư mục máy tính.", 80);
            } else {
                saveAs(fullBlob, fileName);
                if (onProgress) onProgress("Đã tải xuống file tổng.", 80);
            }
            return;
        }

        if (mode === 'split') {
            // Kiểm tra bookmark có sẵn 
            const availableBookmarks = await WordService.getAvailableBookmarks();
            const splitGroups = [
                {
                    folder: "1. To trinh Nhan su thi cong",
                    bookmarks: ["TT_BCH"],
                    fileName: "0.QD thanh lap BCH.docx"
                },
                {
                    folder: "1. To trinh Nhan su thi cong",
                    bookmarks: ["TT_NhanSu"],
                    fileName: "1. To trinh Nhan su thi cong.docx"
                },
                {
                    folder: "2. To trinh May moc thiet bi",
                    bookmarks: ["TT_MayMoc"],
                    fileName: "0. To trinh May moc thiet bi.docx"
                },
                {
                    folder: "3. To trinh Vat lieu su dung",
                    bookmarks: ["TT_VatLieu"],
                    fileName: "0. To trinh Vat lieu su dung.docx"
                },
                {
                    folder: "4. To trinh Ke hoach",
                    bookmarks: ["TT_KeHoach"],
                    fileName: "0. To trinh Ke hoach.docx"
                },
                {
                    folder: "5. To trinh Tien do Thi cong",
                    bookmarks: ["TT_TienDoThiCong"],
                    fileName: "0. To trinh Tien do Thi cong.docx"
                },
                {
                    folder: "6. To trinh Thi nghiem vat lieu",
                    bookmarks: ["TT_ThiNghiem"],
                    fileName: "0. To trinh Thi nghiem vat lieu.docx"
                }
            ];
            const rootFolderName = baseName || 'HoSo';
            console.log(`Total required export entries: ${splitGroups.length}`);

            const zip = useZip ? new JSZip() : null;
            let filesAdded = 0;
            let filesFailed = 0;

            // Log fullBlob info
            console.log(`📄 Full document blob: ${fullBlob ? fullBlob.size + ' bytes' : 'NULL'}`);
            console.log(`📂 Export mode: ${outputMode}, selected folder: ${folderHandle ? 'yes' : 'no'}, root folder: ${rootFolderName}`);

            for (let i = 0; i < splitGroups.length; i++) {
                const group = splitGroups[i];
                if (onProgress) onProgress(`Đang xử lý ${i+1}/${splitGroups.length}: ${group.folder}...`, 20 + Math.floor((i / splitGroups.length) * 60));
                
                const ooxmlParts = [];
                for (const bmName of group.bookmarks) {
                    const ooxml = await WordService.getBookmarkOoxml(bmName);
                    if (ooxml) {
                        ooxmlParts.push(ooxml);
                        console.log(`  ✓ Bookmark '${bmName}' retrieved, length: ${ooxml.length}`);
                    } else {
                        console.warn(`  ✗ Bookmark '${bmName}' not found or empty`);
                    }
                }

                // Tách file - luôn thêm vào ZIP với fallback nếu cần
                let fileBlob = null;

                if (ooxmlParts.length > 0) {
                    // Thử tách theo bookmark
                    fileBlob = await WordService.createSplitDocx(fullBlob, ooxmlParts);
                    if (fileBlob) {
                        console.log(`✅ Split thành công: ${group.folder}/${group.fileName} (${fileBlob.size} bytes)`);
                    } else {
                        console.warn(`⚠️ createSplitDocx thất bại, fallback sang full document: ${group.fileName}`);
                        fileBlob = fullBlob; // Fallback to full document
                    }
                } else {
                    console.warn(`⚠️ Bookmark không tìm thấy cho '${group.folder}', fallback sang full document`);
                    fileBlob = fullBlob; // Fallback to full document
                }

                if (fileBlob && fileBlob.size > 0) {
                    if (useZip) {
                        console.log(`   → Adding to ZIP: ${group.folder}/${group.fileName} (${fileBlob.size} bytes)`);
                        zip.folder(group.folder).file(group.fileName, fileBlob);
                    } else {
                        const path = `${rootFolderName}/${group.folder}/${group.fileName}`;
                        await WordService._saveBlobToFolder(folderHandle, path, fileBlob);
                        console.log(`   → Lưu file vào thư mục: ${path} (${fileBlob.size} bytes)`);
                    }
                    filesAdded++;
                } else {
                    console.error(`❌ File blob invalid or empty: ${group.fileName}`);
                    filesFailed++;
                }
            }

            console.log(`📊 Summary: ${filesAdded} files added, ${filesFailed} files failed`);

            if (useZip) {
                if (onProgress) onProgress("📦 Đang đóng gói file ZIP...", 85);
                const zipContent = await zip.generateAsync({ type: "blob" });
                console.log(`📦 ZIP file created: ${zipContent.size} bytes`);
                if (zipContent.size === 0) {
                    console.error("❌ WARNING: ZIP file is EMPTY (0 bytes)!");
                }
                if (folderHandle) {
                    await WordService._saveBlobToFolder(folderHandle, `HoSo_${baseName}_Da_Tach.zip`, zipContent);
                    console.log(`✅ Lưu ZIP vào thư mục đã chọn: HoSo_${baseName}_Da_Tach.zip`);
                    if (onProgress) onProgress("Đã lưu tệp ZIP.", 95);
                } else {
                    saveAs(zipContent, `HoSo_${baseName}_Da_Tach.zip`);
                    if (onProgress) onProgress("Đã tải xuống tệp ZIP.", 95);
                }
            } else {
                console.log('✅ Đã lưu tất cả tệp DOCX tách riêng vào thư mục đã chọn.');
                if (onProgress) onProgress("Đã xuất hoàn tất toàn bộ các tệp.", 95);
            }
        }
    },

    /**
     * Nhập dữ liệu ngược từ văn bản vào Add-in
     */
    importDataFromDoc: async () => {
        return await Word.run(async (context) => {
            const result = {
                duAn: {},
                nhanSu: [],
                mayMoc: [],
                vatLieu: [],
                thiNghiem: []
            };

            // 1. Đọc Content Controls (Thông tin dự án) với bộ lọc Alias thông minh
            const controls = context.document.contentControls;
            controls.load("items/tag,items/text");
            await context.sync();

            // Mở rộng tên thay thế (Aliases) cho các trường
            const fullTagMap = {
                "DuAn": ["DuAn", "tenDuAn", "TenDuAn", "Project", "Ten Du An"],
                "GoiThau": ["GoiThau", "goiThau", "tenGoiThau", "Package", "Goi Thau"],
                "DVTC": ["DVTC", "dvtc", "donViThiCong", "Contractor", "Don vi thi cong"],
                "DaiDienCDT": ["DaiDienCDT", "daiDienCDT", "chuDauTu", "CDT", "Client", "Dai dien CDT"],
                "TVGS": ["TVGS", "tvgs", "tuVanGiamSat", "Supervisor", "Tu van giam sat"],
                "SoHD": ["SoHD", "soHD", "ContractNumber", "Number", "So HD", "So hop dong", "SoHĐ"],
                "NgayKhoiCong": ["NgayKhoiCong", "ngayKhoiCong", "NgayKC", "StartDate", "Ngay khoi cong"],
                "NgayHoanThanh": ["NgayHoanThanh", "ngayHoanThanh", "NgayHT", "EndDate", "Ngay hoan thanh"]
            };

            const inventory = { tags: [], tables: [], variables: [] };

            controls.items.forEach(ctrl => {
                inventory.tags.push({ tag: ctrl.tag, text: ctrl.text });

                // Khớp dữ liệu dự án qua bộ lọc Alias
                for (const [key, aliases] of Object.entries(fullTagMap)) {
                    if (aliases.some(a => a.toLowerCase() === (ctrl.tag || "").toLowerCase())) {
                        if (ctrl.text && ctrl.text.trim().length > 0 && ctrl.text.trim() !== " " && !ctrl.text.includes("<<")) {
                            const internalKey = {
                                "DuAn": "tenDuAn",
                                "GoiThau": "goiThau",
                                "DVTC": "dvtc",
                                "DaiDienCDT": "daiDienCDT",
                                "TVGS": "tvgs",
                                "SoHD": "soHD",
                                "NgayKhoiCong": "ngayKhoiCong",
                                "NgayHoanThanh": "ngayHoanThanh"
                            }[key];
                            result.duAn[internalKey] = ctrl.text.trim();
                        }
                    }
                }
            });

            // 1.1 Đọc Document Variables (Phòng trường hợp file cũ dùng DOCVARIABLE)
            try {
                const docVars = context.document.variables;
                docVars.load("items/name,items/value");
                await context.sync();
                docVars.items.forEach(v => {
                    inventory.variables.push({ name: v.name, value: v.value });
                    for (const [key, aliases] of Object.entries(fullTagMap)) {
                        if (aliases.some(a => a.toLowerCase() === v.name.toLowerCase())) {
                            const internalKey = {
                                "DuAn": "tenDuAn", "GoiThau": "goiThau", "DVTC": "dvtc",
                                "DaiDienCDT": "daiDienCDT", "TVGS": "tvgs", "SoHD": "soHD",
                                "NgayKhoiCong": "ngayKhoiCong", "NgayHoanThanh": "ngayHoanThanh"
                            }[key];
                            if (!result.duAn[internalKey] && v.value && v.value.trim().length > 0) {
                                result.duAn[internalKey] = v.value.trim();
                            }
                        }
                    }
                });
            } catch (e) {
                console.log("DocVars not supported or empty");
            }

            // 2. Đọc Tables (Danh sách dữ liệu) - ƯU TIÊN bmNhanSu3 cho Nhân sự
            const tables = context.document.tables;
            tables.load("items");

            async function getValuesFromBookmark(bmName) {
                try {
                    const bm = context.document.bookmarks.getItemOrNullObject(bmName);
                    bm.load("isNullObject");
                    await context.sync();

                    if (bm.isNullObject) return null;

                    const range = bm.getRange();

                    // Thử lấy bảng trực tiếp trong range
                    let tab = range.tables.getFirstOrNullObject();
                    tab.load("values,isNullObject");
                    await context.sync();

                    if (tab && !tab.isNullObject) return tab.values;

                    // Nếu không thấy, thử tìm bảng bao quanh (parent table) nếu bookmark nằm trong ô
                    const parentTab = range.parentTable.load("values,isNullObject");
                    await context.sync();
                    if (parentTab && !parentTab.isNullObject) return parentTab.values;

                } catch (e) {
                    console.log(`Lỗi tìm Bookmark ${bmName}:`, e.message);
                }
                return null;
            }

            // Ưu tiên lấy Nhân Sự từ bmNhanSu3
            let nsFoundFromBookmark = false;
            const nsValues = await getValuesFromBookmark("bmNhanSu3");
            if (nsValues && nsValues.length > 1) {
                inventory.tables.push({ index: "bmNhanSu3", header: nsValues[0].join(" ") });
                for (let r = 1; r < nsValues.length; r++) {
                    const row = nsValues[r];
                    if (row[1] && row[1].trim() !== "" && row[1].length > 1) {
                        result.nhanSu.push([
                            row[0] || (result.nhanSu.length + 1).toString(),
                            row[1] || "", row[2] || "", row[3] || "", row[4] || ""
                        ]);
                    }
                }
                if (result.nhanSu.length > 0) nsFoundFromBookmark = true;
            }

            await context.sync();

            for (let i = 0; i < tables.items.length; i++) {
                const table = tables.items[i];
                table.load("values");
                await context.sync();

                const values = table.values;
                if (values.length < 1) continue;

                const columnCount = values[0]?.length || 0;
                const headerRowText = values[0]?.join(" ") || "";
                const normHeader = WordService.normalizeTextForSearch(headerRowText);

                // --- NEW: Heuristic for Summary Tables (2 columns, Key-Value pairs) ---
                // Chuyển values[0][0] sang String để tránh lỗi .length trên kiểu dữ liệu khác
                const firstCellText = values[0] && values[0][0] ? String(values[0][0]) : "";
                if (columnCount === 2 || (columnCount > 1 && firstCellText.length < 50)) {
                    for (let r = 0; r < values.length; r++) {
                        const row = values[r];
                        if (!row || row.length < 2) continue;
                        
                        const keyText = WordService.normalizeTextForSearch(row[0]);
                        const valText = (row[1] ?? "").toString().trim();
                        
                        if (valText.length > 0) {
                            if (keyText.includes("so hd") || keyText.includes("so hop dong")) {
                                if (!result.duAn.soHD) result.duAn.soHD = valText;
                            } else if (keyText === "ten du an") {
                                if (!result.duAn.tenDuAn) result.duAn.tenDuAn = valText;
                            } else if (keyText === "ten goi thau" || keyText === "goi thau") {
                                if (!result.duAn.goiThau) result.duAn.goiThau = valText;
                            } else if (keyText.includes("don vi thi cong")) {
                                if (!result.duAn.dvtc) result.duAn.dvtc = valText;
                            }
                        }
                    }
                }

                // Dấu hiệu nhận biết bảng Nhân sự (Phải có ít nhất 4 cột để tránh bảng tóm tắt ở trang 1)
                const isNhanSuTable = (normHeader.includes("ho va ten") || normHeader.includes("ten nhan su") || normHeader.includes("ho ten"))
                    && normHeader.includes("chuc danh")
                    && columnCount >= 4;

                // NEU DA TIM THAY TU BOOKMARK HOAC DA CO DU LIEU, TU CHOI CAC BANG NHAN SU KHAC
                if ((nsFoundFromBookmark || result.nhanSu.length > 0) && isNhanSuTable) {
                    continue;
                }

                inventory.tables.push({ index: i, header: headerRowText });

                // Nhận diện bảng Nhân sự (Nếu chưa có dữ liệu mới quét)
                if (isNhanSuTable) {
                    for (let r = 1; r < values.length; r++) {
                        const row = values[r];
                        if (row[1] && row[1].trim() !== "") {
                            result.nhanSu.push([
                                row[0] || (result.nhanSu.length + 1).toString(),
                                row[1] || "", row[2] || "", row[3] || "", row[4] || ""
                            ]);
                        }
                    }
                    continue; // Đã lấy xong nhân sự, chuyển sang bảng khác
                }
                // Nhận diện bảng Máy móc
                else if (normHeader.includes("thiet bi") || normHeader.includes("xe may") || normHeader.includes("may moc")) {
                    for (let r = 1; r < values.length; r++) {
                        const row = values[r];
                        if (row[1] && row[1].trim() !== "") {
                            result.mayMoc.push([
                                row[0] || (result.mayMoc.length + 1).toString(),
                                row[1] || "",
                                row[2] || "",
                                row[3] || "",
                                row[4] || "",
                                row[5] || ""
                            ]);
                        }
                    }
                }
                // Nhận diện bảng Vật liệu
                else if (normHeader.includes("vat tu") || normHeader.includes("vat lieu")) {
                    for (let r = 1; r < values.length; r++) {
                        const row = values[r];
                        if (row[1] && row[1].trim() !== "") {
                            result.vatLieu.push([
                                row[0] || (result.vatLieu.length + 1).toString(),
                                row[1] || "",
                                row[2] || "",
                                row[3] || "",
                                row[4] || ""
                            ]);
                        }
                    }
                }
                // Nhận diện bảng Thí nghiệm
                else if (normHeader.includes("thi nghiem")) {
                    for (let r = 1; r < values.length; r++) {
                        const row = values[r];
                        if (row[1]) {
                            result.thiNghiem.push([
                                row[0] || (result.thiNghiem.length + 1).toString(),
                                row[1] || "",
                                row[2] || "",
                                row[3] || "",
                                row[4] || ""
                            ]);
                        }
                    }
                }
            }

            return result;
        });
    },

    /**
     * Đồng bộ Style Toàn Document
     */
    applyModernStyleToDocument: async () => {
        // (Không ghi đè style toàn bộ document để giữ nguyên thiết lập của người dùng)
    }
};
