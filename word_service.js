/**
 * WORD SERVICE (Bản Hoàn Toàn Replice VBA - 1:1)
 * Hoạt động mạnh mẽ trên Word 2016-2021+
 */

/* global Word, Office, JSZip, saveAs */

export const WordService = {
    /**
     * Thay thế Bookmark bằng Text (Dùng Search để tránh lỗi API Bookmark cũ)
     */
    replaceTag: async (tagName, newText) => {
        await Word.run(async (context) => {
            const results = context.document.search(tagName, { matchCase: false });
            results.load("items");
            await context.sync();
            results.items.forEach((item) => {
                item.insertText(newText, "Replace");
            });
            await context.sync();
        });
    },

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
     * Thay thế placeholder text trong document (ví dụ <<DuAn>>)
     */
    replaceInDocument: async (searchText, replaceText, tag = null) => {
        await Word.run(async (context) => {
            const results = context.document.body.search(searchText, { matchCase: false, matchWholeWord: false });
            results.load("items");
            await context.sync();
            results.items.forEach(item => {
                const textToInsert = (replaceText && replaceText.length > 0) ? replaceText : " ";
                const textRange = item.insertText(textToInsert, "Replace");
                if (tag) {
                    const cc = textRange.insertContentControl();
                    cc.tag = tag;
                    cc.title = tag;
                    cc.appearance = "BoundingBox";
                }
            });
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
     * Cập nhật tất cả field trong document (DOCVARIABLE, DATE, v.v.)
     */
    updateAllFields: async () => {
        await Word.run(async (context) => {
            const fields = context.document.body.fields;
            fields.load("items");
            await context.sync();
            fields.items.forEach(field => {
                try { field.update(); } catch (e) {}
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
            'à':'a','á':'a','ạ':'a','ả':'a','ã':'a','â':'a','ầ':'a','ấ':'a','ậ':'a','ẩ':'a','ẫ':'a','ă':'a','ằ':'a','ắ':'a','ặ':'a','ẳ':'a','ẵ':'a',
            'è':'e','é':'e','ẹ':'e','ẻ':'e','ẽ':'e','ê':'e','ề':'e','ế':'e','ệ':'e','ể':'e','ễ':'e',
            'ì':'i','í':'i','ị':'i','ỉ':'i','ĩ':'i',
            'ò':'o','ó':'o','ọ':'o','ỏ':'o','õ':'o','ô':'o','ồ':'o','ố':'o','ộ':'o','ổ':'o','ỗ':'o','ơ':'o','ờ':'o','ớ':'o','ợ':'o','ở':'o','ỡ':'o',
            'ù':'u','ú':'u','ụ':'u','ủ':'u','ũ':'u','ư':'u','ừ':'u','ứ':'u','ự':'u','ử':'u','ữ':'u',
            'ỳ':'y','ý':'y','ỵ':'y','ỷ':'y','ỹ':'y',
            'đ':'d'
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
     * Xuất dữ liệu vào bảng (Căn chỉnh 1:1 theo logic VBA Column Alignment)
     */
    xuatBang: async (data, keyword, bookmarkName = null) => {
        await Word.run(async (context) => {
            let targetTable = null;
            let targetColCount = 0;

            // Bước 1: Thử tìm bảng trực tiếp trong vùng Bookmark
            if (bookmarkName) {
                try {
                    const bm = context.document.bookmarks.getItemOrNullObject(bookmarkName);
                    bm.load("isNullObject");
                    await context.sync();

                    if (!bm.isNullObject) {
                        const bmRange = bm.range;
                        const tablesInRange = bmRange.tables;
                        tablesInRange.load("items");
                        await context.sync();

                        if (tablesInRange.items.length > 0) {
                            targetTable = tablesInRange.items[0];
                            const firstRow = targetTable.rows.getFirst();
                            firstRow.load("values");
                            await context.sync();
                            targetColCount = firstRow.values[0].length;
                            console.log(`xuatBang: bookmark ${bookmarkName} found table with ${targetColCount} cols in range`);
                        } else {
                            // Bookmark nằm trước bảng – tìm bảng tiếp theo sau bookmark
                            const allTables = context.document.tables;
                            allTables.load("items");
                            await context.sync();

                            for (let t = 0; t < allTables.items.length; t++) {
                                const table = allTables.items[t];
                                const tableRange = table.getRange();
                                const relation = tableRange.compareLocationWith(bmRange);
                                await context.sync();

                                if (relation.value === "After" || relation.value === "AdjacentAfter") {
                                    targetTable = table;
                                    const firstRow = targetTable.rows.getFirst();
                                    firstRow.load("values");
                                    await context.sync();
                                    targetColCount = firstRow.values[0].length;
                                    console.log(`xuatBang: bookmark ${bookmarkName} fallback to next table #${t}`);
                                    break;
                                }
                            }
                        }
                    } else {
                        console.warn(`xuatBang: không tìm bookmark ${bookmarkName}, sẽ tìm theo keyword`);
                    }
                } catch (err) {
                    console.warn("xuatBang@bookmark", err);
                }
            }

            // Bước 2: Fallback – quét tất cả bảng, chọn bảng theo số thứ tự từ tên bookmark
            if (!targetTable) {
                const tables = context.document.tables;
                tables.load("items");
                await context.sync();

                const matchedTables = [];
                for (let i = 0; i < tables.items.length; i++) {
                    const table = tables.items[i];
                    const firstRow = table.rows.getFirst();
                    firstRow.load("values");
                    await context.sync();

                    const rowText = firstRow.values[0].join(" ");
                    const normRow = WordService.normalizeTextForSearch(rowText);
                    const keywords = keyword.split('|').map(k => WordService.normalizeTextForSearch(k));

                    if (keywords.some(k => normRow.includes(k))) {
                        matchedTables.push({ table, colCount: firstRow.values[0].length });
                    }
                }

                if (matchedTables.length > 0) {
                    let matchIndex = 0;
                    if (bookmarkName) {
                        const match = bookmarkName.match(/(\d+)$/);
                        if (match) {
                            matchIndex = parseInt(match[1], 10) - 1;
                        }
                    }
                    if (matchIndex >= matchedTables.length) matchIndex = matchedTables.length - 1;
                    if (matchIndex < 0) matchIndex = 0;

                    targetTable = matchedTables[matchIndex].table;
                    targetColCount = matchedTables[matchIndex].colCount;
                    console.log(`xuatBang: fallback tìm thấy ${matchedTables.length} bảng, chọn bảng số ${matchIndex + 1} cho ${bookmarkName}`);
                }
            }

            if (!targetTable) {
                console.warn(`xuatBang: bảng đích không tìm thấy cho bookmark="${bookmarkName}" keyword="${keyword}"`);
                return;
            }

            targetTable.load("rowCount");
            await context.sync();
            const rowCount = targetTable.rowCount;

            // Xóa dữ liệu cũ (giữ Header)
            if (rowCount > 1) {
                try {
                    targetTable.deleteRows(1, rowCount - 1);
                    await context.sync();
                } catch (err) {
                    targetTable.load("rows/items");
                    await context.sync();
                    for (let j = rowCount - 1; j >= 1; j--) {
                        try { targetTable.rows.items[j].delete(); } catch (e) {}
                    }
                    await context.sync();
                }
            }

            if (!data || data.length === 0) return;

            const colCount = targetColCount || data[0].length;
            const newRowsValues = data.map(row => {
                const normalizedRow = [];
                for (let j = 0; j < colCount; j++) {
                    normalizedRow.push(row[j] || "");
                }
                return normalizedRow;
            });

            targetTable.addRows("End", newRowsValues.length, newRowsValues);
            targetTable.headerRowCount = 1; // Ép lặp lại tiêu đề ngay sau khi thêm hàng
            
            targetTable.load("rows/items");
            targetTable.load("rows/cells/items");
            await context.sync();

            targetTable.rows.items.forEach((currentRow, rIdx) => {
                if (rIdx === 0) {
                    currentRow.font.bold = true;
                    currentRow.cells.items.forEach(cell => {
                        cell.horizontalAlignment = "Centered";
                        cell.verticalAlignment = "Center";
                    });
                } else {
                    currentRow.font.bold = false;
                    currentRow.cells.items.forEach((cell, cIdx) => {
                        const colName = keyword.toLowerCase();
                        let alignment = "Left";
                        if (cIdx === 0) alignment = "Centered";
                        
                        if (colName.includes("thiết bị")) {
                            if (cIdx === 2 || cIdx === 3 || cIdx === 5) alignment = "Centered";
                            if (cIdx === 4) alignment = "Justified";
                        } else if (colName.includes("vật tư") || colName.includes("thí nghiệm")) {
                            if (cIdx === 3 || cIdx === 4) alignment = "Centered";
                            if (cIdx === 2) alignment = "Justified";
                        } else if (colName.includes("họ và tên")) {
                            if (cIdx === 2 || cIdx === 3 || cIdx === 4) alignment = "Centered";
                        }
                        
                        // Áp dụng căn lề thông qua ParagraphFormat để có hiệu lực cao nhất
                        cell.getRange().paragraphFormat.alignment = (alignment === "Centered") ? "Centered" : alignment;
                        cell.verticalAlignment = "Center";
                    });
                }
            });

            // (Không ghi đè font/size để giữ nguyên thiết lập của người dùng)
            await context.sync();
        });
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
                    try { variables.getItem(key).delete(); } catch (e) {}
                }
                await context.sync();
            } catch (e) {}

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
                                try { table.rows.items[j].delete(); } catch (e) {}
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
        const { folderHandle = null, outputMode = 'zip' } = options || {};
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
            const folderName = baseName || 'HoSo';
            const fileName = `HoSoTongHop_${baseName}.docx`;
            if (folderHandle) {
                await WordService._saveBlobToFolder(folderHandle, `${folderName}/${fileName}`, fullBlob);
                console.log(`✅ Lưu file tổng vào thư mục đã chọn: ${folderName}/${fileName}`);
            } else {
                saveAs(fullBlob, fileName);
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
                    fileName: "QD thanh lap BCH.docx"
                },
                {
                    folder: "1. To trinh Nhan su thi cong",
                    bookmarks: ["TT_NhanSu"],
                    fileName: "To trinh Nhan su thi cong.docx"
                },
                {
                    folder: "2. To trinh May moc thiet bi",
                    bookmarks: ["TT_MayMoc"],
                    fileName: "To trinh May moc thiet bi.docx"
                },
                {
                    folder: "3. To trinh Vat lieu su dung",
                    bookmarks: ["TT_VatLieu"],
                    fileName: "To trinh Vat lieu su dung.docx"
                },
                {
                    folder: "4. To trinh Ke hoach",
                    bookmarks: ["TT_KeHoach"],
                    fileName: "To trinh Ke hoach.docx"
                },
                {
                    folder: "5. To trinh Tien do Thi cong",
                    bookmarks: ["TT_TienDoThiCong"],
                    fileName: "To trinh Tien do Thi cong.docx"
                },
                {
                    folder: "6. To trinh Thi nghiem vat lieu",
                    bookmarks: ["TT_ThiNghiem"],
                    fileName: "To trinh Thi nghiem vat lieu.docx"
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

            for (const group of splitGroups) {
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
                const zipContent = await zip.generateAsync({ type: "blob" });
                console.log(`📦 ZIP file created: ${zipContent.size} bytes`);
                if (zipContent.size === 0) {
                    console.error("❌ WARNING: ZIP file is EMPTY (0 bytes)!");
                }
                if (folderHandle) {
                    await WordService._saveBlobToFolder(folderHandle, `HoSo_${baseName}_Da_Tach.zip`, zipContent);
                    console.log(`✅ Lưu ZIP vào thư mục đã chọn: HoSo_${baseName}_Da_Tach.zip`);
                } else {
                    saveAs(zipContent, `HoSo_${baseName}_Da_Tach.zip`);
                }
            } else {
                console.log('✅ Đã lưu tất cả tệp DOCX tách riêng vào thư mục đã chọn.');
            }
        }
    },

    /**
     * Đồng bộ Style Toàn Document
     */
    applyModernStyleToDocument: async () => {
        // (Không ghi đè style toàn bộ document để giữ nguyên thiết lập của người dùng)
    }
};
