/**
 * WORD SERVICE (Converted from VBA XuLyBangVaTagTen.bas)
 * Interacts with the Word Document Object Model.
 */

/* global Word */

export const WordService = {
    /**
     * Thay thế Tag (Bookmark) bằng văn bản mới
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
     * Xuất dữ liệu vào Bảng (Table) theo Bookmark
     * Tự động áp dụng font hiện đại (Segoe UI)
     */
    xuatBang: async (data, bmName) => {
        await Word.run(async (context) => {
            // SỬA LỖI: Sử dụng getByName chuẩn của Office JS
            const bookmark = context.document.bookmarks.getByName(bmName);
            bookmark.load("isNullObject");
            await context.sync();

            if (bookmark.isNullObject) {
                console.error(`Không tìm thấy bookmark: ${bmName}`);
                return;
            }

            const range = bookmark.getRange();
            const table = range.tables.getFirst();
            table.load(["rows", "style", "isNullObject"]);
            await context.sync();

            if (table.isNullObject) {
                console.error(`Không tìm thấy bảng tại bookmark: ${bmName}`);
                return;
            }

            // Xóa các dòng cũ (giữ lại header)
            const rowCount = table.rows.items.length;
            for (let i = rowCount - 1; i >= 1; i--) {
                table.rows.items[i].delete();
            }

            // Thêm dữ liệu mới
            const newRowsValues = data.map(row => row.map(cell => cell || ""));
            table.addRows("End", newRowsValues.length, newRowsValues);

            // --- ĐỊNH DẠNG MODERN STYLE ---
            const tableRange = table.getRange();
            tableRange.font.name = "Segoe UI";
            tableRange.font.size = 11;
            
            // Header đậm
            const headerRow = table.rows.getFirst();
            headerRow.font.bold = true;
            headerRow.shadingColor = "#F1F5F9"; // Light Slate Blue bg
            headerRow.font.color = "#1E293B"; // Dark Slate Gray text

            // --- CĂN CHỈNH CHI TIẾT (Alignment mapping from VBA) ---
            await context.sync();
            
            // Căn giữa cột STT cho mọi bảng (Cột 1)
            for (let i = 0; i < table.rows.items.length; i++) {
                table.rows.items[i].cells.items[0].horizontalAlignment = "Center";
                table.rows.items[i].cells.items[0].verticalAlignment = "Center";
            }

            // Phân loại xử lý theo Bookmark
            switch (bmName.toLowerCase()) {
                case "bmnhansu":
                case "bmnhansu2":
                case "bmnhansu3":
                    for (let i = 0; i < table.rows.items.length; i++) {
                        table.rows.items[i].cells.items[2].horizontalAlignment = "Center"; // Chức danh
                        if (table.rows.items[i].cells.items.length >= 4) {
                            table.rows.items[i].cells.items[3].horizontalAlignment = "Center"; // Chuyên ngành
                        }
                    }
                    break;
                case "bmmaymoc":
                    for (let i = 0; i < table.rows.items.length; i++) {
                        table.rows.items[i].cells.items[2].horizontalAlignment = "Center";
                        table.rows.items[i].cells.items[3].horizontalAlignment = "Center";
                        table.rows.items[i].cells.items[5].horizontalAlignment = "Center";
                    }
                    break;
            }

            await context.sync();
        });
    },

    /**
     * Làm mới toàn bộ văn bản sang Font hiện đại
     */
    applyModernStyleToDocument: async () => {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.font.name = "Segoe UI";
            body.font.size = 12;
            await context.sync();
        });
    }
};
