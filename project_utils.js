/**
 * PROJECT UTILS (Converted from VBA XuLyLayTenDuAn.bas)
 */

export const ProjectUtils = {
    /**
     * Bỏ dấu tiếng Việt (Optimized JS version)
     */
    boDau: (str) => {
        if (!str) return "";
        return str
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/đ/g, "d")
            .replace(/Đ/g, "D");
    },

    /**
     * Tách Xã, Huyện, Tỉnh chuẩn phân cấp
     */
    layXaHuyenTinh: (tenDuAn) => {
        const tmp = " " + ProjectUtils.boDau(tenDuAn).toLowerCase();
        
        const arrThon = ["thon ", "ban ", "ap ", "to ", "khu "];
        const arrXa = ["xa ", "phuong ", "thi tran "];
        const arrHuyen = ["huyen ", "quan ", "thanh pho ", "thi xa "];
        const arrTinh = ["tinh ", "thanh pho "];
        
        const stopForThon = ["xa ", "phuong ", "thi tran ", "huyen ", "quan ", "thanh pho ", "thi xa ", "tinh "];
        const stopForXa = ["huyen ", "quan ", "thanh pho ", "thi xa ", "tinh "];
        const stopForHuyen = ["tinh "];
        
        let ketQua = [];
        
        const thon = ProjectUtils.timSauTuKhoa(tmp, arrThon, stopForThon);
        if (thon) ketQua.push(thon);
        
        const xa = ProjectUtils.timSauTuKhoa(tmp, arrXa, stopForXa);
        if (xa) ketQua.push(xa);
        
        const huyen = ProjectUtils.timSauTuKhoa(tmp, arrHuyen, stopForHuyen);
        if (huyen) ketQua.push(huyen);
        
        const tinh = ProjectUtils.timSauTuKhoa(tmp, arrTinh, []);
        if (tinh && !ketQua.includes(tinh)) ketQua.push(tinh);
        
        return ProjectUtils.sanitizeFileName(ketQua.join("_"));
    },

    /**
     * Tìm và tách chuỗi sau từ khóa (Regex based)
     */
    timSauTuKhoa: (s, arrTuKhoa, stopKeywords) => {
        for (const kw of arrTuKhoa) {
            const regex = new RegExp(`(?:^|[\\s,.\\-(-])${kw}([^,.\\-_;/\\)\\n\\r]+)`, "i");
            const match = s.match(regex);
            if (match && match[1]) {
                let result = match[1].trim();
                // Check for stop keywords manually for safety
                for (const skw of stopKeywords) {
                    const skwPos = result.toLowerCase().indexOf(skw.toLowerCase());
                    if (skwPos !== -1) {
                        result = result.substring(0, skwPos).trim();
                    }
                }
                return result;
            }
        }
        return "";
    },

    /**
     * Làm sạch tên file
     */
    sanitizeFileName: (fileName) => {
        if (!fileName) return "";
        return fileName
            .replace(/[/\\:*?"<>|]/g, "")
            .replace(/_{2,}/g, "_")
            .replace(/\s{2,}/g, " ")
            .trim();
    }
};
