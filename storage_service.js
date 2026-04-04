/**
 * STORAGE SERVICE (Bản nâng cấp - Hỗ trợ JSON Project Data)
 */

/* global Office */

export const StorageService = {
    /**
     * Lưu trữ dữ liệu JSON vào Settings của Document
     */
    setProjectData: async (key, data) => {
        return new Promise((resolve) => {
            const jsonData = JSON.stringify(data);
            Office.context.document.settings.set(key, jsonData);
            Office.context.document.settings.saveAsync((result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(true);
                } else {
                    console.error("Lỗi lưu trữ:", result.error);
                    resolve(false);
                }
            });
        });
    },

    /**
     * Lấy dữ liệu JSON từ Settings của Document
     */
    getProjectData: async (key) => {
        const jsonData = Office.context.document.settings.get(key);
        if (!jsonData) return null;
        try {
            return JSON.parse(jsonData);
        } catch (e) {
            console.error("Lỗi giải mã JSON:", e);
            return null;
        }
    }
};
