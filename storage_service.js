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
    },

    /**
     * Lưu Directory Handle vào IndexedDB (Vì Office Settings không lưu được Object native)
     */
    saveFolderHandle: async (handle) => {
        try {
            const db = await StorageService._openDB();
            const tx = db.transaction("handles", "readwrite");
            const store = tx.objectStore("handles");
            await store.put(handle, "exportFolder");
            return true;
        } catch (e) {
            console.error("Lỗi lưu Handle vào IDB:", e);
            return false;
        }
    },

    /**
     * Lấy Directory Handle từ IndexedDB
     */
    getFolderHandle: async () => {
        try {
            const db = await StorageService._openDB();
            const tx = db.transaction("handles", "readonly");
            const store = tx.objectStore("handles");
            const request = store.get("exportFolder");
            return new Promise((resolve) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => resolve(null);
            });
        } catch (e) {
            return null;
        }
    },

    _openDB: () => {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open("HoSoAddinDB", 1);
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains("handles")) {
                    db.createObjectStore("handles");
                }
            };
            request.onsuccess = (e) => resolve(e.target.result);
            request.onerror = (e) => reject(e.target.error);
        });
    }
};
