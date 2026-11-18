(function () {
  fetch('https://api.allorigins.win/raw?url=' + encodeURIComponent(https://script.google.com/macros/s/AKfycbxUyxZ1wMOqrZL84yNoA-4w2ynCJpKUniYlb2ijW76eAZj3sUTVzxsUF9tyfe_KuyLwsw/exec");
  const API_ORGANIZATIONS_URL = `${API_URL}?sheet=DANHSACHCOQUAN`;
  const DATABASE_NAME = "TemplateStorage";
  const STORE_NAME = "files";
  let databasePromise;

  async function getRecords() {
    const response = await fetch(API_URL);
    return response.json();
  }

  async function getOrganizations() {
    const response = await fetch(API_ORGANIZATIONS_URL);
    return response.json();
  }

  async function getRecordActionCapabilities() {
    const response = await fetch(`${API_URL}?capabilities=records&_=${Date.now()}`, { cache: "no-store" });
    return response.json();
  }

  function write(payload) {
    return fetch(API_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify(payload)
    });
  }

  function openTemplateDatabase() {
    if (!databasePromise) {
      databasePromise = new Promise((resolve, reject) => {
        const request = indexedDB.open(DATABASE_NAME, 1);
        request.onerror = () => reject(request.error);
        request.onupgradeneeded = event => {
          const database = event.target.result;
          if (!database.objectStoreNames.contains(STORE_NAME)) {
            database.createObjectStore(STORE_NAME);
          }
        };
        request.onsuccess = () => resolve(request.result);
      });
    }

    return databasePromise;
  }

  async function getTemplate() {
    const database = await openTemplateDatabase();
    return new Promise((resolve, reject) => {
      const request = database
        .transaction([STORE_NAME], "readonly")
        .objectStore(STORE_NAME)
        .get("mau_phuc_dap");
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  }

  async function saveTemplate(name, data) {
    const database = await openTemplateDatabase();
    return new Promise((resolve, reject) => {
      const transaction = database.transaction([STORE_NAME], "readwrite");
      transaction.objectStore(STORE_NAME).put({ name, data }, "mau_phuc_dap");
      transaction.oncomplete = resolve;
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error);
    });
  }

  async function deleteTemplate() {
    const database = await openTemplateDatabase();
    return new Promise((resolve, reject) => {
      const transaction = database.transaction([STORE_NAME], "readwrite");
      transaction.objectStore(STORE_NAME).delete("mau_phuc_dap");
      transaction.oncomplete = resolve;
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error);
    });
  }

  window.AppData = {
    getRecords,
    getOrganizations,
    getRecordActionCapabilities,
    createRecord: write,
    updateRecord: write,
    deleteRecord: write,
    addOrganization: write,
    updateOrganizationMail: write,
    getTemplate,
    saveTemplate,
    deleteTemplate
  };
})();
