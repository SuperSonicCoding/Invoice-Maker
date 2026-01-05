const {contextBridge, ipcRenderer} = require('electron');

contextBridge.exposeInMainWorld('api',
  {
    getCurrentCompany: () => ipcRenderer.invoke('get-current-company'),
    getCompanies: () => ipcRenderer.invoke('get-companies'),
    createCompany: (data) => ipcRenderer.invoke('create-company', data),
    createInitialCompany: (data) => ipcRenderer.invoke('create-initial-company', data),
  }
)


