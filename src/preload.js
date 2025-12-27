const {contextBridge, ipcRenderer} = require('electron');

contextBridge.exposeInMainWorld('api',
  {
    getCurrentCompany: () => ipcRenderer.invoke('get-current-company'),
    createCompany: (data) => ipcRenderer.invoke('create-company', data),
  }
)


