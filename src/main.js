try {
  require('electron-reloader')(module);
} catch {} // reloads electron app when it detects change

// Modules to control application life and create native browser window
const { app, BrowserWindow, ipcMain } = require('electron')
const path = require('node:path')
const sqlite3 = require('sqlite3').verbose();

// const appPath = app.getPath('userData');
const dbPath = path.join(__dirname, 'invoice_companies.db');
console.log('dbPath:', dbPath);
const db = new sqlite3.Database(dbPath);
console.log('db', db);


function createWindow () {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })

  // and load the index.html of the app.
  mainWindow.loadFile('index.html')

  // Open the DevTools.
  mainWindow.webContents.openDevTools()
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})

ipcMain.handle('get-current-company', async (event) => {
  return new Promise((resolve, reject) => {
    db.get('SELECT * FROM companies WHERE id=1', (err, user) => {
      if (err) {
        console.log('err', err);
        reject(err.message);
      }
      
      if (user) {
        resolve(user);
        console.log('user', user);
      } else {
        reject('No user found.');
      }
    });
  });
})

// handles creating a new company
ipcMain.handle('create-company', async (event, data) => {
  return new Promise((resolve, reject)  => {
    const {name, number, address, city, zipCode } = data;
    const insert = db.prepare('INSERT INTO companies (name, invoice_number, address, city, zip_code) VALUES (?,?,?,?,?)');
    insert.run(name, number, address, city, zipCode, function(err) {
      if (err) {
        reject(err.message);
      } else {
        resolve({ id: this.lastId, message: 'Data posted successfully' });
        insert.finalize(); // use when completely done with statement
      }
    });
  });
});

// handles creating the initial company profile
ipcMain.handle('create-initial-company', async (event, data) => {
  return new Promise((resolve, reject)  => {
    const {name, address, city, zipCode, phoneNumber } = data;
    const insert = db.prepare('INSERT INTO companies (name, address, city, zip_code, phone_number) VALUES (?,?,?,?,?)');
    insert.run(name, address, city, zipCode, phoneNumber, function(err) {
      if (err) {
        reject(err.message);
      } else {
        resolve({ id: this.lastId, message: 'Data posted successfully' });
        insert.finalize(); // use when completely done with statement
      }
    });
  });
});

// handles getting all of the companies
ipcMain.handle('get-companies', async (event) => {
  return new Promise((resolve, reject) => {
    db.all('SELECT * FROM companies WHERE id>1 ORDER BY name ASC', (err, companies) => {
      if (err) {
        console.log('err', err);
        reject(err.message);
      }

      if (companies) {
        resolve(companies);
        console.log('companies', companies);
      } else {
        reject('No companies found.');
      }
    });
  });
});

// handles updating data for a company
ipcMain.handle('update-company', async (event, data) => {
  return new Promise((resolve, reject) => {
    console.log('update');
    const {name, invoiceNumber, address, city, zipCode, quantity, unitPrice, description, id } = data;
    const update = db.prepare('UPDATE companies SET name=?,invoice_number=?,address=?,city=?,zip_code=?,quantity=?,unit_price=?, description=? WHERE id=?');
    update.run(name, invoiceNumber, address, city, zipCode, quantity, unitPrice, description, id, function(err) {
      if (err) {
        console.log('err', err);
        reject(err.message);
      } else {
        resolve({ id: this.lastId, message: 'Data updated successfully' });
        update.finalize();
      }
    });
  });
});
