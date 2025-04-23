// main.js
const { app, BrowserWindow, ipcMain, clipboard, screen, Notification } = require('electron');
const exceljs = require('exceljs');
const path = require('path'); 
const os = require('os');
 
let mainWindow;
 
function createWindow() {
  const { width: screenWidth, height: screenHeight } = screen.getPrimaryDisplay().workAreaSize;
  const winWidth = 285;
  const winHeight = 510;
  const x = screenWidth - winWidth;
  const y = (screenHeight - winHeight) / 2;
 
  mainWindow = new BrowserWindow({
    width: winWidth,
    height: winHeight,
    x: x,
    y: y,
    frame: false,
    resizable: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
    alwaysOnTop: true,
  });
 
  mainWindow.loadFile('index.html');
 
  mainWindow.on('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.on('closed', function () {
    mainWindow = null;
  });

  ipcMain.on('quit-app', () => {
    app.quit();
  });

  ipcMain.handle('clipboard-read-text', () => {
    return clipboard.readText();
  });

  ipcMain.on('trans-text-copied', (event, {text, userID}) => {
    copyToTransID(text,userID)
      .then(() => {
        console.log('Text copied to Excel');
      })
      .catch((error) => {
        console.log('Error copying text to Excel:', error);
      });
  });
 
  ipcMain.on('EC-text-copied', (event, text) => {
    copyToErrorCode(text)
      .then(() => {
        console.log('Text copied to Excel');
      })
      .catch((error) => {
        console.log('Error copying text to Excel:', error);
      });
  });
 
  ipcMain.on('ED-text-copied', (event, text) => {
    copyToErrorDesc(text)
      .then(() => {
        console.log('Text copied to Excel');
       
      })
      .catch((error) => {
        console.log('Error copying text to Excel:', error);
      });
  });
 
  //IPC for Record-Time
 
  ipcMain.handle('text-copied-2', async (event, selectedOption) => {
    const username = os.userInfo().username;
    const homeDir = os.homedir(); // Get the user's home directory
    const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
    const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
    const workbook = new exceljs.Workbook();
 
    try {
        await workbook.xlsx.readFile(excelFilePath);
    } catch (error) {
        console.log('Error reading Excel file:', error.message);
        console.log('Creating a new workbook.');
    }
 
    const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
    let nextRow = 1;
 
    const headersExist = sheet.getCell(1, 8).value !== null;
 
    if (!headersExist) {
        const headers = ['AWB Action'];
        headers.forEach((header, columnIndex) => {
            const cell = sheet.getCell(nextRow, columnIndex + 8);
            cell.value = header;
            cell.font = { bold: true };
        });
 
        nextRow++;
    } else {
        while (sheet.getCell(nextRow, 8).value) {
            nextRow++;
        }
    }
 
    const targetColumns = ['AWB Action'];
    if (startTimeRow !== null) {
      targetColumns.forEach((columnName, columnIndex) => {
        sheet.getCell(startTimeRow, columnIndex + 8).value = selectedOption;
    });
  } else {
    console.log(`Start time is not recorded yet. "${selectedOption}" cannot be inserted.`);
    printWithNotification('Error',`Start time is not recorded yet. "${selectedOption}" cannot be inserted.`);
    // Handle the case where start time is not recorded yet
    return;
  }
    
    try {
      await workbook.xlsx.writeFile(excelFilePath);
      printWithNotification('Success', `Text "${selectedOption}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
      console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    } catch (writeError) {
      printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
    }
  });
 
 
 
ipcMain.handle('text-copied-3', async (event, selectedOption) => {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
  const workbook = new exceljs.Workbook();
 
  try {
      await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
      console.log('Error reading Excel file:', error.message);
      console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 9).value !== null;
 
  if (!headersExist) {
      const headers = ['Process'];
      headers.forEach((header, columnIndex) => {
          const cell = sheet.getCell(nextRow, columnIndex + 9);
          cell.value = header;
          cell.font = { bold: true };
      });
 
      nextRow++;
  } else {
      while (sheet.getCell(nextRow, 9).value) {
          nextRow++;
      }
  }
 
  const targetColumns = ['Process'];
  if (startTimeRow !== null) {
    targetColumns.forEach((columnName, columnIndex) => {
      sheet.getCell(startTimeRow, columnIndex + 9).value = selectedOption;
  });
  } else {
      console.log(`Start time is not recorded yet. "${selectedOption}" cannot be inserted.`);
      printWithNotification('Error',`Start time is not recorded yet. "${selectedOption}" cannot be inserted.`);
      // Handle the case where start time is not recorded yet
      return;
    }

    try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${selectedOption}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
  });
 
 
ipcMain.handle('record-time', () => {
  return getCurrentTimestamp();
});
 
ipcMain.handle('record-time2', async () => {
  const{currentTime2,row}=await getCurrentTimestamp2();
 
  // Calculate TAT and update the 'TAT' column
  await calculateTAT(row, row);
  return currentTime2;
});
  // Add the following code to handle the 'text-copied' event
  ipcMain.handle('text-copied', (event, text) => {
    const username = os.userInfo().username;
    const homeDir = os.homedir(); // Get the user's home directory
    const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
    const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
    const workbook = new exceljs.Workbook();
 
  try {
    workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
  const headersExist = sheet.getCell(1, 1).value !== null;
 
  if (!headersExist) {
    const headers = ['Text'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 1);
      cell.value = header;
      cell.font = { bold: true };
    });
 
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 1).value) {
      nextRow++;
    }
  }
 
  ipcMain.on('insert-text', async (event, rowData) => {
    const username = os.userInfo().username;
    const homeDir = os.homedir(); // Get the user's home directory
    const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
    const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
    const workbook = new exceljs.Workbook();
      try {
          await workbook.xlsx.readFile(excelFilePath);
      } catch (error) {
          console.log('Error reading Excel file:', error.message);
          console.log('Creating a new workbook.');
      }
  
      const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
  
      // Find the next available row
      let nextRow = sheet.rowCount + 1;
  
      // Insert user input text into the specified column (e.g., 12th column)
      sheet.getCell(nextRow, 12).value = rowData.text; // Assuming 'text' property contains user input
  
      try {
          await workbook.xlsx.writeFile(excelFilePath);
          console.log(`Text "${rowData.text}" appended to Excel at Row ${nextRow}, Column: 12`);
          return { success: true, message: `Text "${rowData.text}" appended to Excel at Row ${nextRow}, Column: 12`};
      } catch (writeError) {
          console.error('Error writing to Excel file:', writeError.message);
          return { success: false, message: 'Error writing to Excel file: ' + writeError.message };
      }
  });

  const targetColumns = ['Text'];
  targetColumns.forEach((columnName, columnIndex) => {
    sheet.getCell(nextRow, columnIndex + 1).value = text;
  });
 
  workbook.xlsx.writeFile(excelFilePath);
});
 
  const isAlwaysOnTop = mainWindow.isAlwaysOnTop();
  console.log('Is window always on top?', isAlwaysOnTop);
}
 
app.on('ready', createWindow);
 
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
 
app.on('activate', function () {
  if (mainWindow === null) {
    createWindow();
  }
});
 
function printWithNotification(title, message) {
  const notification = new Notification({
    title: title,
    body: message
  });
  notification.show();

  setTimeout(() => {
    notification.close();
  }, 1500);
}


async function copyToTransID(text,userID) {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
  const workbook = new exceljs.Workbook();
 
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Creating a new workbook.');
    printWithNotification('Error', 'Error reading Excel file: ' + error.message);
    printWithNotification('Info', 'Creating a new workbook.');
  }

  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 3).value !== null;
 
  if (!headersExist) {
    const headers = ['Tracking ID'];
    const headers1 = ['Processed By'];
    const headers2 = ['Sr. No.'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 3);
      cell.value = header;
      cell.font = { bold: true };
    });
    headers1.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 11);
      cell.value = header;
      cell.font = { bold: true };
    });
    headers2.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 1);
      cell.value = header;
      cell.font = { bold: true };
    });
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 3).value) {
      nextRow++;
    }
  }
 
  const serialNumber = nextRow - 1;
 
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
 
  const targetColumns = ['Tracking ID'];
  if (startTimeRow !== null) {
    targetColumns.forEach((columnName, columnIndex) => {
      sheet.getCell(startTimeRow, columnIndex + 3).value = text;
      sheet.getCell(startTimeRow, columnIndex + 11).value = userID;
  });
  } else {
    console.log("Start time is not recorded yet. Tracking ID cannot be inserted.");
    printWithNotification('Error',"Start time is not recorded yet. Tracking ID cannot be inserted.");
    // Handle the case where start time is not recorded yet
    return;
  }
 
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
}

// Similarly, modify copyToErrorCode and copyToErrorDesc functions

async function copyToErrorCode(text) {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
  const workbook = new exceljs.Workbook();
 
    try {
      await workbook.xlsx.readFile(excelFilePath);
    } catch (error) {
      console.log('Creating a new workbook.');
      printWithNotification('Error', 'Error reading Excel file: ' + error.message);
      printWithNotification('Info', 'Creating a new workbook.');
    }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 4).value !== null;
 
  if (!headersExist) {
    const headers = ['Error Code'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 4);
      cell.value = header;
      cell.font = { bold: true };
    });
 
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 4).value) {
      nextRow++;
    }
  }
 
  const serialNumber = nextRow - 1;
 
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
 
  const targetColumns = ['Error Code'];
  if (startTimeRow !== null) {
  targetColumns.forEach((columnName, columnIndex) => {
    sheet.getCell(startTimeRow, columnIndex + 4).value = text;
  });
} else {
  console.log("Start time is not recorded yet. Error Code cannot be inserted.");
  printWithNotification('Error',"Start time is not recorded yet. Error Code cannot be inserted.");
  // Handle the case where start time is not recorded yet
  return;
}
 
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
}

async function copyToErrorDesc(text) {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
  const workbook = new exceljs.Workbook();
 
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 5).value !== null;
 
  if (!headersExist) {
    const headers = ['Error Description'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 5);
      cell.value = header;
      cell.font = { bold: true };
    });
 
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 5).value) {
      nextRow++;
    }
  }
 
  const serialNumber = nextRow - 1;
 
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
 
  const targetColumns = ['Error Description'];
  if (startTimeRow !== null) {
    targetColumns.forEach((columnName, columnIndex) => {
      sheet.getCell(startTimeRow, columnIndex + 5).value = text;
  });
  } else {
    console.log("Start time is not recorded yet. Error Description cannot be inserted.");
    printWithNotification('Error',"Start time is not recorded yet. Error Description cannot be inserted.");
    // Handle the case where start time is not recorded yet
    return;
  }
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
    console.log(`Text "${text}" pasted to Excel at Row ${startTimeRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
}
 
let startTimeRow = null; // Variable to store the row where start time is inserted
 
async function getCurrentTimestamp() {

  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
  const workbook = new exceljs.Workbook();
 
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Error reading Excel file:', error.message);
    console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  // Find the next available row (assuming there is a 'Time' column in the first row)
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 6).value !== null;
 
  if (!headersExist) {
    const headers = ['Start Time'];
    const headers1 = ['Assign Date'];
    const headers2 = ['Processed By'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 6);
      cell.value = header;
      cell.font = { bold: true };
    });
    headers1.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 2);
      cell.value = header;
      cell.font = { bold: true };
    });
    headers2.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 11);
      cell.value = header;
      cell.font = { bold: true };
    });
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 6).value) {
      nextRow++;
    }
  }

    // Record the row where start time is inserted
    startTimeRow = nextRow;

  const serialNumber = nextRow - 1;
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
 
// Get the current time as a JavaScript Date object
  const currentTime1 = new Date();
 
  const currentTime1IST = currentTime1.toLocaleTimeString('en-IN', { hour12: false, timeZone: 'Asia/Kolkata' });
  const currentDate = currentTime1.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    timeZone: 'Asia/Kolkata'
  });
  const targetColumns = ['Start Time'];
  targetColumns.forEach((columnName, columnIndex) => {
    // Set the time in the 'Time' column
    sheet.getCell(nextRow, columnIndex + 6).value = currentTime1IST;
    // Set the number format for the cell to display time only
    sheet.getCell(nextRow, columnIndex + 6).numFmt = 'hh:mm:ss';
  });
  // Set the date in the 'Date' column
  sheet.getCell(nextRow, 2).value = currentDate;
  sheet.getCell(nextRow, 2).numFmt = 'dd/mm/yyyy';

  console.log(`Time "${currentTime1.toLocaleTimeString('en-US', { hour12: false })}" recorded to Excel at Row ${nextRow}, Column:${targetColumns}`);
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${currentTime1IST}" pasted to Excel at Row ${nextRow}, Column: ${targetColumns}`);
  
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
}
 

async function getCurrentTimestamp2() {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
  const workbook = new exceljs.Workbook(); 
 
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Error reading Excel file:', error.message);
    console.log('Creating a new workbook.');
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  // Find the next available row (assuming there is a 'Time' column in the first row)
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 7).value !== null;
 
  if (!headersExist) {
    const headers = ['End Time'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 7);
      cell.value = header;
      cell.font = { bold: true };
    });
 
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 7).value) {
      nextRow++;
    }
  }
  const serialNumber = nextRow - 1;
  // Insert serial number
  sheet.getCell(nextRow, 1).value = serialNumber;
 
  // Get the current time as a JavaScript Date object
  const currentTime2 = new Date();
  const currentTime2IST = currentTime2.toLocaleTimeString('en-IN', { hour12: false, timeZone: 'Asia/Kolkata' });
 
  const targetColumns = ['End Time'];
  if (startTimeRow !== null) {
  targetColumns.forEach((columnName, columnIndex) => {
    // Set the time in the 'Time' column
    sheet.getCell(startTimeRow, columnIndex + 7).value = currentTime2IST;
    // Set the number format for the cell to display time only
    sheet.getCell(startTimeRow, columnIndex + 7).numFmt = 'hh:mm:ss';
  });
} else {
  console.log("Start time is not recorded yet. End Time cannot be inserted.");
  printWithNotification('Error',"Start time is not recorded yet. End Time cannot be inserted.");
  // Handle the case where start time is not recorded yet
  return;
}
 
  console.log('Start Time:', sheet.getCell(nextRow, 6).value);
  console.log('End Time:', sheet.getCell(startTimeRow, 7).value);

  console.log(`Time: "${currentTime2.toLocaleTimeString('en-US', { hour12: false })}" recorded to Excel at Row ${nextRow}, Column:${targetColumns}`);
 
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    printWithNotification('Success', `Text "${currentTime2IST}" pasted to Excel at Row ${nextRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    printWithNotification('Error', 'Error writing to Excel file: ' + writeError.message);
  }
  return{currentTime2, row: nextRow};
}
 
// Helper function to format milliseconds to time (hh:mm:ss)
function formatMillisecondsToTime(milliseconds){
  const totalSeconds = Math.floor(milliseconds / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
 
  return `${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
}
// Helper function to pad single-digit numbers with a leading zero
function pad(number) {
  return number < 10 ? `0${number}` : number;
}
 
async function calculateTAT(startRow, endRow) {
  const username = os.userInfo().username;
  const homeDir = os.homedir(); // Get the user's home directory
  const excelFileName = `${username}_MI5_Agent_Cases.xlsx`;
  const excelFilePath = path.join(homeDir, 'OneDrive - Deloitte (O365D)', 'TEST', excelFileName);
  const workbook = new exceljs.Workbook();
 
  try {
    await workbook.xlsx.readFile(excelFilePath);
  } catch (error) {
    console.log('Error reading Excel file:', error.message);
    return;
  }
 
  const sheet = workbook.getWorksheet('Sheet1') || workbook.addWorksheet('Sheet1');
 
  // Find the next available row (assuming there is a 'Time' column in the first row)
  let nextRow = 1;
 
  const headersExist = sheet.getCell(1, 10).value !== null;
 
  if (!headersExist) {
    const headers = ['TAT'];
    headers.forEach((header, columnIndex) => {
      const cell = sheet.getCell(nextRow, columnIndex + 10);
      cell.value = header;
      cell.font = { bold: true };
    });
 
    nextRow++;
  } else {
    while (sheet.getCell(nextRow, 10).value) {
      nextRow++;
    }
  }
 
  // Ensure 'Start Time' and 'End Time' columns exist
  const startColumn = 6; // Column index for 'Start Time'
  const endColumn = 7;   // Column index for 'End Time'
  if (!sheet.getCell(1, startColumn).value || !sheet.getCell(1, endColumn).value) {
    console.log('Invalid Excel format. "Start Time" or "End Time" columns do not exist.');
    return;
  }
 
  const startTimeCell = sheet.getCell(startRow, startColumn);
  const endTimeCell = sheet.getCell(endRow, endColumn);
 
  let startTime = startTimeCell.text;  // Use .text instead of .value
  let endTime = endTimeCell.text;      // Use .text instead of .value
 
  // Check if both 'Start Time' and 'End Time' have valid values
  if (!startTime || !endTime) {
    console.log('Invalid start or end time format. Cannot calculate TAT.');
    return;
  }
 
  // Convert the time strings to JavaScript Date objects
  const startDate = new Date(`01/01/2000 ${startTime}`);
  const endDate = new Date(`01/01/2000 ${endTime}`);
 
  // Calculate TAT
  const tatMilliseconds = endDate.getTime() - startDate.getTime();
  const tatFormatted = formatMillisecondsToTime(tatMilliseconds);
 
  // Update the 'TAT' column
  const targetColumns = ['TAT'];
  targetColumns.forEach((columnName, columnIndex) => {
    sheet.getCell(endRow, columnIndex+10).value = tatFormatted;
  });  

  startTime = "00:00:00"; // Initialize to zero
  endTime = "00:00:00"; // Initialize to zero
 
  // Save the changes to the Excel file
  try {
    await workbook.xlsx.writeFile(excelFilePath);
    console.log(`TAT "${tatFormatted}" calculated and updated in Excel at Row ${endRow}, Column: ${targetColumns}`);
  } catch (writeError) {
    console.error('Error writing to Excel file:', writeError.message);
  }
}