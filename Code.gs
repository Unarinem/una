/**
 * =================================================================
 * | TICKET SYSTEM BACKEND - v6 ENHANCED (Final)                 |
 * =================================================================
 * This script acts as a Web API for an external frontend.
 * Handles image uploads, PDF ticket generation from a dynamic Slides ID,
 * advanced validation, offline sync, CSV import, and public validation.
 */

// --- CONFIGURATION ---
const SPREADSHEET_ID = '1-GZC_wVC56V8EZUhRZE9U_0pGbOSrKMHg_W-iSmY8w4';

// !!! IMPORTANT: FOLDER & TEMPLATE IDs !!!
const FALLBACK_SLIDES_TEMPLATE_ID = '1B36_c9w92gYoeBHXPum3-RlZLCnpuVCUsP9E0Haj3FI'; 

const FOLDER_IDS = {
  qrCodes: '1GizUGOvJkbf9nYJb37zBXF2ZNMfemMrB',
  generatedTickets: '1ZiolMDwReMOueAiDTgOs6qwbhusr-ar7',
  profileImages: '1u3uAXAvIi6SxE-rNNtI87eXhL52Remsj'
};

// --- GLOBAL VARIABLES ---
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const ticketsSheet = ss.getSheetByName('Tickets');
const usersSheet = ss.getSheetByName('Users');
const settingsSheet = ss.getSheetByName('Settings');
const logSheet = ss.getSheetByName('Log');

/**
 * Ensures a sheet with the given name exists, creating it if it doesn't.
 */
function ensureSheet(name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange("A1:Z1").setFontWeight("bold");
  }
  return sheet;
}

const userLogsSheet = ensureSheet('UserLogs', ['Timestamp', 'User Email', 'Action', 'Ticket Number', 'Details']);

/**
 * *** CRITICAL FIX FOR "Failed to fetch" ERROR ***
 * Handles pre-flight CORS requests from the browser.
 * This function is necessary for external websites (like Google Sites) to call the API.
 */
function doOptions(e) {
  return ContentService.createTextOutput()
    .withHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    });
}

/**
 * Handles GET requests for public ticket validation
 */
function doGet(e) {
  const action = e.parameter.action;
  const ticketNumber = e.parameter.ticket;
  
  if (action === 'validatePublic' && ticketNumber) {
    try {
      const result = validateTicketPublic(ticketNumber);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
        .setMimeType(ContentService.MimeType.JSON)
        .withHeaders({'Access-Control-Allow-Origin': '*'});
    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
        .setMimeType(ContentService.MimeType.JSON)
        .withHeaders({'Access-Control-Allow-Origin': '*'});
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid request' }))
    .setMimeType(ContentService.MimeType.JSON)
    .withHeaders({'Access-Control-Allow-Origin': '*'});
}

/**
 * Main API router for all POST requests from the frontend.
 */
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    let responseData;

    switch (request.action) {
      case 'getInitialData':
        responseData = getInitialData();
        break;
      case 'getUsers':
        responseData = getUsers();
        break;
      case 'handleGenerateTickets':
        responseData = handleGenerateTickets(request.payload);
        break;
      case 'processTicketValidation':
        responseData = processTicketValidation(request.payload.ticketNumber, request.payload.scannerId);
        break;
      case 'saveSettings':
        responseData = saveSettings(request.payload.settings, request.payload.files);
        break;
      case 'addUser':
        responseData = addUser(request.payload);
        break;
      case 'removeUser':
        responseData = removeUser(request.payload);
        break;
      case 'deleteTicket':
        responseData = deleteTicket(request.payload);
        break;
      case 'getUserLogs':
        responseData = getUserLogs(request.payload);
        break;
      case 'syncOfflineScans':
        responseData = syncOfflineScans(request.payload);
        break;
      case 'getHourlyAnalytics':
        responseData = getHourlyAnalytics();
        break;
      case 'importCsvTickets':
        responseData = importCsvTickets(request.payload);
        break;
      case 'validateTicketPublic':
        responseData = validateTicketPublic(request.payload.ticketNumber);
        break;
      default:
        throw new Error('Unknown action: ' + request.action);
    }

    const response = { status: 'success', data: responseData };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON)
      .withHeaders({'Access-Control-Allow-Origin': '*'});

  } catch (error) {
    logAction('SYSTEM_ERROR', 'API_ERROR', error.stack);
    const errorResponse = { status: 'error', message: error.message };
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON)
      .withHeaders({'Access-Control-Allow-Origin': '*'});
  }
}

// =================================================================
// |               NEW API FUNCTIONS                               |
// =================================================================

/**
 * Syncs offline scanned tickets with the database
 */
function syncOfflineScans(offlineScans) {
  const results = {
    successful: [],
    failed: [],
    duplicates: []
  };

  offlineScans.forEach(scan => {
    try {
      const ticketNumber = scan.ticketNumber;
      const timestamp = scan.timestamp;
      const scannerId = scan.scannerId;
      
      // Find the ticket
      const textFinder = ticketsSheet.createTextFinder(ticketNumber);
      const found = textFinder.findNext();
      
      if (!found) {
        results.failed.push({ ticketNumber: ticketNumber, reason: 'Ticket not found' });
        return;
      }

      const row = found.getRow();
      const ticketData = ticketsSheet.getRange(row, 1, 1, 8).getValues()[0];
      
      // Check if already scanned
      if (ticketData[2] === 'Scanned') {
        results.duplicates.push({ 
          ticketNumber: ticketNumber, 
          reason: 'Already scanned by ' + ticketData[5] + ' at ' + new Date(ticketData[4]).toLocaleString()
        });
        return;
      }

      // Update ticket status
      ticketsSheet.getRange(row, 3).setValue('Scanned');
      ticketsSheet.getRange(row, 5).setValue(new Date(timestamp));
      ticketsSheet.getRange(row, 6).setValue(scannerId);
      
      // Log the action with offline timestamp
      logAction(scannerId, 'SCAN_SUCCESS_OFFLINE', 'Offline scan synced for ticket: ' + ticketNumber, ticketNumber);
      
      results.successful.push({ ticketNumber: ticketNumber, syncedAt: new Date() });
      
    } catch (error) {
      results.failed.push({ ticketNumber: scan.ticketNumber, reason: error.message });
    }
  });

  return results;
}

/**
 * Gets hourly analytics data for dashboard charts
 */
function getHourlyAnalytics() {
  const logs = userLogsSheet.getDataRange().getValues();
  const headers = logs.shift();
  const timestampIndex = headers.indexOf('Timestamp');
  const actionIndex = headers.indexOf('Action');

  // Filter for successful scans only
  const scanLogs = logs.filter(row => 
    row[actionIndex] === 'SCAN_SUCCESS' || row[actionIndex] === 'SCAN_SUCCESS_OFFLINE'
  );

  // Group by hour
  const hourlyData = {};
  
  scanLogs.forEach(row => {
    const timestamp = new Date(row[timestampIndex]);
    const hour = timestamp.getHours();
    const date = timestamp.toDateString();
    const key = date + '-' + hour;
    
    if (!hourlyData[key]) {
      hourlyData[key] = {
        hour: hour,
        date: date,
        count: 0,
        timestamp: new Date(date + ' ' + hour + ':00:00')
      };
    }
    hourlyData[key].count++;
  });

  // Convert to array and sort by timestamp
  const sortedData = Object.values(hourlyData).sort((a, b) => a.timestamp - b.timestamp);
  
  return sortedData;
}

/**
 * Imports tickets from CSV data
 */
function importCsvTickets(csvData) {
  const csvContent = csvData.csvContent;
  const ticketType = csvData.ticketType;
  const lines = csvContent.split('\n').filter(line => line.trim() !== '');
  
  if (lines.length === 0) {
    throw new Error('CSV file is empty');
  }

  // Skip header row if it contains "GuestName" or similar
  const startIndex = lines[0].toLowerCase().includes('guest') || lines[0].toLowerCase().includes('name') ? 1 : 0;
  const guestNames = lines.slice(startIndex).map(line => line.trim().replace(/"/g, ''));
  
  if (guestNames.length === 0) {
    throw new Error('No guest names found in CSV');
  }

  const newTicketsData = [];
  const generatedAt = new Date();
  const user = Session.getEffectiveUser().getEmail();
  const settings = getEventSettings();
  let currentRow = ticketsSheet.getLastRow();

  // Check if template is available
  if (!settings.slidesTemplateId && !FALLBACK_SLIDES_TEMPLATE_ID) {
    throw new Error("Ticket Generation Failed: Please set a Google Slides Template URL in the settings.");
  }

  for (let i = 0; i < guestNames.length; i++) {
    currentRow++;
    const guestName = guestNames[i];
    const ticketNumber = 'TK' + ticketType.toUpperCase().slice(0,3) + '241031' + String(currentRow).padStart(4, '0');
    
    // Generate QR code
    const qrCodeBlob = UrlFetchApp.fetch('https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=' + encodeURIComponent(ticketNumber)).getBlob().setName(ticketNumber + '_qr.png');
    const qrCodeFile = DriveApp.getFolderById(FOLDER_IDS.qrCodes).createFile(qrCodeBlob);
    qrCodeFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const qrCodeUrl = 'https://drive.google.com/uc?id=' + qrCodeFile.getId();

    // Generate PDF with guest name
    const ticketPdfUrl = createIndividualTicketWithGuest(ticketNumber, qrCodeFile, settings, guestName);

    const newRow = [
      ticketNumber, ticketType, 'Not Scanned', generatedAt, '', '', ticketPdfUrl, qrCodeUrl, guestName
    ];
    newTicketsData.push(newRow);
  }

  // Add guest name column header if it doesn't exist
  if (ticketsSheet.getRange(1, 9).getValue() !== 'Guest Name') {
    ticketsSheet.getRange(1, 9).setValue('Guest Name');
  }

  if (newTicketsData.length > 0) {
    ticketsSheet.getRange(ticketsSheet.getLastRow() + 1, 1, newTicketsData.length, newTicketsData[0].length).setValues(newTicketsData);
  }
  
  logAction(user, 'IMPORT_CSV_TICKETS', 'Imported ' + guestNames.length + ' tickets from CSV for ' + ticketType + '.');
  
  return {
    count: guestNames.length,
    tickets: newTicketsData.map(t => ({
      ticketNumber: t[0], ticketType: t[1], scanStatus: t[2], generatedAt: t[3], 
      scannedAt: t[4], scannedBy: t[5], pdfUrl: t[6], qrCodeUrl: t[7], guestName: t[8]
    }))
  };
}

/**
 * Public ticket validation (no authentication required)
 */
function validateTicketPublic(ticketNumber) {
  if (!ticketNumber || ticketNumber.trim() === '') {
    return { status: 'INVALID', message: 'Please enter a ticket number.' };
  }

  const textFinder = ticketsSheet.createTextFinder(ticketNumber.trim());
  const found = textFinder.findNext();
  
  if (!found) {
    return { status: 'INVALID', message: 'This ticket number does not exist in our system.' };
  }

  const row = found.getRow();
  const ticketData = ticketsSheet.getRange(row, 1, 1, 9).getValues()[0];
  const ticketInfo = {
    ticketNumber: ticketData[0], 
    ticketType: ticketData[1], 
    scanStatus: ticketData[2],
    scannedAt: ticketData[4], 
    scannedBy: ticketData[5],
    guestName: ticketData[8] || null
  };

  if (ticketInfo.scanStatus === 'Scanned') {
    return { 
      status: 'SCANNED', 
      message: 'This ticket has already been used.',
      ticketNumber: ticketInfo.ticketNumber, 
      ticketType: ticketInfo.ticketType,
      guestName: ticketInfo.guestName,
      scannedAt: ticketInfo.scannedAt ? new Date(ticketInfo.scannedAt).toLocaleString() : null
    };
  }

  return { 
    status: 'VALID', 
    message: 'This is a valid ticket that has not been used yet.',
    ticketNumber: ticketInfo.ticketNumber, 
    ticketType: ticketInfo.ticketType,
    guestName: ticketInfo.guestName
  };
}

// =================================================================
// |               API-Exposed Functions                           |
// =================================================================

function getInitialData() {
  const settings = getEventSettings();
  const tickets = getTickets();
  const users = getUsers();
  const stats = getDashboardStats(tickets, settings.ticketTypes || []);
  
  return {
    userInfo: { email: Session.getEffectiveUser().getEmail() },
    settings: settings,
    tickets: tickets,
    users: users,
    stats: stats
  };
}

function handleGenerateTickets(generationForm) {
  const count = generationForm.count;
  const ticketType = generationForm.ticketType;
  const newTicketsData = [];
  const generatedAt = new Date();
  const user = Session.getEffectiveUser().getEmail();
  const settings = getEventSettings();
  let currentRow = ticketsSheet.getLastRow();

  // Constructive feedback check
  if (!settings.slidesTemplateId && !FALLBACK_SLIDES_TEMPLATE_ID) {
    throw new Error("Ticket Generation Failed: Please set a Google Slides Template URL in the settings.");
  }

  for (let i = 0; i < count; i++) {
    currentRow++;
    const ticketNumber = 'TK' + ticketType.toUpperCase().slice(0,3) + '241031' + String(currentRow).padStart(4, '0');
    
    const qrCodeBlob = UrlFetchApp.fetch('https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=' + encodeURIComponent(ticketNumber)).getBlob().setName(ticketNumber + '_qr.png');
    
    const qrCodeFile = DriveApp.getFolderById(FOLDER_IDS.qrCodes).createFile(qrCodeBlob);
    qrCodeFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const qrCodeUrl = 'https://drive.google.com/uc?id=' + qrCodeFile.getId();

    const ticketPdfUrl = createIndividualTicket(ticketNumber, qrCodeFile, settings);

    const newRow = [
      ticketNumber, ticketType, 'Not Scanned', generatedAt, '', '', ticketPdfUrl, qrCodeUrl, ''
    ];
    newTicketsData.push(newRow);
  }

  // Ensure guest name column exists
  if (ticketsSheet.getRange(1, 9).getValue() !== 'Guest Name') {
    ticketsSheet.getRange(1, 9).setValue('Guest Name');
  }

  if (newTicketsData.length > 0) {
    ticketsSheet.getRange(ticketsSheet.getLastRow() + 1, 1, newTicketsData.length, newTicketsData[0].length).setValues(newTicketsData);
  }
  
  logAction(user, 'GENERATE_TICKETS', 'Generated ' + count + ' x ' + ticketType + ' tickets.');
  
  return newTicketsData.map(t => ({
      ticketNumber: t[0], ticketType: t[1], scanStatus: t[2], generatedAt: t[3], 
      scannedAt: t[4], scannedBy: t[5], pdfUrl: t[6], qrCodeUrl: t[7], guestName: t[8]
  }));
}

function saveSettings(settingsObj, files) {
  if (files.eventPicture) {
    settingsObj.eventPicture = saveBase64ImageToDrive(files.eventPicture, 'event_profile_pic', FOLDER_IDS.profileImages);
  }
  if (files.ticketTemplate) {
     settingsObj.ticketTemplate = saveBase64ImageToDrive(files.ticketTemplate, 'ticket_template_bg', FOLDER_IDS.profileImages);
  }

  const settingsRange = settingsSheet.getRange('A2:B' + (settingsSheet.getLastRow() || 2));
  const settingsData = settingsRange.getValues();
  
  for (const key in settingsObj) {
    if (settingsObj.hasOwnProperty(key)) {
        let value = settingsObj[key];
        let rowIndex = settingsData.findIndex(row => row[0] === key) + 2;

        if (key === 'slidesTemplateUrl' && value) {
            const match = value.match(/presentation\/d\/([a-zA-Z0-9-_]+)/);
            value = (match && match[1]) ? match[1] : value;
            settingsSheet.getRange(rowIndex, 2).setValue(value);
            settingsObj.slidesTemplateId = value;
        } else {
          if (Array.isArray(value)) value = value.join(',');
          if (rowIndex > 1) {
              settingsSheet.getRange(rowIndex, 2).setValue(value);
          } else {
              settingsSheet.appendRow([key, value]);
          }
        }
    }
  }

  logAction(Session.getEffectiveUser().getEmail(), 'SAVE_SETTINGS', 'Event settings updated.');
  return getEventSettings();
}

function getUserLogs(email) {
  if (!email) return [];
  const allLogs = userLogsSheet.getDataRange().getValues();
  const headers = allLogs.shift();
  const userEmailIndex = headers.indexOf('User Email');
  
  const userLogs = allLogs.filter(row => row[userEmailIndex] === email);

  return userLogs.map(row => {
    let logEntry = {};
    headers.forEach((header, index) => {
      logEntry[header] = row[index];
    });
    return logEntry;
  });
}

// =================================================================
// |               Internal Business Logic                         |
// =================================================================

function getEventSettings() {
  const settingsData = settingsSheet.getRange('A2:B' + (settingsSheet.getLastRow() || 2)).getValues();
  const settings = {};
  for (let row of settingsData) {
    if (row[0]) {
      if (row[0] === 'ticketTypes' && row[1]) {
        settings[row[0]] = row[1].toString().split(',').map(item => item.trim());
      } else {
        settings[row[0]] = row[1];
      }
    }
  }
  return settings;
}

function getTickets() {
  if (ticketsSheet.getLastRow() < 2) return [];
  const numCols = Math.max(8, ticketsSheet.getLastColumn());
  const data = ticketsSheet.getRange(2, 1, ticketsSheet.getLastRow() - 1, numCols).getValues();
  return data.map(row => ({
    ticketNumber: row[0], ticketType: row[1], scanStatus: row[2], generatedAt: row[3],
    scannedAt: row[4], scannedBy: row[5], pdfUrl: row[6], qrCodeUrl: row[7], guestName: row[8] || null
  }));
}

function getUsers() {
    if (usersSheet.getLastRow() < 2) return [];
    const data = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 2).getValues();
    return data.filter(row => row[0]).map(row => ({ email: row[0], role: row[1] }));
}

function getDashboardStats(tickets, allTicketTypes) {
  const totalGenerated = tickets.length;
  const totalScanned = tickets.filter(t => t.scanStatus === 'Scanned').length;
  
  const breakdown = {};
  (allTicketTypes || []).forEach(type => {
    breakdown[type] = { generated: 0, scanned: 0 };
  });

  tickets.forEach(t => {
    if (!breakdown[t.ticketType]) {
        breakdown[t.ticketType] = { generated: 0, scanned: 0 };
    }
    breakdown[t.ticketType].generated++;
    if (t.scanStatus === 'Scanned') {
      breakdown[t.ticketType].scanned++;
    }
  });

  return {
    totalGenerated: totalGenerated, 
    totalScanned: totalScanned,
    attendanceRate: totalGenerated > 0 ? (totalScanned / totalGenerated) * 100 : 0,
    breakdown: breakdown
  };
}

function processTicketValidation(ticketNumber, scannerId) {
  const user = scannerId || Session.getEffectiveUser().getEmail();
  if (!ticketNumber || ticketNumber.trim() === '') {
    logAction(user, 'SCAN_FAIL', 'QR code is empty or invalid.', '');
    return { status: 'INVALID', message: 'QR code is empty or invalid.' };
  }

  const textFinder = ticketsSheet.createTextFinder(ticketNumber);
  const found = textFinder.findNext();
  if (!found) {
    logAction(user, 'SCAN_FAIL', 'Ticket not found: ' + ticketNumber, ticketNumber);
    return { status: 'INVALID', message: 'This ticket does not exist in the system.' };
  }

  const row = found.getRow();
  const numCols = Math.max(8, ticketsSheet.getLastColumn());
  const ticketData = ticketsSheet.getRange(row, 1, 1, numCols).getValues()[0];
  const ticketInfo = {
      ticketNumber: ticketData[0], ticketType: ticketData[1], scanStatus: ticketData[2],
      scannedAt: ticketData[4], scannedBy: ticketData[5], guestName: ticketData[8] || null
  };

  if (ticketInfo.scanStatus === 'Scanned') {
    logAction(user, 'SCAN_DUPLICATE', 'Duplicate scan for ticket: ' + ticketNumber, ticketNumber);
    return { 
      status: 'INVALID', 
      message: 'Already scanned by ' + ticketInfo.scannedBy + ' at ' + new Date(ticketInfo.scannedAt).toLocaleString() + '.', 
      ticketNumber: ticketInfo.ticketNumber, 
      ticketType: ticketInfo.ticketType, 
      guestName: ticketInfo.guestName 
    };
  }

  ticketsSheet.getRange(row, 3).setValue('Scanned');
  ticketsSheet.getRange(row, 5).setValue(new Date());
  ticketsSheet.getRange(row, 6).setValue(user);
  
  logAction(user, 'SCAN_SUCCESS', 'Successfully scanned ticket: ' + ticketNumber, ticketNumber);
  return { 
    status: 'VALID', 
    message: 'Access Granted. Welcome!', 
    ticketNumber: ticketInfo.ticketNumber, 
    ticketType: ticketInfo.ticketType, 
    guestName: ticketInfo.guestName 
  };
}

function addUser(newUser) {
    usersSheet.appendRow([newUser.email, newUser.role, new Date()]);
    logAction(Session.getEffectiveUser().getEmail(), 'ADD_USER', 'Added user: ' + newUser.email + ' with role ' + newUser.role + '.');
    return newUser;
}

function removeUser(email) {
    const data = usersSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === email) {
            usersSheet.deleteRow(i + 1);
            logAction(Session.getEffectiveUser().getEmail(), 'REMOVE_USER', 'Removed user: ' + email + '.');
            return 'User ' + email + ' removed.';
        }
    }
    throw new Error('User ' + email + ' not found.');
}

function deleteTicket(ticketNumber) {
    const data = ticketsSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === ticketNumber) {
            ticketsSheet.deleteRow(i + 1);
            logAction(Session.getEffectiveUser().getEmail(), 'DELETE_TICKET', 'Deleted ticket: ' + ticketNumber + '.');
            return 'Ticket ' + ticketNumber + ' deleted.';
        }
    }
    throw new Error('Ticket ' + ticketNumber + ' not found.');
}

// =================================================================
// |                      HELPER FUNCTIONS                         |
// =================================================================

function createIndividualTicket(ticketNumber, qrCodeFile, settings) {
  return createTicketPdf(ticketNumber, qrCodeFile, settings, null);
}

function createIndividualTicketWithGuest(ticketNumber, qrCodeFile, settings, guestName) {
  return createTicketPdf(ticketNumber, qrCodeFile, settings, guestName);
}

function createTicketPdf(ticketNumber, qrCodeFile, settings, guestName) {
  const templateId = settings.slidesTemplateId || FALLBACK_SLIDES_TEMPLATE_ID;
  if (!templateId) {
    return 'https://drive.google.com/uc?id=' + qrCodeFile.getId();
  }
  try {
    const ticketFolder = DriveApp.getFolderById(FOLDER_IDS.generatedTickets);
    const templateFile = DriveApp.getFileById(templateId);
    const newTicketSlideFile = templateFile.makeCopy(ticketNumber + '_Ticket', ticketFolder);
    
    const presentation = SlidesApp.openById(newTicketSlideFile.getId());
    const slide = presentation.getSlides()[0];

    slide.replaceAllText('{{TICKET_NUMBER}}', ticketNumber);
    slide.replaceAllText('{{EVENT_NAME}}', settings.eventName || '');
    slide.replaceAllText('{{EVENT_VENUE}}', settings.eventVenue || '');
    slide.replaceAllText('{{EVENT_DATETIME}}', settings.eventDateTime ? new Date(settings.eventDateTime).toLocaleString() : '');
    
    // Replace guest name placeholder if provided
    if (guestName) {
      slide.replaceAllText('{{GUEST_NAME}}', guestName);
    } else {
      slide.replaceAllText('{{GUEST_NAME}}', '');
    }

    const shapes = slide.getShapes();
    const qrPlaceholder = shapes.find(s => s.getText().asString().trim().includes('{{QR_CODE}}'));
    
    if (qrPlaceholder) {
      const qrImage = slide.insertImage(qrCodeFile);
      qrImage.setLeft(qrPlaceholder.getLeft()).setTop(qrPlaceholder.getTop()).setWidth(qrPlaceholder.getWidth()).setHeight(qrPlaceholder.getHeight());
      qrPlaceholder.remove();
    } else {
      slide.insertImage(qrCodeFile);
    }

    presentation.saveAndClose();

    const pdfBlob = newTicketSlideFile.getAs('application/pdf').setName(ticketNumber + '.pdf');
    const pdfFile = ticketFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    newTicketSlideFile.setTrashed(true);

    return 'https://drive.google.com/uc?id=' + pdfFile.getId() + '&export=download';
  } catch(e) {
    logAction('SYSTEM_ERROR', 'PDF_GENERATION_FAIL', e.stack);
    return 'https://drive.google.com/uc?id=' + qrCodeFile.getId();
  }
}

function saveBase64ImageToDrive(base64Data, fileName, folderId) {
  const decoded = Utilities.base64Decode(base64Data.split(',')[1]);
  const blob = Utilities.newBlob(decoded, base64Data.split(',')[0].split(':')[1].split(';')[0], fileName + '_' + new Date().getTime());
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId();
}

function logAction(user, action, details, ticketNumber) {
  ticketNumber = ticketNumber || '';
  try {
    const timestamp = new Date();
    logSheet.appendRow([timestamp, user, action, details]);
    if (action.startsWith('SCAN_')) {
        userLogsSheet.appendRow([timestamp, user, action, ticketNumber, details]);
    }
  } catch (e) {
    console.error("Logging failed: " + e.message);
  }
}