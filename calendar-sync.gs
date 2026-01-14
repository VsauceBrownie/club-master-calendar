/************************************
 * Club Master Calendar Sync (Simplified)
 * Efficient sync for up to 80 clubs with hundreds of events each
 * 
 * SETUP INSTRUCTIONS:
 * 1. Replace YOUR_CALENDAR_ID below with your Google Calendar ID
 * 2. Set up your spreadsheet with the required tabs (see TABS configuration)
 * 3. Add club information to the ClubRegistry tab
 * 4. Run 'Setup Auto Sync' to enable automatic syncing
 ************************************/

// ========== CONFIGURATION - USER CUSTOMIZATION REQUIRED ==========
// TODO: Replace with your actual Google Calendar ID
// Format: 'calendar-id@group.calendar.google.com' or your personal calendar email
const MASTER_CALENDAR_ID = "YOUR_CALENDAR_ID_HERE";

// Maximum execution time in milliseconds (Google Apps Script limit is 6 minutes)
// Reduce if you frequently hit timeout limits
const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;

// ========== SPREADSHEET STRUCTURE CONFIGURATION ==========
// These are the tab names used in your master spreadsheet
// You can customize these names, but ensure they match your actual tabs
const TABS = {
  REGISTRY: "ClubRegistry",     // Tab containing list of clubs and their sheet URLs
  APPROVED: "Approved",        // Tab containing approved events to sync to calendar
  SYNC_STATE: "SyncState"      // Tab tracking sync status for each club
};

// ========== COLUMN CONFIGURATION ==========
// These define the column positions (0-indexed) in each sheet
// Modify these if your sheets have different column layouts
const COLS = {
  // ClubRegistry tab columns: [ClubID, ClubName, ClubSheetURL]
  REGISTRY: { ClubID: 0, ClubName: 1, ClubSheetURL: 2 },
  
  // Individual club Events tab columns: [EventName, Date, StartTime, EndTime, Location, Description, DeleteFlag]
  EVENTS: { EventName: 0, Date: 1, StartTime: 2, EndTime: 3, Location: 4, Description: 5, DeleteFlag: 6 },
  
  // Approved tab columns: [ClubID, ClubName, EventName, Date, Start, End, Location, Description, CalendarEventID, ApprovedTimestamp]
  APPROVED: { ClubID: 0, ClubName: 1, EventName: 2, Date: 3, Start: 4, End: 5, Location: 6, Description: 7, CalendarEventID: 8, ApprovedTimestamp: 9 },
  
  // SyncState tab columns: [ClubID, LastSync, RowCount, Checksum]
  SYNC_STATE: { ClubID: 0, LastSync: 1, RowCount: 2, Checksum: 3 }
};

/* ========== UI Interface ========== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Master Calendar')
    .addItem('Quick Sync (Changed Only)', 'quickSync')
    .addItem('Full Sync All Clubs', 'fullSync')
    .addItem('Setup Auto Sync', 'setupAutoSync')
    .addToUi();
}

function quickSync() {
  const result = syncClubs(false);
  showResult(`Quick Sync: ${result.clubs} clubs, ${result.events} events processed`);
}

function autoQuickSync() {
  const result = syncClubs(false);
  console.log(`Quick Sync: ${result.clubs} clubs, ${result.events} events processed`);
}

function fullSync() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Confirm Full Sync', 'Process all clubs regardless of changes?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    const result = syncClubs(true);
    showResult(`Full Sync: ${result.clubs} clubs, ${result.events} events processed`);
  }
}

function showResult(message) {
  SpreadsheetApp.getUi().alert('Sync Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ========== Main Sync Logic ========== */
function syncClubs(forceAll = false) {
  const startTime = Date.now();
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const calendar = CalendarApp.getCalendarById(MASTER_CALENDAR_ID);
  
  if (!calendar) throw new Error("Calendar not found");
  
  // Get clubs and sync state
  const clubs = getClubs(masterSS);
  const syncState = getSyncState(masterSS);
  const approvedState = getApprovedState(masterSS);
  
  let clubsProcessed = 0;
  let eventsProcessed = 0;
  const newSyncState = [];
  
  for (const club of clubs) {
    if (Date.now() - startTime > MAX_EXECUTION_TIME) break;
    
    // Check if club needs syncing
    const lastSync = syncState[club.id];
    if (!forceAll && !needsSync(club, lastSync)) continue;
    
    try {
      const result = syncOneClub(calendar, masterSS, approvedState, club);
      eventsProcessed += result.eventsProcessed;
      clubsProcessed++;
      
      // Update sync state
      newSyncState.push({
        clubId: club.id,
        lastSync: new Date(),
        rowCount: result.rowCount,
        checksum: result.checksum
      });
      
    } catch (error) {
      console.log(`Error syncing club ${club.id}: ${error}`);
    }
  }
  
  // Save new sync state
  updateSyncState(masterSS, newSyncState);
  
  return { clubs: clubsProcessed, events: eventsProcessed };
}

function syncOneClub(calendar, masterSS, approvedState, club) {
  const clubSS = SpreadsheetApp.openByUrl(club.url);
  const eventsSheet = getSheet(clubSS, "Events");
  const events = getSheetData(eventsSheet);
  
  const approvedSheet = getSheet(masterSS, TABS.APPROVED);
  let eventsProcessed = 0;
  const toDelete = [];
  const toAdd = [];
  const toUpdate = [];
  
  // Process each event
  for (const row of events) {
    const event = parseEvent(row, club);
    if (!event) continue;
    
    const existing = approvedState[event.key];
    eventsProcessed++;
    
    if (event.deleteFlag) {
      if (existing) {
        deleteCalendarEvent(calendar, existing.eventId, event);
        toDelete.push(existing.rowIndex);
        delete approvedState[event.key];
      }
    } else if (!existing) {
      const calEvent = createCalendarEvent(calendar, event);
      const newRow = createApprovedRow(event, calEvent.getId());
      toAdd.push(newRow);
      approvedState[event.key] = { eventId: calEvent.getId() };
    } else if (eventChanged(existing.row, event)) {
      updateCalendarEvent(calendar, existing.eventId, event);
      const updatedRow = createApprovedRow(event, existing.eventId);
      toUpdate.push({ rowIndex: existing.rowIndex, row: updatedRow });
    }
  }
  
  // Batch apply changes to sheet
  applySheetChanges(approvedSheet, toDelete, toAdd, toUpdate);
  
  return {
    eventsProcessed,
    rowCount: events.length,
    checksum: calculateChecksum(events)
  };
}

/* ========== Change Detection ========== */
function needsSync(club, lastSyncData) {
  if (!lastSyncData) return true;
  
  try {
    const clubSS = SpreadsheetApp.openByUrl(club.url);
    const eventsSheet = getSheet(clubSS, "Events");
    const currentRowCount = eventsSheet.getLastRow() - 1; // Exclude header
    
    // Quick check: different row count
    if (currentRowCount !== lastSyncData.rowCount) return true;
    
    // Deep check: data changed
    const events = getSheetData(eventsSheet);
    const currentChecksum = calculateChecksum(events);
    return currentChecksum !== lastSyncData.checksum;
    
  } catch (error) {
    return true; // Sync if we can't check
  }
}

function calculateChecksum(data) {
  const content = JSON.stringify(data);
  let hash = 0;
  for (let i = 0; i < content.length; i++) {
    hash = ((hash << 5) - hash) + content.charCodeAt(i);
    hash = hash & hash;
  }
  return hash.toString();
}

function parseEvent(row, club) {
  const name = getString(row[COLS.EVENTS.EventName]);
  const date = row[COLS.EVENTS.Date];
  const startTime = row[COLS.EVENTS.StartTime];
  const endTime = row[COLS.EVENTS.EndTime] || startTime;
  
  if (!name || !date || !startTime) return null;
  
  const location = getString(row[COLS.EVENTS.Location]);
  const description = getString(row[COLS.EVENTS.Description]);
  const deleteFlag = getString(row[COLS.EVENTS.DeleteFlag]).toLowerCase() === 'yes';
  
  const startDT = combineDateTime(date, startTime);
  const endDT = combineDateTime(date, endTime);
  const key = `${club.id}||${name}||${formatDate(date)}||${formatTime(startTime)}`;
  
  return {
    key,
    clubId: club.id,
    clubName: club.name,
    name,
    location,
    description,
    deleteFlag,
    startDT,
    endDT,
    title: `${club.name} - ${name}`
  };
}

function createCalendarEvent(calendar, event) {
  return calendar.createEvent(event.title, event.startDT, event.endDT, {
    location: event.location || undefined,
    description: event.description || undefined
  });
}

function updateCalendarEvent(calendar, eventId, event) {
  try {
    const calEvent = CalendarApp.getEventById(eventId);
    if (calEvent) {
      calEvent.setTitle(event.title);
      calEvent.setTime(event.startDT, event.endDT);
      calEvent.setLocation(event.location || "");
      calEvent.setDescription(event.description || "");
    }
  } catch (error) {
    console.log(`Failed to update event ${eventId}: ${error}`);
  }
}

function deleteCalendarEvent(calendar, eventId, event) {
  try {
    const calEvent = CalendarApp.getEventById(eventId);
    if (calEvent) {
      calEvent.deleteEvent();
      return;
    }
  } catch (error) {
    console.log(`Failed to delete by ID ${eventId}, trying search: ${error}`);
  }
  
  // Fallback: search and delete
  try {
    const events = calendar.getEvents(event.startDT, event.endDT, { search: event.name });
    events.forEach(e => {
      if (e.getTitle() === event.title) e.deleteEvent();
    });
  } catch (error) {
    console.log(`Fallback delete failed: ${error}`);
  }
}

/* ========== Sheet Operations ========== */
function applySheetChanges(sheet, toDelete, toAdd, toUpdate) {
  // Delete rows (from bottom up to avoid index shifting)
  toDelete.sort((a, b) => b - a).forEach(rowIndex => {
    try { sheet.deleteRow(rowIndex); } catch (e) { console.log(`Delete row error: ${e}`); }
  });
  
  // Add new rows
  if (toAdd.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAdd.length, toAdd[0].length).setValues(toAdd);
  }
  
  // Update existing rows
  toUpdate.forEach(update => {
    try {
      sheet.getRange(update.rowIndex, 1, 1, update.row.length).setValues([update.row]);
    } catch (e) { console.log(`Update row error: ${e}`); }
  });
}

function createApprovedRow(event, eventId) {
  return [
    event.clubId, event.clubName, event.name,
    formatDate(event.startDT), formatTime(event.startDT), formatTime(event.endDT),
    event.location, event.description, eventId, new Date()
  ];
}

function eventChanged(existingRow, newEvent) {
  return (
    existingRow[COLS.APPROVED.EventName] !== newEvent.name ||
    existingRow[COLS.APPROVED.Location] !== newEvent.location ||
    existingRow[COLS.APPROVED.Description] !== newEvent.description ||
    existingRow[COLS.APPROVED.Start] !== formatTime(newEvent.startDT) ||
    existingRow[COLS.APPROVED.End] !== formatTime(newEvent.endDT)
  );
}

/* ========== Data Access Helpers ========== */
function getClubs(masterSS) {
  const sheet = getSheet(masterSS, TABS.REGISTRY);
  return getSheetData(sheet).map(row => ({
    id: getString(row[COLS.REGISTRY.ClubID]),
    name: getString(row[COLS.REGISTRY.ClubName]),
    url: getString(row[COLS.REGISTRY.ClubSheetURL])
  })).filter(club => club.id && club.name && club.url);
}

function getSyncState(masterSS) {
  const sheet = getOrCreateSheet(masterSS, TABS.SYNC_STATE, ['ClubID', 'LastSync', 'RowCount', 'Checksum']);
  const data = getSheetData(sheet);
  const state = {};
  
  data.forEach(row => {
    const clubId = getString(row[COLS.SYNC_STATE.ClubID]);
    if (clubId) {
      state[clubId] = {
        lastSync: row[COLS.SYNC_STATE.LastSync],
        rowCount: parseInt(row[COLS.SYNC_STATE.RowCount] || "0"),
        checksum: getString(row[COLS.SYNC_STATE.Checksum])
      };
    }
  });
  
  return state;
}

function getApprovedState(masterSS) {
  const sheet = getSheet(masterSS, TABS.APPROVED);
  const data = getSheetData(sheet);
  const state = {};
  
  data.forEach((row, index) => {
    const clubId = getString(row[COLS.APPROVED.ClubID]);
    const name = getString(row[COLS.APPROVED.EventName]);
    if (clubId && name) {
      const key = `${clubId}||${name}||${row[COLS.APPROVED.Date]}||${row[COLS.APPROVED.Start]}`;
      state[key] = {
        rowIndex: index + 2, // +2 for header row
        row: row,
        eventId: getString(row[COLS.APPROVED.CalendarEventID])
      };
    }
  });
  
  return state;
}

function updateSyncState(masterSS, updates) {
  const sheet = getOrCreateSheet(masterSS, TABS.SYNC_STATE, ['ClubID', 'LastSync', 'RowCount', 'Checksum']);
  const existing = getSheetData(sheet);
  const existingMap = {};
  
  existing.forEach((row, index) => {
    const clubId = getString(row[COLS.SYNC_STATE.ClubID]);
    if (clubId) existingMap[clubId] = index + 2;
  });
  
  const toAdd = [];
  const toUpdate = [];
  
  updates.forEach(update => {
    const newRow = [update.clubId, update.lastSync, update.rowCount, update.checksum];
    const existingRowIndex = existingMap[update.clubId];
    
    if (existingRowIndex) {
      toUpdate.push({ rowIndex: existingRowIndex, row: newRow });
    } else {
      toAdd.push(newRow);
    }
  });
  
  // Apply updates
  toUpdate.forEach(update => {
    sheet.getRange(update.rowIndex, 1, 1, 4).setValues([update.row]);
  });
  
  if (toAdd.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAdd.length, 4).setValues(toAdd);
  }
}

/* ========== Utility Functions ========== */
function getSheet(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" not found`);
  return sheet;
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function getSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

function getString(value) {
  return (value || "").toString().trim();
}

function combineDateTime(date, time) {
  const baseDate = date instanceof Date ? new Date(date) : new Date(date);
  const timeObj = parseTime(time);
  baseDate.setHours(timeObj.hours, timeObj.minutes, 0, 0);
  return baseDate;
}

function parseTime(timeValue) {
  if (timeValue instanceof Date) {
    return { hours: timeValue.getHours(), minutes: timeValue.getMinutes() };
  }
  
  const timeStr = timeValue.toString().trim();
  const ampmMatch = timeStr.match(/^(\d{1,2}):(\d{2})\s*([APap][Mm])$/);
  
  if (ampmMatch) {
    let hours = parseInt(ampmMatch[1]);
    const minutes = parseInt(ampmMatch[2]);
    const ampm = ampmMatch[3].toUpperCase();
    
    if (ampm === "PM" && hours !== 12) hours += 12;
    if (ampm === "AM" && hours === 12) hours = 0;
    
    return { hours, minutes };
  }
  
  const regularMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (regularMatch) {
    return { hours: parseInt(regularMatch[1]), minutes: parseInt(regularMatch[2]) };
  }
  
  throw new Error(`Invalid time format: ${timeStr}`);
}

function formatDate(date) {
  const d = date instanceof Date ? date : new Date(date);
  return `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}-${d.getDate().toString().padStart(2, '0')}`;
}

function formatTime(date) {
  const d = date instanceof Date ? date : new Date(date);
  return `${d.getHours().toString().padStart(2, '0')}:${d.getMinutes().toString().padStart(2, '0')}`;
}

/* ========== Auto Sync Setup ========== */
function setupAutoSync() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger for every 10 minutes
  ScriptApp.newTrigger('autoSync')
    .timeBased()
    .everyMinutes(10)
    .create();
    
  SpreadsheetApp.getUi().alert('Auto Sync Setup', 'Automatic sync will run every 10 minutes', SpreadsheetApp.getUi().ButtonSet.OK);
}

function autoSync() {
  try {
    syncClubs(false); // Quick sync only
  } catch (error) {
    console.log(`Auto sync error: ${error}`);
  }
}

/* ========== FIXED: Improved Key Generation ========== */
function parseEvent(row, club) {
  const name = getString(row[COLS.EVENTS.EventName]);
  const dateCell = row[COLS.EVENTS.Date];
  const startCell = row[COLS.EVENTS.StartTime];
  const endCell = row[COLS.EVENTS.EndTime] || startCell;
  
  if (!name || !dateCell || !startCell) return null;
  
  const location = getString(row[COLS.EVENTS.Location]);
  const description = getString(row[COLS.EVENTS.Description]);
  const deleteFlag = getString(row[COLS.EVENTS.DeleteFlag]).toLowerCase() === 'yes';
  
  const startDT = combineDateTime(dateCell, startCell);
  const endDT = combineDateTime(dateCell, endCell);
  
  // FIXED: Ensure consistent key generation
  const normalizedDate = formatDate(startDT);  // Use the DateTime object for consistency
  const normalizedTime = formatTime(startDT);  // Use the DateTime object for consistency
  const key = `${club.id}||${name}||${normalizedDate}||${normalizedTime}`;
  
  return {
    key,
    clubId: club.id,
    clubName: club.name,
    name,
    location,
    description,
    deleteFlag,
    startDT,
    endDT,
    title: `${club.name} - ${name}`
  };
}

/* ========== IMPROVED: Approved State Loading with Duplicate Detection ========== */
function getApprovedState(masterSS) {
  const sheet = getSheet(masterSS, TABS.APPROVED);
  const data = getSheetData(sheet);
  const state = {};
  
  data.forEach((row, index) => {
    const clubId = getString(row[COLS.APPROVED.ClubID]);
    const name = getString(row[COLS.APPROVED.EventName]);
    const dateStr = row[COLS.APPROVED.Date];
    const startStr = row[COLS.APPROVED.Start];
    
    if (clubId && name && dateStr && startStr) {
      // FIXED: Ensure consistent key generation using stored values
      const normalizedDate = typeof dateStr === 'string' ? dateStr : formatDate(dateStr);
      const normalizedTime = typeof startStr === 'string' ? startStr : formatTime(startStr);
      const key = `${clubId}||${name}||${normalizedDate}||${normalizedTime}`;
      
      // FIXED: Check for duplicate keys and log them
      if (state[key]) {
        console.log(`WARNING: Duplicate key found: ${key}`);
        console.log(`Existing row: ${state[key].rowIndex}, New row: ${index + 2}`);
      }
      
      state[key] = {
        rowIndex: index + 2, // +2 for header row
        row: row,
        eventId: getString(row[COLS.APPROVED.CalendarEventID])
      };
    }
  });
  
  console.log(`Loaded ${Object.keys(state).length} approved events`);
  return state;
}

/* ========== DUPLICATE CLEANUP FUNCTION ========== */
function cleanupDuplicates() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Cleanup Duplicates',
    'This will remove duplicate events from the Approved sheet and calendar. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const approvedSheet = getSheet(masterSS, TABS.APPROVED);
  const calendar = CalendarApp.getCalendarById(MASTER_CALENDAR_ID);
  
  const data = getSheetData(approvedSheet);
  const seen = new Map();
  const toDelete = [];
  
  data.forEach((row, index) => {
    const clubId = getString(row[COLS.APPROVED.ClubID]);
    const name = getString(row[COLS.APPROVED.EventName]);
    const dateStr = row[COLS.APPROVED.Date];
    const startStr = row[COLS.APPROVED.Start];
    const eventId = getString(row[COLS.APPROVED.CalendarEventID]);
    
    if (clubId && name && dateStr && startStr) {
      const key = `${clubId}||${name}||${dateStr}||${startStr}`;
      
      if (seen.has(key)) {
        // This is a duplicate
        console.log(`Found duplicate: ${key} in row ${index + 2}`);
        toDelete.push({
          rowIndex: index + 2,
          eventId: eventId,
          eventName: name
        });
      } else {
        seen.set(key, index + 2);
      }
    }
  });
  
  console.log(`Found ${toDelete.length} duplicates to remove`);
  
  // Delete calendar events and sheet rows
  toDelete.forEach(duplicate => {
    // Delete from calendar
    try {
      if (duplicate.eventId) {
        const calEvent = CalendarApp.getEventById(duplicate.eventId);
        if (calEvent) {
          calEvent.deleteEvent();
          console.log(`Deleted calendar event: ${duplicate.eventId}`);
        }
      }
    } catch (error) {
      console.log(`Error deleting calendar event ${duplicate.eventId}: ${error}`);
    }
  });
  
  // Delete sheet rows (from bottom up)
  const rowsToDelete = toDelete.map(d => d.rowIndex).sort((a, b) => b - a);
  rowsToDelete.forEach(rowIndex => {
    try {
      approvedSheet.deleteRow(rowIndex);
      console.log(`Deleted row: ${rowIndex}`);
    } catch (error) {
      console.log(`Error deleting row ${rowIndex}: ${error}`);
    }
  });
  
  ui.alert('Cleanup Complete', `Removed ${toDelete.length} duplicate events`, ui.ButtonSet.OK);
}

/* ========== FIXED: Add to UI Menu ========== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Master Calendar')
    .addItem('Quick Sync (Changed Only)', 'quickSync')
    .addItem('Full Sync All Clubs', 'fullSync')
    .addItem('Setup Auto Sync', 'setupAutoSync')
    .addSeparator()
    .addItem('Cleanup Duplicates', 'cleanupDuplicates')  // NEW
    .addItem('Debug Sync State', 'debugSyncState')      // NEW
    .addToUi();
}

/* ========== DEBUGGING FUNCTION ========== */
function debugSyncState() {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log("=== DEBUGGING SYNC STATE ===");
  
  // Check approved events
  const approvedState = getApprovedState(masterSS);
  console.log(`Total approved events: ${Object.keys(approvedState).length}`);
  
  // Look for potential duplicates
  const eventCounts = {};
  Object.keys(approvedState).forEach(key => {
    const parts = key.split('||');
    const eventIdentity = `${parts[0]}||${parts[1]}`; // Club + Event name
    eventCounts[eventIdentity] = (eventCounts[eventIdentity] || 0) + 1;
  });
  
  console.log("=== POTENTIAL DUPLICATES ===");
  Object.entries(eventCounts).forEach(([identity, count]) => {
    if (count > 1) {
      console.log(`${identity}: ${count} instances`);
    }
  });
  
  // Check sync state
  const syncState = getSyncState(masterSS);
  console.log(`Clubs in sync state: ${Object.keys(syncState).length}`);
}
/* ========== NUCLEAR RESET: Clear Calendar and Repopulate ========== */
function resetCalendarFromApproved() {
  const ui = SpreadsheetApp.getUi();
  
  // Multi-step confirmation to prevent accidents
  const response1 = ui.alert(
    '⚠️ DESTRUCTIVE OPERATION WARNING ⚠️',
    'This will DELETE ALL events from the master calendar and rebuild from the Approved sheet.\n\n' +
    'This action cannot be undone!\n\n' +
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response1 !== ui.Button.YES) return;
  
  const response2 = ui.alert(
    'Final Confirmation',
    'Type YES in the next dialog to confirm you want to DELETE ALL CALENDAR EVENTS',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response2 !== ui.Button.OK) return;
  
  const confirmation = ui.prompt(
    'Type YES to Proceed',
    'Type "YES" (without quotes) to confirm calendar reset:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (confirmation.getSelectedButton() !== ui.Button.OK || 
      confirmation.getResponseText().trim() !== 'YES') {
    ui.alert('Reset Cancelled', 'Calendar reset was cancelled.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const result = executeCalendarReset();
    ui.alert(
      'Reset Complete ✅',
      `Calendar successfully reset!\n\n` +
      `• Deleted: ${result.deleted} old events\n` +
      `• Created: ${result.created} new events\n` +
      `• Updated: ${result.updated} calendar IDs in Approved sheet`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Reset Failed ❌', `Error during reset: ${error.toString()}`, ui.ButtonSet.OK);
    console.error('Calendar reset error:', error);
  }
}

function executeCalendarReset() {
  const startTime = Date.now();
  console.log('🚀 Starting calendar reset...');
  
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const calendar = CalendarApp.getCalendarById(MASTER_CALENDAR_ID);
  const approvedSheet = getSheet(masterSS, TABS.APPROVED);
  
  if (!calendar) {
    throw new Error('Master calendar not found');
  }
  
  // Step 1: Clear entire calendar
  const deletedCount = clearEntireCalendar(calendar);
  console.log(`🗑️ Deleted ${deletedCount} events from calendar`);
  
  // Step 2: Rebuild from approved sheet
  const approvedData = getSheetData(approvedSheet);
  const results = rebuildCalendarFromApproved(calendar, approvedSheet, approvedData);
  
  const duration = ((Date.now() - startTime) / 1000).toFixed(1);
  console.log(`✅ Reset complete in ${duration}s`);
  
  return {
    deleted: deletedCount,
    created: results.created,
    updated: results.updated
  };
}

/* ========== CALENDAR CLEARING ========== */
function clearEntireCalendar(calendar) {
  console.log('🧹 Clearing entire calendar...');
  
  let deletedCount = 0;
  const batchSize = 50; // Process in batches to avoid timeouts
  
  // Get events in chunks (Google Calendar API limits)
  const today = new Date();
  const futureDate = new Date(today.getTime() + (365 * 24 * 60 * 60 * 1000)); // 1 year ahead
  const pastDate = new Date(today.getTime() - (365 * 24 * 60 * 60 * 1000)); // 1 year back
  
  // Clear future events
  deletedCount += clearDateRange(calendar, today, futureDate, 'future');
  
  // Clear past events  
  deletedCount += clearDateRange(calendar, pastDate, today, 'past');
  
  return deletedCount;
}

function clearDateRange(calendar, startDate, endDate, label) {
  console.log(`🗑️ Clearing ${label} events from ${startDate.toDateString()} to ${endDate.toDateString()}`);
  
  let deletedCount = 0;
  let batchCount = 0;
  const maxBatches = 20; // Safety limit
  
  while (batchCount < maxBatches) {
    try {
      const events = calendar.getEvents(startDate, endDate);
      
      if (events.length === 0) {
        console.log(`✅ No more ${label} events to delete`);
        break;
      }
      
      console.log(`🔄 Batch ${batchCount + 1}: Deleting ${events.length} ${label} events`);
      
      events.forEach(event => {
        try {
          event.deleteEvent();
          deletedCount++;
        } catch (error) {
          console.log(`⚠️ Error deleting event "${event.getTitle()}": ${error}`);
        }
      });
      
      batchCount++;
      
      // Brief pause to avoid rate limits
      if (events.length > 10) {
        Utilities.sleep(1000);
      }
      
    } catch (error) {
      console.log(`❌ Error in batch ${batchCount}: ${error}`);
      break;
    }
  }
  
  console.log(`🗑️ Deleted ${deletedCount} ${label} events in ${batchCount} batches`);
  return deletedCount;
}

/* ========== CALENDAR REBUILDING ========== */
function rebuildCalendarFromApproved(calendar, approvedSheet, approvedData) {
  console.log('🔨 Rebuilding calendar from approved events...');
  
  let created = 0;
  let updated = 0;
  const batchSize = 25;
  const updateRanges = [];
  
  // Process in batches to avoid timeouts
  for (let i = 0; i < approvedData.length; i += batchSize) {
    const batch = approvedData.slice(i, i + batchSize);
    console.log(`🔄 Processing batch ${Math.floor(i/batchSize) + 1}: rows ${i+2} to ${Math.min(i+batchSize+1, approvedData.length+1)}`);
    
    batch.forEach((row, batchIndex) => {
      const globalIndex = i + batchIndex;
      const rowNumber = globalIndex + 2; // +2 for header
      
      try {
        const eventData = parseApprovedRow(row);
        if (!eventData) {
          console.log(`⚠️ Skipping invalid row ${rowNumber}`);
          return;
        }
        
        // Create new calendar event
        const calendarEvent = calendar.createEvent(
          eventData.title,
          eventData.startDT,
          eventData.endDT,
          {
            location: eventData.location || undefined,
            description: eventData.description || undefined
          }
        );
        
        const newEventId = calendarEvent.getId();
        console.log(`✅ Created: "${eventData.title}" -> ${newEventId}`);
        created++;
        
        // Prepare sheet update with new event ID
        const updatedRow = [...row];
        updatedRow[COLS.APPROVED.CalendarEventID] = newEventId;
        updatedRow[COLS.APPROVED.ApprovedTimestamp] = new Date();
        
        updateRanges.push({
          range: approvedSheet.getRange(rowNumber, 1, 1, updatedRow.length),
          values: [updatedRow]
        });
        
        updated++;
        
      } catch (error) {
        console.log(`❌ Error processing row ${rowNumber}: ${error}`);
      }
    });
    
    // Brief pause between batches
    if (i + batchSize < approvedData.length) {
      Utilities.sleep(500);
    }
  }
  
  // Batch update all calendar IDs in the sheet
  console.log('📝 Updating calendar IDs in approved sheet...');
  updateRanges.forEach(update => {
    try {
      update.range.setValues(update.values);
    } catch (error) {
      console.log(`⚠️ Error updating sheet: ${error}`);
    }
  });
  
  return { created, updated };
}

/* ========== APPROVED ROW PARSER ========== */
function parseApprovedRow(row) {
  const clubId = getString(row[COLS.APPROVED.ClubID]);
  const clubName = getString(row[COLS.APPROVED.ClubName]);
  const eventName = getString(row[COLS.APPROVED.EventName]);
  const dateCell = row[COLS.APPROVED.Date];
  const startCell = row[COLS.APPROVED.Start];
  const endCell = row[COLS.APPROVED.End];
  const location = getString(row[COLS.APPROVED.Location]);
  const description = getString(row[COLS.APPROVED.Description]);
  
  // Validate required fields
  if (!clubId || !clubName || !eventName || !dateCell || !startCell) {
    return null;
  }
  
  try {
    const startDT = combineDateTime(dateCell, startCell);
    const endDT = combineDateTime(dateCell, endCell || startCell);
    
    return {
      clubId,
      clubName,
      eventName,
      location,
      description,
      startDT,
      endDT,
      title: `${clubName} - ${eventName}`
    };
  } catch (error) {
    console.log(`Error parsing date/time for event "${eventName}": ${error}`);
    return null;
  }
}

/* ========== SAFER ALTERNATIVE: SELECTIVE RESET ========== */
function resetCalendarSelectiveByDate() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Selective Calendar Reset',
    'Enter date range to reset (leave empty for all events):\n\n' +
    'Start date (YYYY-MM-DD) or leave empty:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const startDateStr = response.getResponseText().trim();
  let startDate = null;
  let endDate = null;
  
  if (startDateStr) {
    const endResponse = ui.prompt(
      'End Date',
      'End date (YYYY-MM-DD):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (endResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const endDateStr = endResponse.getResponseText().trim();
    
    try {
      startDate = new Date(startDateStr);
      endDate = new Date(endDateStr);
      
      if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        throw new Error('Invalid date format');
      }
    } catch (error) {
      ui.alert('Invalid Date', 'Please use YYYY-MM-DD format', ui.ButtonSet.OK);
      return;
    }
  }
  
  try {
    const result = executeSelectiveReset(startDate, endDate);
    ui.alert(
      'Selective Reset Complete',
      `Reset events ${startDate ? `from ${startDate.toDateString()} to ${endDate.toDateString()}` : 'for all dates'}\n\n` +
      `• Processed: ${result.processed} approved events\n` +
      `• Created: ${result.created} calendar events`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Reset Failed', error.toString(), ui.ButtonSet.OK);
  }
}

function executeSelectiveReset(startDate, endDate) {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const calendar = CalendarApp.getCalendarById(MASTER_CALENDAR_ID);
  const approvedSheet = getSheet(masterSS, TABS.APPROVED);
  const approvedData = getSheetData(approvedSheet);
  
  let processed = 0;
  let created = 0;
  
  // Clear calendar events in date range
  if (startDate && endDate) {
    console.log(`Clearing calendar events from ${startDate.toDateString()} to ${endDate.toDateString()}`);
    clearDateRange(calendar, startDate, endDate, 'selected range');
  }
  
  // Rebuild events in range
  approvedData.forEach((row, index) => {
    const eventData = parseApprovedRow(row);
    if (!eventData) return;
    
    // Check if event is in date range (if specified)
    if (startDate && endDate) {
      const eventDate = new Date(eventData.startDT);
      if (eventDate < startDate || eventDate > endDate) {
        return; // Skip events outside range
      }
    }
    
    processed++;
    
    try {
      const calendarEvent = calendar.createEvent(
        eventData.title,
        eventData.startDT,
        eventData.endDT,
        {
          location: eventData.location || undefined,
          description: eventData.description || undefined
        }
      );
      
      // Update sheet with new event ID
      const rowNumber = index + 2;
      approvedSheet.getRange(rowNumber, COLS.APPROVED.CalendarEventID + 1).setValue(calendarEvent.getId());
      
      created++;
    } catch (error) {
      console.log(`Error recreating event "${eventData.title}": ${error}`);
    }
  });
  
  return { processed, created };
}

/* ========== ADD TO UI MENU ========== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Master Calendar')
    .addItem('Quick Sync (Changed Only)', 'quickSync')
    .addItem('Full Sync All Clubs', 'fullSync')
    .addItem('Setup Auto Sync', 'setupAutoSync')
    .addSeparator()
    .addItem('Cleanup Duplicates', 'cleanupDuplicates')
    .addItem('🚨 Reset Calendar (Nuclear)', 'resetCalendarFromApproved')     // NEW
    .addItem('🎯 Reset Calendar (Date Range)', 'resetCalendarSelectiveByDate') // NEW
    .addItem('Debug Sync State', 'debugSyncState')
    .addToUi();
}
