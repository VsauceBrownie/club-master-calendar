# Club Master Calendar Sync

A powerful Google Apps Script solution for synchronizing events from multiple club spreadsheets to a single master Google Calendar. Perfect for organizations, schools, or communities managing events across multiple clubs or departments.

## 🚀 Features

- **Multi-Club Support**: Sync events from up to 80+ clubs efficiently
- **Change Detection**: Only processes clubs with changes to save time
- **Automatic Sync**: Configurable automatic syncing every 10 minutes
- **Duplicate Management**: Built-in duplicate detection and cleanup
- **Calendar Reset Options**: Safe reset functions for maintenance
- **Comprehensive UI**: Easy-to-use menu interface in Google Sheets
- **Error Handling**: Robust error handling and logging
- **Batch Processing**: Optimized for Google Apps Script limits

## 📋 Prerequisites

- Google Account with access to Google Sheets and Google Calendar
- Basic understanding of spreadsheet structure
- Permission to create and manage Google Calendars

## 🔧 Complete Setup Guide

### Step 1: Create Your Master Calendar

1. **Create a New Google Calendar**
   - Go to [Google Calendar](https://calendar.google.com)
   - Click the "+" icon next to "Other calendars"
   - Select "Create new calendar"
   - Name it something like "Master Club Calendar" or "Organization Events"
   - Set appropriate sharing permissions (make sure it's accessible to your script)

2. **Get Your Calendar ID**
   - In Google Calendar, find your calendar in the left sidebar
   - Hover over it and click the three dots (⋮)
   - Select "Settings and sharing"
   - Scroll down to "Integrate calendar"
   - Copy the **Calendar ID** (format: `calendar-id@group.calendar.google.com` or your email)

### Step 2: Set Up Your Master Spreadsheet

1. **Create a New Google Sheet**
   - Go to [Google Sheets](https://sheets.google.com)
   - Create a new spreadsheet named "Master Club Calendar" or similar

2. **Create Required Tabs**
   Create these exact tabs (case-sensitive):

   #### 📊 ClubRegistry Tab
   | Column A | Column B | Column C |
   |----------|----------|----------|
   | ClubID | ClubName | ClubSheetURL |
   
   **Example:**
   ```
   club001    | Drama Club    | https://docs.google.com/spreadsheets/d/.../edit
   club002    | Science Club  | https://docs.google.com/spreadsheets/d/.../edit
   club003    | Music Club    | https://docs.google.com/spreadsheets/d/.../edit
   ```

   #### ✅ Approved Tab (Leave empty - will be auto-populated)
   | Column A | Column B | Column C | Column D | Column E | Column F | Column G | Column H | Column I | Column J |
   |----------|----------|----------|----------|----------|----------|----------|----------|----------|-----------|
   | ClubID | ClubName | EventName | Date | Start | End | Location | Description | CalendarEventID | ApprovedTimestamp |

   #### 📈 SyncState Tab (Leave empty - will be auto-created)
   | Column A | Column B | Column C | Column D |
   |----------|----------|----------|----------|
   | ClubID | LastSync | RowCount | Checksum |

### Step 3: Set Up Individual Club Sheets

For each club listed in your ClubRegistry:

1. **Create or Use Existing Club Spreadsheet**
   - Each club needs their own Google Sheet
   - Ensure the sheet is shared with "Anyone with the link can view" (or appropriate permissions)

2. **Create Events Tab**
   Create an "Events" tab with these exact columns:

   | Column A | Column B | Column C | Column D | Column E | Column F | Column G |
   |----------|----------|----------|----------|----------|----------|----------|
   | EventName | Date | StartTime | EndTime | Location | Description | DeleteFlag |

   **Example Events:**
   ```
   Weekly Meeting | 2024-01-15 | 3:00 PM | 4:00 PM | Room 101 | Regular weekly meeting | 
   Spring Play | 2024-03-20 | 7:00 PM | 9:30 PM | Auditorium | Annual spring performance | 
   Fundraiser | 2024-02-10 | 10:00 AM | 2:00 PM | Cafeteria | Bake sale and raffle | 
   Old Event | 2024-01-01 | 5:00 PM | 6:00 PM | Room 205 | Cancelled event | yes
   ```

   **Important Notes:**
   - **Date format**: YYYY-MM-DD (e.g., 2024-01-15)
   - **Time format**: HH:MM AM/PM (e.g., 3:00 PM) or 24-hour (e.g., 15:00)
   - **DeleteFlag**: Enter "yes" to mark an event for deletion
   - **EndTime**: Optional, defaults to StartTime if empty

### Step 4: Install and Configure the Apps Script

1. **Open the Script Editor**
   - In your master spreadsheet, go to `Extensions` > `Apps Script`
   - Delete any existing code
   - Copy and paste the entire Apps Script code

2. **Configure Your Calendar ID**
   - Find this line at the top of the script:
   ```javascript
   const MASTER_CALENDAR_ID = "YOUR_CALENDAR_ID_HERE";
   ```
   - Replace `YOUR_CALENDAR_ID_HERE` with your actual calendar ID from Step 1

3. **Save and Authorize**
   - Click the Save project icon (💾)
   - Give the project a name (e.g., "Club Master Calendar Sync")
   - Close and reopen the spreadsheet to trigger the authorization flow
   - When prompted, grant permissions for:
     - Spreadsheet access (to read/write data)
     - Calendar access (to create/update/delete events)

### Step 5: Initial Setup and Testing

1. **Verify Menu Appears**
   - Reopen your spreadsheet
   - You should see a "Master Calendar" menu in the toolbar

2. **Populate Club Registry**
   - Add all your clubs to the ClubRegistry tab
   - Ensure all sheet URLs are correct and accessible

3. **Run Initial Sync**
   - Click `Master Calendar` > `Full Sync All Clubs`
   - This will populate the Approved tab with all events
   - Check your Google Calendar to verify events appeared

4. **Set Up Automatic Sync**
   - Click `Master Calendar` > `Setup Auto Sync`
   - This creates a trigger to run every 10 minutes

## 🎯 Menu Functions Explained

### 📅 Sync Operations
- **Quick Sync (Changed Only)**: Only processes clubs with detected changes
- **Full Sync All Clubs**: Processes all clubs regardless of changes (slower but thorough)

### 🛠️ Maintenance Tools
- **Setup Auto Sync**: Creates automatic trigger (runs every 10 minutes)
- **Cleanup Duplicates**: Removes duplicate events from calendar and sheet
- **🚨 Reset Calendar (Nuclear)**: Deletes ALL calendar events and rebuilds from Approved tab
- **🎯 Reset Calendar (Date Range)**: Selective reset for specific date ranges

### 🔍 Debugging
- **Debug Sync State**: Shows diagnostic information in the console

## ⚠️ Important Notes

### Safety First
- **Always test with a backup calendar first**
- **The Nuclear Reset is irreversible** - use with extreme caution
- **Auto sync runs every 10 minutes** by default

### Performance Considerations
- **Large organizations** may need to adjust `MAX_EXECUTION_TIME` (default: 4.5 minutes)
- **Google Apps Script limit** is 6 minutes per execution
- **Batch processing** is used to handle large datasets efficiently

### Permissions Required
- **Spreadsheet access**: Read/write access to all club sheets
- **Calendar access**: Full control over the master calendar
- **Script execution**: Ability to create time-based triggers

## 🐛 Troubleshooting

### Common Issues

#### "Calendar not found" Error
- Verify your calendar ID is correct
- Ensure the calendar is shared with your account
- Check that the calendar hasn't been deleted

#### "Sheet not found" Error
- Verify all sheet URLs in ClubRegistry are accessible
- Check that club sheets have the correct "Events" tab
- Ensure proper sharing permissions on club sheets

#### Events Not Appearing
- Run "Debug Sync State" to check for errors
- Verify column headers match exactly (case-sensitive)
- Check that date/time formats are correct

#### Duplicate Events
- Run "Cleanup Duplicates" from the menu
- Check for inconsistent data in club sheets
- Verify key generation is working correctly

#### Performance Issues
- Reduce `MAX_EXECUTION_TIME` if hitting timeouts
- Consider splitting large clubs into multiple sheets
- Use "Quick Sync" instead of "Full Sync" for regular operation

### Getting Help

1. **Check Execution Logs**
   - Go to Apps Script dashboard
   - View "Executions" tab for recent runs
   - Look for error messages and warnings

2. **Run Debug Sync State**
   - Provides comprehensive diagnostic information
   - Shows sync state for all clubs
   - Identifies potential duplicates

3. **Verify Data Integrity**
   - Check all column headers match exactly
   - Ensure date/time formats are consistent
   - Verify all URLs in ClubRegistry are accessible

## 📊 Advanced Configuration

### Customizing Column Layouts

If you need different column layouts, modify the `COLS` configuration:

```javascript
const COLS = {
  REGISTRY: { ClubID: 0, ClubName: 1, ClubSheetURL: 2 },
  EVENTS: { 
    EventName: 0, 
    Date: 1, 
    StartTime: 2, 
    EndTime: 3, 
    Location: 4, 
    Description: 5, 
    DeleteFlag: 6 
  },
  // ... other configurations
};
```

### Adjusting Sync Frequency

To change auto-sync frequency, modify the `setupAutoSync` function:

```javascript
// Change everyMinutes(10) to your preferred interval
ScriptApp.newTrigger('autoSync')
  .timeBased()
  .everyMinutes(15) // Change this value
  .create();
```

### Custom Tab Names

If you prefer different tab names, update the `TABS` configuration:

```javascript
const TABS = {
  REGISTRY: "MyClubRegistry",     // Your custom tab name
  APPROVED: "MyApprovedEvents",   // Your custom tab name
  SYNC_STATE: "MySyncTracking"    // Your custom tab name
};
```

## 🔄 Backup and Recovery

### Regular Backups
- **Google Sheets**: Automatically versioned by Google
- **Google Calendar**: Can export calendar data regularly
- **Configuration**: Keep a copy of your Apps Script code

### Disaster Recovery
1. **Recreate Calendar**: Create new calendar and update calendar ID
2. **Restore Sheets**: Use Google Sheets version history
3. **Re-run Sync**: Use "Full Sync All Clubs" to repopulate

## 📈 Scaling Considerations

### For Large Organizations (50+ clubs)
- Consider splitting into multiple master spreadsheets
- Adjust `MAX_EXECUTION_TIME` to optimize performance
- Monitor Google Apps Script quota usage

### For High-Volume Events (1000+ events)
- Use "Quick Sync" for regular operation
- Schedule "Full Sync" during off-peak hours
- Consider implementing manual approval workflows

## 📞 Support

This system is designed to be self-maintaining, but if you encounter issues:

1. **Check this README** for troubleshooting steps
2. **Review the execution logs** in Apps Script dashboard
3. **Run Debug Sync State** for diagnostic information
4. **Test with a small dataset** before scaling up

---

## 🎉 You're Ready!

Once you've completed these steps, your Club Master Calendar Sync system will be:
- ✅ Automatically syncing events from all club sheets
- ✅ Detecting and preventing duplicates
- ✅ Providing easy maintenance tools
- ✅ Running on a reliable schedule

Enjoy your centralized, automated event management system! 🚀
