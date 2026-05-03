const URL_TO_CHECK = 'https://ba13.banquealimentaire.org/bon-de-commande-1290';
const REGEX_PATTERN = /https:\/\/ba13\.banquealimentaire\.org\/sites\/default\/files\/.*?\/([^"\/]+?)\.xlsx/;

function checkAndDownload()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let filesSheet = ss.getSheetByName('Files');
	if (!filesSheet)
	{
		filesSheet = ss.insertSheet('Files');
		filesSheet.appendRow(['Date', 'Filename']);
	}

	const content = UrlFetchApp.fetch(URL_TO_CHECK).getContentText();
	const match = content.match(REGEX_PATTERN);

	if (!match)
	{
		console.error('URL not found on page');
		return;
	}

	const fullUrl = match[0];
	const fileName = decodeURIComponent(match[1]);

	const lastRow = filesSheet.getLastRow();
	let lastFileName = '';
	if (lastRow > 1)
	{
		lastFileName = filesSheet.getRange(lastRow, 2).getValue();
	}

	if (fileName !== lastFileName)
	{
		console.log('New file detected: ' + fileName);
		processExcelFile(fullUrl, fileName, ss, filesSheet);
	}
	else
	{
		console.log('No change detected.');
	}
}

function processExcelFile(url, fileName, ss, filesSheet)
{
	const response = UrlFetchApp.fetch(url);
	const blob = response.getBlob();

	// Store in Files sheet
	filesSheet.appendRow([new Date(), fileName]);

	// Import content
	importExcelContent(blob, fileName, ss);
}

function importExcelContent(blob, fileName, ss)
{
	// To convert XLSX to Google Sheets, we need Advanced Drive Service
	// We'll create a temporary file in Drive and then read it
	const fileResource = {
		title: fileName,
		mimeType: MimeType.GOOGLE_SHEETS
	};

	const tempFile = Drive.Files.insert(fileResource, blob, { convert: true });
	const tempSs = SpreadsheetApp.openById(tempFile.id);
	const sourceSheet = tempSs.getSheets()[0];
	const data = sourceSheet.getDataRange().getValues();

	// Create new sheet in target spreadsheet
	const sanitizedName = sanitizeSheetName(fileName);
	let newSheet = ss.getSheetByName(sanitizedName);
	if (newSheet)
	{
		// If sheet exists, delete it as we want the new content
		ss.deleteSheet(newSheet);
	}
	newSheet = ss.insertSheet(sanitizedName);

	if (data.length > 0)
	{
		newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
		trimSheet(newSheet);
	}

	// Clean up temp file
	Drive.Files.remove(tempFile.id);
}

function trimSheet(sheet)
{
	const lastRow = sheet.getLastRow();
	const maxRows = sheet.getMaxRows();
	if (maxRows > lastRow)
	{
		sheet.deleteRows(lastRow + 1, maxRows - lastRow);
	}

	const lastColumn = sheet.getLastColumn();
	const maxColumns = sheet.getMaxColumns();
	if (maxColumns > lastColumn)
	{
		sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
	}
}

function scheduledCheck()
{
	const now = new Date();
	// Get hour in CET (Europe/Paris)
	const hour = parseInt(Utilities.formatDate(now, 'Europe/Paris', 'H'));

	if (hour >= 7 && hour <= 20)
	{
		console.log('Running scheduled check at hour: ' + hour);
		checkAndDownload();
	}
	else
	{
		console.log('Outside of operating hours (7am-8pm CET). Current hour: ' + hour);
	}
}

/**
 * Run this once manually to set up the hourly trigger.
 */
function setupTrigger()
{
	const triggers = ScriptApp.getProjectTriggers();
	const triggerName = 'scheduledCheck';

	// Check if trigger already exists
	const exists = triggers.some(t => t.getHandlerFunction() === triggerName);

	if (!exists)
	{
		ScriptApp.newTrigger(triggerName)
			.timeBased()
			.everyHours(1)
			.create();
		console.log('Trigger created for ' + triggerName);
	}
	else
	{
		console.log('Trigger for ' + triggerName + ' already exists.');
	}
}

function sanitizeSheetName(name)
{
	// Remove forbidden characters: \ / ? * [ ] :
	let sanitized = name.replace(/[\\\/\?\*\[\]\:]/g, '');
	// Truncate to 31 characters
	return sanitized.substring(0, 31);
}
