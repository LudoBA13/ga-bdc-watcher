const CONFIG = {
	URL_TO_CHECK: 'https://ba13.banquealimentaire.org/bon-de-commande-1290',
	REGEX_WEB_XLSX: /https:\/\/ba13\.banquealimentaire\.org\/sites\/default\/files\/.*?\/([^"\/]+?\.xlsx)/,
	REGEX_MENU_PATTERN: /BA13_(\d{3})_.*?(_\d)?\.xlsx$/,
	FORBIDDEN_SHEET_CHARS: /[\\\/\?\*\[\]\:]/g,
	OPERATING_HOURS: { START: 7, END: 20 },
	TIMEZONE: 'Europe/Paris',
	NOTIFICATION_EMAIL: Session.getEffectiveUser().getEmail()
};

/**
 * Main entry point for the hourly trigger.
 */
function scheduledCheck()
{
	const now = new Date;
	const hour = parseInt(Utilities.formatDate(now, CONFIG.TIMEZONE, 'H'));

	if (hour < CONFIG.OPERATING_HOURS.START || hour > CONFIG.OPERATING_HOURS.END)
	{
		console.log('Outside of operating hours. Current hour: ' + hour);
		return;
	}

	console.log('Running scheduled check at hour: ' + hour);
	checkAndDownload();
}

/**
 * Checks the webpage for a new Excel file and downloads it if detected.
 */
function checkAndDownload()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const content = UrlFetchApp.fetch(CONFIG.URL_TO_CHECK).getContentText();
	const match = content.match(CONFIG.REGEX_WEB_XLSX);

	if (!match)
	{
		console.error('XLSX URL not found on page');
		return;
	}

	const fullUrl = match[0];
	const fileName = decodeURIComponent(match[1]);

	if (isAlreadyProcessed(fileName, ss))
	{
		console.log('No change detected or file already processed: ' + fileName);
		return;
	}

	console.log('New file detected: ' + fileName);
	const blob = UrlFetchApp.fetch(fullUrl).getBlob();
	importExcelContent(blob, fileName, ss);
	sendNotificationEmail(fileName, ss);
}

/**
 * Sends a notification email when a new file is detected.
 */
function sendNotificationEmail(fileName, ss)
{
	const recipient = CONFIG.NOTIFICATION_EMAIL;
	const subject = 'New BDC Detected: ' + fileName;
	const body = 'A new BDC has been detected and imported into the spreadsheet.\n\n' +
		'Filename: ' + fileName + '\n' +
		'Spreadsheet: ' + ss.getUrl() + '\n\n' +
		'This is an automated message.';

	MailApp.sendEmail(recipient, subject, body);
	console.log('Notification email sent to ' + recipient);
}

/**
 * Checks if a filename matches the last entry in the 'Files' sheet.
 */
function isAlreadyProcessed(fileName, ss)
{
	const filesSheet = getOrCreateSheet(ss, 'Files', ['Date', 'Filename', 'Sheet Name']);
	const lastRow = filesSheet.getLastRow();
	if (lastRow <= 1)
	{
		return false;
	}
	return filesSheet.getRange(lastRow, 2).getValue() === fileName;
}

/**
 * Converts XLSX blob to Google Sheets and imports it.
 */
function importExcelContent(blob, fileName, ss)
{
	const fileResource = {
		title: fileName,
		mimeType: MimeType.GOOGLE_SHEETS
	};

	const tempFile = Drive.Files.insert(fileResource, blob, { convert: true });
	try
	{
		const tempSs = SpreadsheetApp.openById(tempFile.id);
		const data = tempSs.getSheets()[0].getDataRange().getValues();
		importDataToNewSheet(data, fileName, ss);
	}
	finally
	{
		Drive.Files.remove(tempFile.id);
	}
}

/**
 * Logic to import raw data into a newly created (and sanitized) sheet.
 */
function importDataToNewSheet(data, fileName, ss)
{
	if (!data || data.length === 0)
	{
		return;
	}

	const sanitizedName = sanitizeSheetName(fileName);
	deleteSheetIfExists(ss, sanitizedName);
	
	const newSheet = ss.insertSheet(sanitizedName);
	newSheet.addDeveloperMetadata('originalFileName', fileName);

	logImport(fileName, sanitizedName, ss);

	newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
	trimSheet(newSheet);
	extractArticles(newSheet);
}

/**
 * Manual import from UI prompt.
 */
function importFromSpreadsheetUrl()
{
	const ui = SpreadsheetApp.getUi();
	const response = ui.prompt('Import from URL', 'Enter the URL of the Google Sheet or Excel file (.xlsx) to import:', ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() !== ui.Button.OK)
	{
		return;
	}

	const url = response.getResponseText().trim();
	const targetSs = SpreadsheetApp.getActiveSpreadsheet();

	if (url.toLowerCase().includes('.xlsx'))
	{
		handleExcelUrlImport(url, targetSs, ui);
		return;
	}

	handleGoogleSheetUrlImport(url, targetSs, ui);
}

function handleExcelUrlImport(url, ss, ui)
{
	try
	{
		const fileName = decodeURIComponent(url.split('/').pop().split('?')[0]);
		const blob = UrlFetchApp.fetch(url).getBlob();
		importExcelContent(blob, fileName, ss);
	}
	catch (e)
	{
		ui.alert('Error importing Excel file: ' + e.message);
	}
}

function handleGoogleSheetUrlImport(url, ss, ui)
{
	const id = extractSpreadsheetId(url);
	if (!id)
	{
		ui.alert('Invalid URL. Could not find a Spreadsheet ID or Excel file.');
		return;
	}

	try
	{
		const sourceSs = SpreadsheetApp.openById(id);
		const data = sourceSs.getSheets()[0].getDataRange().getValues();
		importDataToNewSheet(data, sourceSs.getName(), ss);
	}
	catch (e)
	{
		ui.alert('Error opening Google Sheet: ' + e.message);
	}
}

/**
 * Extracts data from a sheet into the 'Articles' summary sheet.
 */
function extractArticles(sheet)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const articlesSheet = getOrCreateSheet(ss, 'Articles', ['Sheet Name', 'Category', 'Article ID', 'Label', 'Unit', 'Quantity']);
	const data = sheet.getDataRange().getValues();
	const displayName = getDisplayName(sheet);

	let extractionState = {
		isRecording: false,
		unit: '',
		category: '',
		headerMap: {},
		results: []
	};

	for (let i = 0; i < data.length; i++)
	{
		const row = data[i];
		const sectionHeader = detectSectionHeader(row[0]);

		if (sectionHeader)
		{
			i = setupSection(extractionState, sectionHeader, data, i, displayName);
			continue;
		}

		if (extractionState.isRecording)
		{
			processArticleRow(row, extractionState, displayName);
		}
	}

	if (extractionState.results.length > 0)
	{
		articlesSheet.getRange(articlesSheet.getLastRow() + 1, 1, extractionState.results.length, 6).setValues(extractionState.results);
		ss.toast('Extracted ' + extractionState.results.length + ' articles.');
	}
}

function detectSectionHeader(cellValue)
{
	const headerRegex = /^Produit.*?(homolog|picerie|tout).*?(COLIS|KILO)/i;
	const match = String(cellValue).match(headerRegex);
	if (!match)
	{
		return null;
	}

	return {
		unit: match[2].toUpperCase() === 'KILO' ? 'kg' : 'colis',
		typeKey: match[1].toLowerCase()
	};
}

function setupSection(state, header, data, currentIndex, sheetName)
{
	state.isRecording = true;
	state.unit = header.unit;
	state.category = mapCategory(header.typeKey);
	
	const columnRow = data[currentIndex + 1];
	if (!columnRow)
	{
		return currentIndex;
	}

	state.headerMap = mapColumns(columnRow);
	validateHeaders(state.headerMap, state.unit, sheetName);
	
	return currentIndex + 1; // Skip the header row
}

function processArticleRow(row, state, sheetName)
{
	const idVal = row[state.headerMap.ARTICLE];
	if (!idVal || isNaN(idVal) || String(idVal).trim() === '')
	{
		state.isRecording = false;
		return;
	}

	const labelVal = row[state.headerMap.DESIGNATION];
	const quantity = calculateQuantity(row, state);

	state.results.push([sheetName, state.category, idVal, labelVal, state.unit, quantity]);
}

function calculateQuantity(row, state)
{
	if (state.unit === 'kg')
	{
		return row[state.headerMap['Max en KG']];
	}

	const poidsBrut = parseFloat(String(row[state.headerMap['Poids brut']]).replace(',', '.')) || 0;
	const nbMax = parseFloat(row[state.headerMap['Nb max de colis']]) || 0;
	return poidsBrut * nbMax;
}

/**
 * Utility functions.
 */

function mapCategory(typeKey)
{
	if (typeKey.includes('homolog'))
	{
		return 'Asso';
	}
	if (typeKey.includes('picerie'))
	{
		return 'ES';
	}
	return 'Asso|ES';
}

function mapColumns(row)
{
	const targets = ['ARTICLE', 'DESIGNATION', 'Max en KG', 'Poids brut', 'Nb max de colis'];
	const map = {};
	targets.forEach(t =>
	{
		map[t] = row.findIndex(cell => String(cell).trim().startsWith(t));
	});
	return map;
}

function validateHeaders(map, unit, sheetName)
{
	if (map.ARTICLE === -1)
	{
		throw new Error('Missing "ARTICLE" in ' + sheetName);
	}
	if (map.DESIGNATION === -1)
	{
		throw new Error('Missing "DESIGNATION" in ' + sheetName);
	}
	if (unit === 'kg' && map['Max en KG'] === -1)
	{
		throw new Error('Missing "Max en KG" in ' + sheetName);
	}
	if (unit === 'colis' && (map['Poids brut'] === -1 || map['Nb max de colis'] === -1))
	{
		throw new Error('Missing shipping columns in ' + sheetName);
	}
}

function extractSpreadsheetId(url)
{
	const dMatch = url.match(/\/d\/([-\w]+)/);
	return dMatch ? dMatch[1] : (url.match(/[-\w]{25,}/) || [])[0];
}

function sanitizeSheetName(name)
{
	const menuMatch = name.match(CONFIG.REGEX_MENU_PATTERN);
	if (menuMatch)
	{
		return 'Menu' + menuMatch[1] + (menuMatch[2] || '');
	}

	let sanitized = name.replace(CONFIG.FORBIDDEN_SHEET_CHARS, '').replace(/\.xlsx$/i, '');
	return sanitized.length > 31 ? sanitized.slice(-31) : sanitized;
}

function getDisplayName(sheet)
{
	const originalName = getOriginalFileName(sheet);
	const menuMatch = originalName.match(CONFIG.REGEX_MENU_PATTERN);
	return menuMatch ? 'Menu' + menuMatch[1] + (menuMatch[2] || '') : originalName;
}

function getOriginalFileName(sheet)
{
	const meta = sheet.getDeveloperMetadata().find(m => m.getKey() === 'originalFileName');
	return meta ? meta.getValue() : sheet.getName();
}

function getOrCreateSheet(ss, name, headers)
{
	let sheet = ss.getSheetByName(name);
	if (!sheet)
	{
		sheet = ss.insertSheet(name);
		sheet.appendRow(headers);
	}
	return sheet;
}

function deleteSheetIfExists(ss, name)
{
	const sheet = ss.getSheetByName(name);
	if (sheet)
	{
		ss.deleteSheet(sheet);
	}
}

function logImport(fileName, sheetName, ss)
{
	const sheet = getOrCreateSheet(ss, 'Files', ['Date', 'Filename', 'Sheet Name']);
	sheet.appendRow([new Date, fileName, sheetName]);
}

function trimSheet(sheet)
{
	const lastRow = sheet.getLastRow();
	const maxRows = sheet.getMaxRows();
	if (maxRows > lastRow)
	{
		sheet.deleteRows(lastRow + 1, maxRows - lastRow);
	}

	const lastCol = sheet.getLastColumn();
	const maxCols = sheet.getMaxColumns();
	if (maxCols > lastCol)
	{
		sheet.deleteColumns(lastCol + 1, maxCols - lastCol);
	}
}

function onOpen()
{
	SpreadsheetApp.getUi().createMenu('BA Tools')
		.addItem('Extract Articles from Current Sheet', 'extractArticlesFromActiveSheet')
		.addItem('Import from Spreadsheet URL', 'importFromSpreadsheetUrl')
		.addSeparator()
		.addItem('Test Email Notification', 'testEmailNotification')
		.addToUi();
}

function testEmailNotification()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	try
	{
		sendNotificationEmail('TEST_FILE.xlsx', ss);
		SpreadsheetApp.getUi().alert('Test email sent to ' + CONFIG.NOTIFICATION_EMAIL);
	}
	catch (e)
	{
		SpreadsheetApp.getUi().alert('Error sending email: ' + e.message);
	}
}

function extractArticlesFromActiveSheet()
{
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	if (['Files', 'Articles'].includes(sheet.getName()))
	{
		SpreadsheetApp.getUi().alert('Cannot extract articles from this sheet.');
		return;
	}
	extractArticles(sheet);
}

function setupTrigger()
{
	const triggers = ScriptApp.getProjectTriggers();
	if (triggers.some(t => t.getHandlerFunction() === 'scheduledCheck'))
	{
		console.log('Trigger already exists.');
		return;
	}
	ScriptApp.newTrigger('scheduledCheck').timeBased().everyHours(1).create();
	console.log('Trigger created.');
}
