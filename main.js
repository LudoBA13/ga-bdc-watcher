const CONFIG = {
	URL_TO_CHECK: 'https://ba13.banquealimentaire.org/bon-de-commande-1290',
	REGEX_WEB_XLSX: /https:\/\/ba13\.banquealimentaire\.org\/sites\/default\/files\/.*?\/([^"\/]+?\.xlsx)/,
	REGEX_MENU_PATTERN: /BA13_(\d{3})_.*?(_\d)?\.xlsx$/,
	FORBIDDEN_SHEET_CHARS: /[\\\/\?\*\[\]\:]/g,
	OPERATING_HOURS: { START: 7, END: 20 },
	TIMEZONE: 'Europe/Paris',
	NOTIFICATION_EMAIL: Session.getEffectiveUser().getEmail(),
	DRIVE_FOLDER_ID: '1wlCuBGmKa8yWJePfA5MmE3Aqpnhggjq8'
};

/**
 * Main entry point for the hourly trigger.
 */
function scheduledCheck()
{
	updateCachedMaxDate();

	const now = new Date;
	const hour = parseInt(Utilities.formatDate(now, CONFIG.TIMEZONE, 'H'));

	if (hour < CONFIG.OPERATING_HOURS.START || hour > CONFIG.OPERATING_HOURS.END)
	{
		return;
	}

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
		return;
	}

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
}

/**
 * Checks if a filename matches the last entry in the 'Files' sheet.
 */
function isAlreadyProcessed(fileName, ss)
{
	const filesSheet = getOrCreateSheet(ss, 'Files', ['Date', 'Filename', 'sheetName', 'File ID']);
	const lastRow = filesSheet.getLastRow();
	if (lastRow <= 1)
	{
		return false;
	}
	return filesSheet.getRange(lastRow, 2).getValue() === fileName;
}

/**
 * Saves the original XLSX blob as a file, and imports a converted version into the spreadsheet.
 */
function importExcelContent(blob, fileName, ss)
{
	// 1. Save original binary file
	const rawFileResource = {
		title: fileName,
		mimeType: MimeType.MICROSOFT_EXCEL,
		parents: [{ id: CONFIG.DRIVE_FOLDER_ID }]
	};
	const rawFile = Drive.Files.insert(rawFileResource, blob, { convert: false });

	// 2. Save converted version to Drive (for processing)
	const sheetFileResource = {
		title: fileName,
		mimeType: MimeType.GOOGLE_SHEETS,
		parents: [{ id: CONFIG.DRIVE_FOLDER_ID }]
	};
	const sheetFile = Drive.Files.insert(sheetFileResource, blob, { convert: true });
	try
	{
		// 3. Import data from the converted file
		const tempSs = SpreadsheetApp.openById(sheetFile.id);
		const data = tempSs.getSheets()[0].getDataRange().getValues();
		importDataToNewSheet(data, fileName, ss);
		const sanitizedName = sanitizeSheetName(fileName);
		logImport(fileName, sanitizedName, ss, rawFile.id);
	}
	finally
	{
		Drive.Files.remove(sheetFile.id);
	}
}

/**
 * Updates the 'CurrentMenu' sheet based on the sheet name of the most recent imported file in 'Files',
 * mapping key-values from the 'Menus' sheet.
 */
function recomputeCurrentMenu()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const filesSheet = ss.getSheetByName('Files');
	if (!filesSheet)
	{
		throw new Error('Sheet "Files" not found.');
	}

	const filesData = filesSheet.getDataRange().getValues();
	if (filesData.length <= 1)
	{
		throw new Error('Sheet "Files" contains no import data.');
	}

	const filesHeaders = filesData[0];
	const dateIndex = filesHeaders.indexOf('Date');
	const sheetNameIndex = filesHeaders.indexOf('sheetName');

	if (dateIndex === -1)
	{
		throw new Error('Column "Date" not found in "Files" sheet.');
	}
	if (sheetNameIndex === -1)
	{
		throw new Error('Column "sheetName" not found in "Files" sheet.');
	}

	let latestTime = -1;
	let targetSheetName = '';

	for (let i = 1; i < filesData.length; i++)
	{
		const dateVal = filesData[i][dateIndex];
		const date = dateVal instanceof Date ? dateVal : new Date(dateVal);
		if (date && !isNaN(date.getTime()) && date.getTime() > latestTime)
		{
			latestTime = date.getTime();
			targetSheetName = String(filesData[i][sheetNameIndex]);
		}
	}

	if (!targetSheetName)
	{
		throw new Error('No target sheetName could be found from the most recent imported file.');
	}

	const menusSheet = ss.getSheetByName('Menus');
	if (!menusSheet)
	{
		throw new Error('Sheet "Menus" not found.');
	}

	const menusData = menusSheet.getDataRange().getValues();
	if (menusData.length <= 1)
	{
		throw new Error('Sheet "Menus" contains no data.');
	}

	const menusHeaders = menusData[0];
	const menusSheetNameIndex = menusHeaders.indexOf('sheetName');
	if (menusSheetNameIndex === -1)
	{
		throw new Error('Column "sheetName" not found in "Menus" sheet.');
	}

	let targetRow = null;
	for (let i = 1; i < menusData.length; i++)
	{
		if (String(menusData[i][menusSheetNameIndex]) === targetSheetName)
		{
			targetRow = menusData[i];
			break;
		}
	}

	if (!targetRow)
	{
		throw new Error('Could not find a row in "Menus" matching sheetName: ' + targetSheetName);
	}

	const currentMenuSheet = getOrCreateSheet(ss, 'CurrentMenu', ['Property', 'Value']);
	currentMenuSheet.clearContents();
	currentMenuSheet.appendRow(['Property', 'Value']);

	for (let j = 0; j < menusHeaders.length; j++)
	{
		const key = menusHeaders[j];
		const val = targetRow[j];
		currentMenuSheet.appendRow([key, val]);
	}
	trimSheet(currentMenuSheet);
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

	newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
	trimSheet(newSheet);
	extractArticles(newSheet);
	logMenuData(newSheet, ss);
	recomputeCurrentMenu();
}

/**
 * Extracts specific menu data from the designated cells of a given sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to extract data from.
 * @returns {{name: string, pickupDateStart: Date|string, pickupDateEnd: Date|string}|null} The extracted data map or null if sheet is invalid.
 */
function extractMenuData(sheet)
{
	if (!sheet)
	{
		return null;
	}

	// Assume H2, H3, H4 based on requirements
	const range = sheet.getRange('H2:H4');
	const values = range.getValues();

	return {
		name: values[0][0],
		pickupDateStart: values[1][0],
		pickupDateEnd: values[2][0]
	};
}

/**
 * Logs menu data into the 'Menus' sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to extract data from.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [targetSheet] Optional target sheet to log to.
 */
function logMenuData(sheet, ss, targetSheet = null)
{
	const data = extractMenuData(sheet);
	if (!data)
	{
		return;
	}

	const period = getPlanningPeriod(data.pickupDateStart, data.pickupDateEnd);
	const planningDates = computePlanningDates(data.pickupDateStart, data.pickupDateEnd);

	const headers = ['menuId', 'sheetName', 'menuName', 'pickupDateStart', 'pickupDateEnd', 'articles', 'planningYear', 'planningMonth', 'planningDates'];
	const menusSheet = targetSheet || getOrCreateSheet(ss, 'Menus', headers);
	menusSheet.appendRow([
		sheet.getName(),
		sheet.getName(),
		data.name,
		data.pickupDateStart,
		data.pickupDateEnd,
		JSON.stringify(getMenuArticles(sheet)),
		period.year,
		period.month,
		JSON.stringify(planningDates)
	]);
}

/**
 * Computes the planning period (year and month) from a date range.
 *
 * @param {Date|string} start The start of the range.
 * @param {Date|string} end The end of the range.
 * @returns {{year: number|string, month: number|string}} The planning year and month.
 */
function getPlanningPeriod(start, end)
{
	const startDate = start instanceof Date ? start : new Date(start);
	const endDate = end instanceof Date ? end : new Date(end);

	if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()))
	{
		return { year: '', month: '' };
	}

	let current = new Date(startDate.getTime());
	while (current <= endDate)
	{
		const tick = dateToTick(current);
		const parsed = parseTick(tick);
		if (parsed && parsed.year && parsed.month)
		{
			return { year: parsed.year, month: parsed.month };
		}
		current.setDate(current.getDate() + 1);
	}

	return { year: '', month: '' };
}

/**
 * Computes a map of planning codes to dates for a given range.
 *
 * @param {Date|string} start The start of the range.
 * @param {Date|string} end The end of the range.
 * @returns {Object<string, Date>} A map of planning codes to dates.
 */
function computePlanningDates(start, end)
{
	const startDate = start instanceof Date ? start : new Date(start);
	const endDate = end instanceof Date ? end : new Date(end);
	const result = {};

	if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()))
	{
		return result;
	}

	let current = new Date(startDate.getTime());
	while (current <= endDate)
	{
		const tick = dateToTick(current);
		const parsed = parseTick(tick);
		if (parsed && parsed.timeslot <= 3 && parsed.code)
		{
			const yyyy = current.getFullYear();
			const mm = String(current.getMonth() + 1).padStart(2, '0');
			const dd = String(current.getDate()).padStart(2, '0');
			const dateStr = yyyy + '-' + mm + '-' + dd;
			result[parsed.code] = dateStr;
		}
		current.setDate(current.getDate() + 1);
	}

	return result;
}

/**
 * Recomputes all menu data in the 'Menus' sheet.
 */
function recomputeAllMenuData()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const headers = ['menuId', 'sheetName', 'menuName', 'pickupDateStart', 'pickupDateEnd', 'articles', 'planningYear', 'planningMonth', 'planningDates'];
	const menusSheet = getOrCreateSheet(ss, 'Menus', headers);

	menusSheet.clearContents();
	menusSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

	const sheets = ss.getSheets().filter(s => /^Menu\d/.test(s.getName()));
	sheets.forEach(sheet =>
	{
		logMenuData(sheet, ss, menusSheet);
	});

	trimSheet(menusSheet);
	recomputeCurrentMenu();
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
		const fileName = sourceSs.getName();
		importDataToNewSheet(data, fileName, ss);
		const sanitizedName = sanitizeSheetName(fileName);
		logImport(fileName, sanitizedName, ss, id);
	}
	catch (e)
	{
		ui.alert('Error opening Google Sheet: ' + e.message);
	}
}

/**
 * Extracts menu articles into a structured array.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to extract articles from.
 * @returns {Array<Object>} The list of articles.
 */
function getMenuArticles(sheet)
{
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
			const idVal = row[extractionState.headerMap.ARTICLE];
			if (!idVal || isNaN(idVal) || String(idVal).trim() === '')
			{
				extractionState.isRecording = false;
				continue;
			}

			const labelVal = row[extractionState.headerMap.DESIGNATION];
			const quantity = row[extractionState.headerMap.MAX_QTY];
			const unitWeight = calculateUnitWeight(row, extractionState);

			extractionState.results.push({
				category: extractionState.category,
				id: idVal,
				label: labelVal,
				unit: extractionState.unit,
				unitWeight: unitWeight,
				maxQty: quantity
			});
		}
	}
	return extractionState.results;
}

/**
 * Extracts data from a sheet into the 'Articles' summary sheet.
 */
function extractArticles(sheet)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const articlesSheet = getOrCreateSheet(ss, 'Articles', ['Sheet Name', 'Category', 'Article ID', 'Label', 'Unit', 'Unit Weight', 'Max Qty']);
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
		articlesSheet.getRange(articlesSheet.getLastRow() + 1, 1, extractionState.results.length, 7).setValues(extractionState.results);
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
	const quantity = row[state.headerMap.MAX_QTY];
	const unitWeight = calculateUnitWeight(row, state);

	state.results.push([sheetName, state.category, idVal, labelVal, state.unit, unitWeight, quantity]);
}

function calculateUnitWeight(row, state)
{
	if (state.unit === 'kg')
	{
		return 1;
	}

	const val = row[state.headerMap['Poids brut du colis']];
	return val ? parseFloat(String(val).replace(',', '.')) : 1;
}

function mapColumns(row)
{
	const map = {
		ARTICLE: row.findIndex(cell => String(cell).trim().startsWith('ARTICLE')),
		DESIGNATION: row.findIndex(cell => String(cell).trim().startsWith('DESIGNATION')),
		'Poids brut du colis': row.findIndex(cell => String(cell).trim().startsWith('Poids brut du colis')),
		MAX_QTY: row.findIndex(cell => /max.*100.*UD/i.test(String(cell).trim()))
	};
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
	if (map.MAX_QTY === -1)
	{
		throw new Error('Missing Max Qty column (expected pattern /max.*100.*UD/i) in ' + sheetName + '. Found: ' + Object.keys(map).join(', '));
	}
}

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
		styleHeader(sheet.getRange(1, 1, 1, headers.length));
	}
	return sheet;
}

function styleHeader(range)
{
	range.setHorizontalAlignment('center')
		.setVerticalAlignment('middle')
		.setBackground('#4a86e8')
		.setFontWeight('bold')
		.setFontColor('white');
}

function trimSheet(sheet)
{
	const maxRows = sheet.getMaxRows();
	const lastRow = sheet.getLastRow();
	const maxCols = sheet.getMaxColumns();
	const lastCol = sheet.getLastColumn();

	// Remove empty rows at the bottom
	if (maxRows > lastRow && lastRow > 0)
	{
		const rowsToDelete = maxRows - lastRow;
		sheet.deleteRows(lastRow + 1, rowsToDelete);
	}

	// Remove empty columns at the right
	if (maxCols > lastCol && lastCol > 0)
	{
		const colsToDelete = maxCols - lastCol;
		sheet.deleteColumns(lastCol + 1, colsToDelete);
	}
}

function deleteSheetIfExists(ss, name)
{
	const sheet = ss.getSheetByName(name);
	if (sheet)
	{
		ss.deleteSheet(sheet);
	}
}

function logImport(fileName, sheetName, ss, fileId)
{
	const sheet = getOrCreateSheet(ss, 'Files', ['Date', 'Filename', 'sheetName', 'File ID']);
	sheet.appendRow([new Date, fileName, sheetName, fileId]);
}

function onOpen()
{
	SpreadsheetApp.getUi().createMenu('BA Tools')
		.addItem('Extract all menu data', 'recomputeAllMenuData')
		.addSeparator()
		.addItem('Setup Trigger', 'setupTrigger')
		.addToUi();
}

function setupTrigger()
{
	const triggers = ScriptApp.getProjectTriggers();
	if (triggers.some(t => t.getHandlerFunction() === 'scheduledCheck'))
	{
		return;
	}
	ScriptApp.newTrigger('scheduledCheck').timeBased().everyHours(1).create();
}
