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

	// Import content
	importDataToNewSheet(data, fileName, ss);

	// Clean up temp file
	Drive.Files.remove(tempFile.id);
}

function importDataToNewSheet(data, fileName, ss)
{
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
		extractArticles(newSheet);
	}
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

function onOpen()
{
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('BA Tools')
		.addItem('Extract Articles from Current Sheet', 'extractArticlesFromActiveSheet')
		.addItem('Import from Spreadsheet URL', 'importFromSpreadsheetUrl')
		.addToUi();
}

function importFromSpreadsheetUrl()
{
	const ui = SpreadsheetApp.getUi();
	const response = ui.prompt('Import from URL', 'Enter the URL of the Google Sheet or Excel file (.xlsx) to import:', ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() === ui.Button.OK)
	{
		const url = response.getResponseText().trim();
		const targetSs = SpreadsheetApp.getActiveSpreadsheet();

		// Check if it's an XLSX file
		if (url.toLowerCase().includes('.xlsx'))
		{
			try
			{
				// Extract filename from URL
				const parts = url.split('/');
				let fileName = parts[parts.length - 1].split('?')[0];
				fileName = decodeURIComponent(fileName).replace(/\.xlsx$/i, '');
				
				const res = UrlFetchApp.fetch(url);
				const blob = res.getBlob();
				importExcelContent(blob, fileName, targetSs);
				ui.alert('Imported successfully from Excel: ' + fileName);
				return;
			}
			catch (e)
			{
				ui.alert('Error importing Excel file: ' + e.message);
				return;
			}
		}

		// Try to extract Google Sheet ID
		// First try the standard /d/ID/ pattern
		let id = null;
		const dMatch = url.match(/\/d\/([-\w]+)/);
		if (dMatch)
		{
			id = dMatch[1];
		}
		else
		{
			// Fallback to searching for any string that looks like an ID
			const gSheetMatch = url.match(/[-\w]{25,}/);
			if (gSheetMatch)
			{
				id = gSheetMatch[0];
			}
		}

		if (!id)
		{
			ui.alert('Invalid URL. Could not find a Spreadsheet ID or Excel file.');
			return;
		}

		try
		{
			const sourceSs = SpreadsheetApp.openById(id);
			const fileName = sourceSs.getName();
			const sourceSheet = sourceSs.getSheets()[0]; // Get the first sheet
			const data = sourceSheet.getDataRange().getValues();

			importDataToNewSheet(data, fileName, targetSs);
			ui.alert('Imported successfully from Google Sheet: ' + fileName);
		}
		catch (e)
		{
			ui.alert('Error opening Google Sheet: ' + e.message);
		}
	}
}

function extractArticlesFromActiveSheet()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();
	if (sheet.getName() === 'Files' || sheet.getName() === 'Articles')
	{
		SpreadsheetApp.getUi().alert('Cannot extract articles from this sheet.');
		return;
	}
	extractArticles(sheet);
}

function extractArticles(sheet)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let articlesSheet = ss.getSheetByName('Articles');
	if (!articlesSheet)
	{
		articlesSheet = ss.insertSheet('Articles');
		articlesSheet.appendRow(['Sheet Name', 'Category', 'Article ID', 'Label', 'Unit', 'Quantity']);
	}

	const data = sheet.getDataRange().getValues();
	const sheetName = sheet.getName();
	const headerRegex = /^Produit.*?(homolog|picerie|tout).*?(COLIS|KILO)/i;

	let isRecording = false;
	let currentUnit = '';
	let currentCategory = '';
	let headerMap = {};
	let extractedData = [];

	for (let i = 0; i < data.length; i++)
	{
		const row = data[i];
		const firstCell = String(row[0]);

		const headerMatch = firstCell.match(headerRegex);
		if (!isRecording && headerMatch)
		{
			// Header found, look for columns on next line
			isRecording = true;
			currentUnit = headerMatch[2].toUpperCase() === 'KILO' ? 'kg' : 'colis';
			
			const typeKey = headerMatch[1].toLowerCase();
			if (typeKey.includes('homolog'))
			{
				currentCategory = 'Asso';
			}
			else if (typeKey.includes('picerie'))
			{
				currentCategory = 'ES';
			}
			else if (typeKey.includes('tout'))
			{
				currentCategory = 'Asso|ES';
			}

			const nextRow = data[i + 1];
			if (!nextRow)
			{
				continue;
			}

			// Detect headers
			const requiredHeaders = ["ARTICLE", "DESIGNATION", "Max en KG", "Poids brut", "Nb max de colis"];
			headerMap = {};
			
			for (const target of requiredHeaders)
			{
				headerMap[target] = -1;
				for (let j = 0; j < nextRow.length; j++)
				{
					const cellVal = String(nextRow[j]).trim();
					if (cellVal.startsWith(target))
					{
						headerMap[target] = j;
						break;
					}
				}
			}

			// Validate headers
			if (headerMap["ARTICLE"] === -1)
			{
				throw new Error('Missing "ARTICLE" column in sheet ' + sheetName);
			}
			if (headerMap["DESIGNATION"] === -1)
			{
				throw new Error('Missing "DESIGNATION" column in sheet ' + sheetName);
			}
			
			if (currentUnit === 'kg' && headerMap["Max en KG"] === -1)
			{
				throw new Error('Missing "Max en KG" column in sheet ' + sheetName);
			}
			if (currentUnit === 'colis')
			{
				if (headerMap["Poids brut"] === -1)
				{
					throw new Error('Missing "Poids brut" column in sheet ' + sheetName);
				}
				if (headerMap["Nb max de colis"] === -1)
				{
					throw new Error('Missing "Nb max de colis" column in sheet ' + sheetName);
				}
			}

			i++; // Skip the column header row
			continue;
		}

		if (isRecording)
		{
			const idVal = row[headerMap["ARTICLE"]];
			const labelVal = row[headerMap["DESIGNATION"]];

			// Stop recording the current section if we hit a non-numeric value in the article column
			if (idVal === '' || isNaN(idVal) || String(idVal).trim() === '')
			{
				isRecording = false;
				continue;
			}

			let quantity = 0;
			if (currentUnit === 'kg')
			{
				quantity = row[headerMap["Max en KG"]];
			}
			else
			{
				const rawPoidsBrut = row[headerMap["Poids brut"]];
				let poidsBrutStr = String(rawPoidsBrut);
				// Handle French locale (comma as decimal separator)
				const sanitizedPoidsBrut = poidsBrutStr.replace(',', '.');
				const poidsBrut = parseFloat(sanitizedPoidsBrut) || 0;
				
				const nbMax = parseFloat(row[headerMap["Nb max de colis"]]) || 0;
				quantity = poidsBrut * nbMax;
			}

			extractedData.push([sheetName, currentCategory, idVal, labelVal, currentUnit, quantity]);
		}
	}

	if (extractedData.length > 0)
	{
		articlesSheet.getRange(articlesSheet.getLastRow() + 1, 1, extractedData.length, 6).setValues(extractedData);
		SpreadsheetApp.getActiveSpreadsheet().toast('Extracted ' + extractedData.length + ' articles.');
	}
}

function sanitizeSheetName(name)
{
	// Remove forbidden characters: \ / ? * [ ] :
	let sanitized = name.replace(/[\\\/\?\*\[\]\:]/g, '');
	// Truncate to 31 characters, keeping the rightmost part
	return sanitized.length > 31 ? sanitized.slice(-31) : sanitized;
}
