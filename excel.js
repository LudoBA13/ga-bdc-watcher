/**
 * Generates an Excel order from a template and returns it as a blob.
 * 
 * @param {Object} orderData The order details.
 * @param {string} orderData.codeVif The partner code.
 * @param {string} orderData.pickupDate The pickup date (YYYY-MM-DD).
 * @param {Object} orderData.articles Map of Article ID (string) to Quantity (number/string).
 * @return {Object} {blob: Blob, logs: string}
 */
function generateOrderExcelBlob(orderData)
{
	const logs = [];
	logs.push('Order Generation Started: ' + new Date().toISOString());
	logs.push('Order Data: ' + JSON.stringify(orderData));

	const templateId = getLatestTemplateId();
	if (!templateId)
	{
		throw new Error('Could not find a valid template in the Files sheet.');
	}

	// 1. Create a temporary Google Sheet from the Excel template
	const tempSheetId = createTemporarySheetFromExcel(templateId, 'Order_' + orderData.codeVif);

	try
	{
		const ss = SpreadsheetApp.openById(tempSheetId);
		const sheet = ss.getSheets()[0]; // Assume first sheet is the one to fill

		// 2. Fill static fields
		// codeVif -> D27
		sheet.getRange('D27').setValue(orderData.codeVif);
		// pickupDate -> D31
		sheet.getRange('D31').setValue(orderData.pickupDate);

		// 3. Fill articles
		// Range A40:A110 contains Article IDs
		const rangeA = sheet.getRange('A40:A110');
		const ids = rangeA.getValues();
		const quantities = new Array(ids.length).fill(['']);

		// Process only articles present in orderData
		const missingArticles = [];
		for (const [articleId, qty] of Object.entries(orderData.articles))
		{
			if (!qty || parseFloat(qty) === 0)
			{
				continue;
			}

			// Find corresponding row in template
			let found = false;
			for (let i = 0; i < ids.length; i++)
			{
				if (normalizeArticleId(ids[i][0]) === articleId)
				{
					quantities[i] = [qty];
					logs.push('Matched Article ID ' + articleId + ' with Qty: ' + qty);
					found = true;
					break;
				}
			}
			
			if (!found)
			{
				const msg = 'WARNING: Article ID ' + articleId + ' provided in order, but not found in template range (A40:A110).';
				logs.push(msg);
				missingArticles.push(msg);
			}
		}

		if (missingArticles.length > 0)
		{
			logs.push('\nSummary of missing articles:\n' + missingArticles.join('\n'));
		}

		// Write quantities to column H (8th column)
		sheet.getRange(40, 8, quantities.length, 1).setValues(quantities);

		// 4. Force calculation
		SpreadsheetApp.flush();

		// 5. Export back to XLSX
		const url = 'https://docs.google.com/spreadsheets/d/' + tempSheetId + '/export?format=xlsx';
		const token = ScriptApp.getOAuthToken();
		const response = UrlFetchApp.fetch(url, {
			headers: {
				'Authorization': 'Bearer ' + token
			}
		});

		const fileName = 'Commande_' + orderData.codeVif + '_' + orderData.pickupDate + '.xlsx';
		return {
			blob: response.getBlob().setName(fileName),
			logs: logs.join('\n')
		};
	}
	finally
	{
		// 6. Cleanup
		Drive.Files.remove(tempSheetId);
	}
}

/**
 * Retrieves the latest File ID from the 'Files' sheet.
 */
function getLatestTemplateId()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('Files');
	if (!sheet)
	{
		return null;
	}

	const lastRow = sheet.getLastRow();
	if (lastRow <= 1)
	{
		return null;
	}

	// File ID is in the 4th column
	return sheet.getRange(lastRow, 4).getValue();
}

/**
 * Creates a Google Sheet copy of an Excel file.
 */
function createTemporarySheetFromExcel(excelFileId, title)
{
	const file = DriveApp.getFileById(excelFileId);
	const blob = file.getBlob();
	
	const resource = {
		title: title,
		mimeType: MimeType.GOOGLE_SHEETS
	};

	const tempFile = Drive.Files.insert(resource, blob);
	return tempFile.id;
}

/**
 * Normalizes a value to a string representation of a number to ignore leading zeros.
 */
function normalizeArticleId(value)
{
	if (value === null || value === undefined || value === '')
	{
		return null;
	}
	
	// Convert to string and strip non-numeric characters for potential hidden issues
	const str = String(value).trim();
	const num = parseInt(str.replace(/\D/g, ''), 10);
	return isNaN(num) ? null : String(num);
}
