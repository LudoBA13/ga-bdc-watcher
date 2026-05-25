/**
 * Generates an Excel order from a template and sends it by email.
 * 
 * @param {Object} orderData The order details.
 * @param {string} orderData.email The user's email.
 * @param {string} orderData.codeVif The partner code.
 * @param {string} orderData.pickupDate The pickup date (YYYY-MM-DD).
 * @param {Object} orderData.articles Map of Article ID (string) to Quantity (number/string).
 */
function sendExcelOrder(orderData)
{
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
		const quantities = [];

		for (let i = 0; i < ids.length; i++)
		{
			const cellValue = ids[i][0];
			const articleId = normalizeArticleId(cellValue);
			
			if (articleId && orderData.articles[articleId])
			{
				quantities.push([orderData.articles[articleId]]);
			}
			else
			{
				quantities.push(['']);
			}
		}

		// Write quantities to column H (8th column)
		sheet.getRange(40, 8, quantities.length, 1).setValues(quantities);

		// 4. Force calculation if needed (SpreadsheetApp usually handles this)
		SpreadsheetApp.flush();

		// 5. Send email to owner with Excel attachment
		emailExcelFile(ss, orderData);
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
 * Normalizes a cell value to a numeric Article ID if possible.
 */
function normalizeArticleId(value)
{
	if (typeof value === 'number')
	{
		return String(Math.floor(value));
	}
	
	const str = String(value).trim();
	if (/^\d+$/.test(str))
	{
		return str;
	}
	
	return null;
}

/**
 * Emails the spreadsheet as an Excel file to the owner.
 */
function emailExcelFile(spreadsheet, orderData)
{
	const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?format=xlsx';
	const token = ScriptApp.getOAuthToken();
	const response = UrlFetchApp.fetch(url, {
		headers: {
			'Authorization': 'Bearer ' + token
		}
	});

	const fileName = 'Commande_' + orderData.codeVif + '_' + orderData.pickupDate + '.xlsx';
	const blob = response.getBlob().setName(fileName);

	const recipient = Session.getEffectiveUser().getEmail();
	const subject = 'Nouvelle commande - ' + orderData.codeVif;
	const body = 'Veuillez trouver ci-joint la commande pour le partenaire ' + orderData.codeVif + 
		'.\n\nDate d\'enlèvement prévue : ' + orderData.pickupDate + 
		'\nClient : ' + orderData.email;

	MailApp.sendEmail({
		to: recipient,
		subject: subject,
		body: body,
		attachments: [blob]
	});
}
