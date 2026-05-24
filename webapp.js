/**
 * Retrieves the content of the CurrentArticles sheet.
 */
function getCurrentArticles()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('CurrentArticles');
	if (!sheet)
	{
		return [];
	}

	const data = sheet.getDataRange().getValues();
	if (data.length <= 1)
	{
		return [];
	}

	// Returns data without the header row
	return data.slice(1);
}

/**
 * Web App entry point.
 */
function doGet()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	
	// Get maxDate from ACStructures
	let maxDate = '';
	const sheetOrg = ss.getSheetByName('ACStructures');
	if (sheetOrg)
	{
		const data = sheetOrg.getDataRange().getValues();
		const dateIndex = data[0].indexOf('Date de mise à jour');
		for (let i = 1; i < data.length; i++)
		{
			if (data[i][dateIndex] > maxDate)
			{
				maxDate = data[i][dateIndex];
			}
		}
	}

	// Get menuId from Articles
	let menuId = '';
	const sheetArt = ss.getSheetByName('Articles');
	if (sheetArt)
	{
		const lastRow = sheetArt.getLastRow();
		if (lastRow > 1)
		{
			menuId = sheetArt.getRange(lastRow, 1).getValue();
		}
	}

	const template = HtmlService.createTemplateFromFile('Index');
	template.maxDate = maxDate;
	template.menuId = menuId;
	return template.evaluate()
		.setTitle('Formulaire de Commande BDC')
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns data from Articles sheet as objects, with a cache-buster value.
 */
function getCurrentArticles()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('Articles');
	if (!sheet)
	{
		return { articles: [], menuId: '' };
	}

	const data = sheet.getDataRange().getValues();
	if (data.length <= 1)
	{
		return { articles: [], menuId: '' };
	}

	const menuId = sheet.getRange(sheet.getLastRow(), 1).getValue();
	const headers = data[0];
	const result = [];
	for (let i = 1; i < data.length; i++)
	{
		const row = data[i];
		const obj = {};
		headers.forEach((header, index) =>
		{
			obj[header] = row[index];
		});
		result.push(obj);
	}
	return { articles: result, menuId: menuId };
}

/**
 * Look up organization data from ACStructures sheet.
 */
function lookupOrganization(codeVif)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('ACStructures');
	if (!sheet)
	{
		return null;
	}

	const data = sheet.getDataRange().getValues();
	const headers = data[0];
	const codeIndex = headers.indexOf('Code VIF');
	const requestedFields = ['Nom', 'UD', 'Planning', 'Modes de distribution de l\'aide alimentaire'];

	for (let i = 1; i < data.length; i++)
	{
		if (String(data[i][codeIndex]) !== String(codeVif))
		{
			continue;
		}

		const result = {};
		requestedFields.forEach(field =>
		{
			result[field] = data[i][headers.indexOf(field)];
		});
		return result;
	}
	return null;
}

/**
 * Handle form submission.
 */
function processForm(formData)
{
	console.log('Form received:', formData);
	return 'Commande reçue pour ' + formData.email;
}
