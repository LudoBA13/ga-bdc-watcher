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
	return HtmlService.createTemplateFromFile('Index').evaluate()
		.setTitle('Formulaire de Commande BDC')
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns data from Articles sheet as objects.
 */
function getCurrentArticles()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('Articles');
	if (!sheet)
	{
		return [];
	}

	const data = sheet.getDataRange().getValues();
	if (data.length <= 1)
	{
		return [];
	}

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
	return result;
}

/**
 * Mock lookup for Code VIF.
 */
function lookupOrganization(codeVif)
{
	if (codeVif === '1139998')
	{
		return {
			name: 'Test AP',
			ud: 100,
			planning: '1er lundi 8h30'
		};
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
