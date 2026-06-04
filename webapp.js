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
	
	// Get menu metadata from CurrentMenu sheet
	const menuSheet = ss.getSheetByName('CurrentMenu');
	const menuData = {};
	if (menuSheet)
	{
		const data = menuSheet.getDataRange().getValues();
		for (let i = 1; i < data.length; i++)
		{
			menuData[data[i][0]] = data[i][1];
		}
	}

	const template = HtmlService.createTemplateFromFile('Index');
	template.menuName = menuData.menuName || 'Inconnu';
	template.pickupDateStart = menuData.pickupDateStart || 'N/A';
	template.pickupDateEnd = menuData.pickupDateEnd || 'N/A';
	
	return template.evaluate()
		.setTitle('Formulaire de Commande BDC')
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns data from CurrentArticles sheet as objects, with a cache-buster value.
 */
function getCurrentArticles()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('CurrentArticles');
	if (!sheet)
	{
		return { articles: [], menuId: '' };
	}

	const data = sheet.getDataRange().getValues();
	if (data.length <= 1)
	{
		return { articles: [], menuId: '' };
	}

	const headers = data[0];
	const result = [];
	for (let i = 1; i < data.length; i++)
	{
		const row = data[i];
		const obj = {};
		for (const [index, header] of headers.entries())
		{
			obj[header] = row[index];
		}
		result.push(obj);
	}

	const menuId = result.length > 0 ? result[result.length - 1]['Sheet Name'] : '';
	PropertiesService.getScriptProperties().setProperty('menuId', menuId);
	return { articles: result, menuId: menuId };
}

/**
 * Updates the maxDate cache.
 */
function updateCachedMaxDate()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheetOrg = ss.getSheetByName('ACStructures');
	if (!sheetOrg)
	{
		return;
	}

	let maxDate = '';
	const data = sheetOrg.getDataRange().getValues();
	const dateIndex = data[0].indexOf('Date de mise à jour');
	for (let i = 1; i < data.length; i++)
	{
		if (data[i][dateIndex] > maxDate)
		{
			maxDate = data[i][dateIndex];
		}
	}
	PropertiesService.getScriptProperties().setProperty('maxDate', maxDate);
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
	const requestedFields = ['Nom', 'UD', 'Planning', 'Modes de distribution de l\'aide alimentaire', 'Entrepôt d\'enlèvement', 'Passages Sec'];

	for (let i = 1; i < data.length; i++)
	{
		if (String(data[i][codeIndex]) !== String(codeVif))
		{
			continue;
		}

		const result = {};
		for (const field of requestedFields)
		{
			result[field] = data[i][headers.indexOf(field)];
		}
		result['Passages Sec'] = parseInt(result['Passages Sec'], 10) || 1;
		return result;
	}
	return null;
}

/**
 * Logs order to 'Orders' sheet.
 */
function logOrder(formData)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName('Orders');
	if (!sheet)
	{
		sheet = ss.insertSheet('Orders');
		sheet.appendRow(['Date', 'Email', 'Code VIF', 'Pickup Date', 'Comments', 'Articles', 'User Key']);
	}

	const row = [
		new Date(),
		formData.email,
		formData.codeVif,
		formData.pickupDate,
		formData.comments,
		JSON.stringify(formData.articles),
		Session.getTemporaryActiveUserKey()
	];
	sheet.appendRow(row);
}

/**
 * Handle form submission.
 */
function processForm(formData)
{
	try
	{
		logOrder(formData);
		const result = generateOrderExcelBlob(formData);
		const blob = result.blob;
		const logs = result.logs;
		
		// Retrieve partner organization details
		const orgData = lookupOrganization(formData.codeVif);
		let orgInfo = 'Nom: Inconnu\n';
		if (orgData)
		{
			orgInfo = Object.entries(orgData)
				.map(([key, value]) => key + ': ' + value)
				.join('\n');
		}

		const ownerEmail = Session.getEffectiveUser().getEmail();
		const orgName = orgData ? orgData['Nom'] : 'Inconnu';
		const subject = 'Nouvelle commande - ' + orgName + ' (' + formData.codeVif + ')';
		const body = 'Veuillez trouver ci-joint la commande pour le partenaire ' + orgName + ' (' + formData.codeVif + ')' + 
			'.\n\nDate d\'enlèvement prévue : ' + formData.pickupDate + 
			'\nClient : ' + formData.email +
			'\n\n--- Détails Partenaire ---\n' + orgInfo +
			'\n\nCommentaires : ' + (formData.comments || 'Aucun') +
			'\n\n--- Debug Logs ---\n' + (logs || 'No logs available.');

		MailApp.sendEmail({
			to: formData.email,
			bcc: ownerEmail,
			subject: subject,
			body: body,
			attachments: [blob]
		});

		return 'Votre commande a été envoyée avec succès.';
	}
	catch (e)
	{
		console.error('Error processing form:', e);
		throw new Error('Erreur lors de l\'envoi de la commande : ' + e.message);
	}
}
