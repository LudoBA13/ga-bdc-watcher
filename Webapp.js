/**
 * Web App entry point.
 */
function doGet()
{
	const template = HtmlService.createTemplateFromFile('Index');
	const menuData = getCurrentMenu();
	const menuId = menuData.menuId || '';
	PropertiesService.getScriptProperties().setProperty('menuId', menuId);

	menuData.articles = (menuData.articles || []).map(a =>
	{
		return {
			'Category': a.category,
			'Article ID': a.id,
			'Label': a.label,
			'Unit': a.unit,
			'Unit Weight': a.unitWeight,
			'Max Qty': a.maxQty
		};
	});

const MIN_PICKUP_DELAY = 5;

/**
 * Retrieves valid pickup dates for the current menu.
 */
function getValidPickupDates()
{
	const menuData = getCurrentMenu();
	const startDate = new Date(menuData.pickupDateStart);
	const endDate = new Date(menuData.pickupDateEnd);
	const validDates = [];

	const minPickupDate = new Date();
	minPickupDate.setDate(minPickupDate.getDate() + MIN_PICKUP_DELAY);
	minPickupDate.setHours(0, 0, 0, 0);

	let currentDate = new Date(startDate);
	while (currentDate <= endDate)
	{
		if (currentDate >= minPickupDate)
		{
			validDates.push(currentDate.toISOString().split('T')[0]);
		}
		currentDate.setDate(currentDate.getDate() + 1);
	}
	return validDates;
}

	const formatDate = (date) =>
	{
		if (!(date instanceof Date))
		{
			return date;
		}
		const months = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
		return date.getDate() + ' ' + months[date.getMonth()] + ' ' + date.getFullYear();
	};

	template.pickupDateStart = formatDate(menuData.pickupDateStart);
	template.pickupDateEnd = formatDate(menuData.pickupDateEnd);
	template.menuId = menuId;
	template.menuName = menuData.menuName || 'Inconnu';
	template.maxDate = PropertiesService.getScriptProperties().getProperty('maxDate') || '';
	template.validPickupDates = getValidPickupDates();

	// Embed raw data for immediate client-side rendering
	template.embeddedMenuData = JSON.stringify(menuData);

	return template.evaluate()
		.setTitle('Formulaire de Commande BA 13')
		.setFaviconUrl('https://ba13.banquealimentaire.org/themes/custom/customer/favicon.ico')
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns the data from the 'CurrentMenu' sheet as an object.
 *
 * @returns {Object} The current menu data.
 */
function getCurrentMenu()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('CurrentMenu');
	if (!sheet)
	{
		return {};
	}

	const data = sheet.getDataRange().getValues();
	const result = {};

	for (let i = 1; i < data.length; i++)
	{
		const key = data[i][0];
		let val = data[i][1];
		if (key === 'articles' && typeof val === 'string' && val.startsWith('['))
		{
			try
			{
				val = JSON.parse(val);
			}
			catch (e)
			{
				console.error('Error parsing articles JSON:', e);
			}
		}
		result[key] = val;
	}

	return result;
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
	const requestedFields = ['Nom', 'UD', 'Planning', '$planning', 'Modes de distribution de l\'aide alimentaire', 'Entrepôt d\'enlèvement', 'Passages Sec', 'Le partenaire reçoit-il des produits CNES ?', 'Le partenaire reçoit-il des produits du FSE + ?'];

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

		// Retrieve partner organization details
		const orgData = lookupOrganization(formData.codeVif);
		let orgInfo = 'Nom: Inconnu\n';
		if (orgData)
		{
			orgInfo = Object.entries(orgData)
				.map(([key, value]) => key + ': ' + value)
				.join('\n');
		}

		const menuData = getCurrentMenu();
		const serverObservations = generateObservations(formData, menuData, orgData);
		Logger.log('Server-side observations for ' + formData.codeVif + ': ' + JSON.stringify(serverObservations));

		// Use client-side observations passed in formData
		const clientObsText = (formData.observations && formData.observations.length > 0) ? formData.observations.map(o => '- ' + o).join('\n') : 'Aucune';

		const ownerEmail = Session.getEffectiveUser().getEmail();
		const orgName = orgData ? orgData['Nom'] : 'Inconnu';
		const subject = 'Nouvelle commande - ' + orgName + ' (' + formData.codeVif + ')';
		const body = 'Veuillez trouver ci-joint la commande pour le partenaire ' + orgName + ' (' + formData.codeVif + ')' +
			'.\n\nDate d\'enlèvement prévue : ' + formData.pickupDate +
			'\nClient : ' + formData.email +
			'\n\n--- Observations ---\n' + clientObsText +
			'\n\n--- Détails Partenaire ---\n' + orgInfo +
			'\n\nCommentaires : ' + (formData.comments || 'Aucun');

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
		throw new Error('Erreur lors de l\'envoi de la commande : ' + e.message);
	}
}
