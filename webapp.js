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
