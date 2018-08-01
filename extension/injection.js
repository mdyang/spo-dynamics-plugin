var dynamicsInstance = 'dynsearch1.crm.dynamics.com';

var result = null;

var getUrlParameter = function getUrlParameter(sParam) {
    var sPageURL = decodeURIComponent(window.location.search.substring(1)),
        sURLVariables = sPageURL.split('&'),
        sParameterName,
        i;

    for (i = 0; i < sURLVariables.length; i++) {
        sParameterName = sURLVariables[i].split('=');

        if (sParameterName[0] === sParam) {
            return sParameterName[1] === undefined ? true : sParameterName[1];
        }
    }
};

function performSearch() {
	var crmResultContainer = $('#crmResult');
	var searchBox = $('input[placeholder="Search in SharePoint"]');
	if (crmResultContainer) {
		$('#crmResult').remove();
	}

	// var searchTerm = getUrlParameter('q').replace(/\+/g, ' ');
	var searchTerm = searchBox.val();
    console.log("fannout to Dynamics Online for search term: " + searchTerm);

    chrome.storage.local.get(['token'], function(result) {
        if (result.token) {
            $.ajax({
                method: 'post',
                url: 'https://' + dynamicsInstance + '/XRMServices/2011/Organization.svc/web',
                contentType: 'text/xml',
                datatype: 'xml', 
                headers: { 
                    'SOAPAction': 'http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/Execute',
                    'Authorization': "Bearer " + result.token
                },
                success: (data) => {
					var types = [
						{
							type: 'account',
							display: 'Accounts',
							background: '#794300',
							icon: '/_imgs/NavBar/ActionImgs/Account_32.png'
						}, 
						{
							type: 'contact', 
							display: 'Contacts',
							background: '#005088',
							icon: '/_imgs/NavBar/ActionImgs/Contact_32.png'
						},
						{
							type: 'lead', 
							display: 'Leads',
							background: '#0071C5',
							icon: '/_imgs/NavBar/ActionImgs/Lead_32.png'
						},
						{
							type: 'opportunity',
							display: 'Opportunies',
							background: '#3E7239',
							icon: '/_imgs/NavBar/ActionImgs/Opportunity_32.png'
						}];
						
					var crmGroup = $('<div id="crmResult">').css({'width': '100%', 'margin-top': '15px'});					
					crmGroup.append($('<div class="FolderGroup-module__header___3aEV4">').append($('<span style="margin-right:5px">').text('Results from ')).append($('<a>').text('Dynamics 365').attr({'href': 'https://' + dynamicsInstance, 'target': '_blank'})));
					
					var hasResult = false;
					var typesContainer = $('<div>').css('width', '100%');
					for (var i = 0; i < types.length; i ++)
					{
						var type = types[i].type;
						var display = types[i].display;
						var icon = types[i].icon;
						var background = types[i].background;
						var results = $(getTypeResult(data, type)).find('a\\:Entity');
						if (results.length) {
							hasResult = true;
						}

						var typeGroup = $('<div>').css({'width': '25%', 'display': 'inline-block', 'vertical-align': 'top'});
						typeGroup.append($('<div>').text(display).css('font-weight', "bold"));
						$(results).each((j) => {
							var result = results[j];
							var idPropName = type + 'id';
							var name = getFirstAvailablePropertyValue(result, ['fullname', 'name']);
							var id = getPropertyValue(result, idPropName);
							typeGroup.append(
								$('<div>')
									.append(
										$('<div style="display:inline-block;width:32px;height:32px;margin:5px 5px 5px 0;vertical-align:middle">')
											.css('background-color', background)
											.append($('<img>')
												.attr({
												'src': 'https://' + dynamicsInstance + icon})
											)
									)
									.append(
										$('<div style="display:inline-block;vertical-align:middle">').append($('<a>')
											.attr({
												'href': 'https://' + dynamicsInstance + '/main.aspx?etn=' + type + '&pagetype=entityrecord&id=' + id, 
												'target': '_blank'})
											.text(name))
									)
								);
							console.log('found ' + type + ': ' + name);
						});
						
						typesContainer.append(typeGroup);
					}
					
					crmGroup.append(typesContainer);
					if (hasResult) {
						$('.SPSearchUX-module__searchFilters___s1xp2').parent().after(crmGroup);
					}
				},
                data: '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Header><SdkClientVersion xmlns="http://schemas.microsoft.com/xrm/2011/Contracts">9.0</SdkClientVersion></s:Header><s:Body><Execute xmlns="http://schemas.microsoft.com/xrm/2011/Contracts/Services" xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><request i:type="b:ExecuteQuickFindRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts" xmlns:b="http://schemas.microsoft.com/crm/2011/Contracts"><a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic"><a:KeyValuePairOfstringanyType><b:key>SearchText</b:key><b:value i:type="c:string" xmlns:c="http://www.w3.org/2001/XMLSchema">{{search_keyword}}</b:value></a:KeyValuePairOfstringanyType><a:KeyValuePairOfstringanyType><b:key>EntityGroupName</b:key><b:value i:type="c:string" xmlns:c="http://www.w3.org/2001/XMLSchema">Mobile Client Search</b:value></a:KeyValuePairOfstringanyType><a:KeyValuePairOfstringanyType><b:key>EntityNames</b:key><b:value i:nil="true" /></a:KeyValuePairOfstringanyType><a:KeyValuePairOfstringanyType><b:key>AppModule</b:key><b:value i:nil="true" /></a:KeyValuePairOfstringanyType></a:Parameters><a:RequestId i:nil="true" /><a:RequestName>ExecuteQuickFind</a:RequestName></request></Execute></s:Body></s:Envelope>'.replace('{{search_keyword}}', searchTerm),
				error: () => {
					refetchToken(performSearch);
				}
            });
        }
    });
}

function getTypeResult(xml, type) {
    var resultSets = $(xml).find('a\\:QuickFindResult');
    var filteredResultSets = resultSets.filter((i) => {
        return type == $($(resultSets[i]).find('a\\:EntityName')[0]).text();
    });
    
    return filteredResultSets && filteredResultSets.length > 0 ? filteredResultSets[0] : null;
}

function getPropertyValue(entity, propName) {
    var props = $(entity).find('a\\:KeyValuePairOfstringanyType');
    var filteredProp = props.filter((i) => {
        return propName == $($(props[i]).find('c\\:key')[0]).text();
    });
    
    return filteredProp && filteredProp.length > 0 ? $(filteredProp).find('c\\:value').text() : null;
}

function getFirstAvailablePropertyValue(entity, propNames) {
    for (var i = 0; i < propNames.length; i ++)
    {
        var value = getPropertyValue(entity, propNames[i]);
        if (value)
        {
            return value;
        }
    }
    
    return null;
}

function attach() {
    const searchButton = $('button[aria-label="Search"]');
    if (!searchButton) {
        return;
    }
    searchButton.on('click', performSearch);
}

attach();
