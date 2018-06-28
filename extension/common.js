function refetchToken(callback)
{
	$.get('http://meyang-hp2:8080/api/Token/1', (data) => { 
		chrome.storage.local.set({token: data}, function() {
			callback();
		});
	});
}