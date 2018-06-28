function success()
{
    $('#checking').hide();
    $('#itworks').show();
}

function fail()
{
    $('#checking').hide();
    $('#itworks').hide();
}

function test()
{
    chrome.storage.local.get(['token'], function(result) {
        if (result.token) {
            $.ajax({
                url: "https://mengdongy2.crm5.dynamics.com/api/data/v9.0", 
                headers: { Authorization: "Bearer " + result.token },
                success: success,
                error: refetchToken
            });
        }
		else {
			refetchToken(test);
		}
    });
}

$(document).ready(() => {
    test();
});