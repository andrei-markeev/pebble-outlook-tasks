var UI = require('ui');
var ajax = require('ajax');
var Settings = require('settings');
var Feature = require('platform/feature');

var appUrl = 'https://markeev.com/pebble/tasks.html';

var clientId='d20692cf-c8fc-4fb9-9ee4-d2eb877d669c'; // Azure app id (https://manage.windowsazure.com/)
var clientSecret=''; // Azure app secret
var auth_url = 'https://login.microsoftonline.com/common/oauth2/token?api-version=beta';
var baseApiUrl = 'https://graph.microsoft.com/beta';

Settings.config({ url: appUrl }, function() {
    localStorage.removeItem('refresh_token');
    console.log('Settings closed!');
    retrieveTasks();
});

if (!Settings.option('auth_code'))
{
    showText('Account is not configured. Please visit app settings on phone.');
    Pebble.openURL(appUrl);
    return;
}
else
    retrieveTasks();

function retrieveTasks()
{
  if (localStorage.getItem('refresh_token') !== null)
    authWithRefreshToken();
  else
    authWithAccessCode();
   
}

function authWithRefreshToken()
{
    console.log('Auth with refresh token');


    ajax({
            url: auth_url,
            method: 'POST',
            data: {
                grant_type: 'refresh_token',
                refresh_token: localStorage.getItem('refresh_token'),
                redirect_uri: appUrl,
                client_id: clientId,
                client_secret: clientSecret
            }
        },
        authorizedCallback,
        showError
    );
}

function authWithAccessCode()
{
    console.log('Auth with access code');

    ajax({
            url: auth_url,
            method: 'POST',
            data: {
                grant_type: 'authorization_code',
                code: Settings.option('auth_code'),
                redirect_uri: appUrl,
                client_id: clientId,
                client_secret: clientSecret
            }
        },
        authorizedCallback,
        showError
    );
}

var access_token = '';
function authorizedCallback(data)
{
    var dataObj = JSON.parse(data);
    access_token = dataObj.access_token;
    var refresh_token = dataObj.refresh_token;
    localStorage.setItem('refresh_token', refresh_token);
    
    // get list of task folders
    ajax(
      {
        url: baseApiUrl + '/me/outlook/taskFolders',
        headers: { "Authorization": "Bearer " + access_token }
      },
      function(data) {
        showMenu('Folders', data, 'name', folderSelected);
      },
      showError
    );
}

var folder_id = '';
function folderSelected(e, id)
{
    folder_id = id;

    // get tasks in folder
    ajax(
        {
            url: baseApiUrl + '/me/outlook/taskFolders/' + folder_id + "/tasks?$filter=status ne 'completed'&$orderby=importance,createdDateTime asc",
            headers: { "Authorization": "Bearer " + access_token }
        },
        function(data) {
          showMenu('Tasks', data, 'subject', taskSelected);
        },
        showError
    );
}

var task_id = '';
function taskSelected(e, id)
{
    task_id = id;
    
    var card = new UI.Card({
        body: e.item.title + "\n\n* Long click select to mark as completed!",
        scrollable: true
    });
    card.on('longClick', 'select', function(){ markAsCompleted(card, e.menu); });
    card.show();
}

function markAsCompleted(card, menu) {
    console.log('POST ' + baseApiUrl + '/me/outlook/tasks/' + task_id + '/complete');
    
    ajax(
        {
            method: 'POST',
            url: baseApiUrl + '/me/outlook/tasks/' + task_id + '/complete',
            headers: { "Authorization": "Bearer " + access_token }
        },
        function(data) {
            card.hide();
            menu.hide();
            folderSelected(null, folder_id);
        },
        showError
    );
}


function showMenu(title, data, title_property, select_callback)
{
    var dataObj = JSON.parse(data);
    
    if (dataObj.value.length === 0) {
        showText(title + ' not found.');
        return;
    }

    var menuitems = [];
    var ids = [];
    for (var i=0;i<dataObj.value.length;i++)
    {
        menuitems.push({
            title: dataObj.value[i][title_property] || '(untitled)'
        });
        
        ids.push(dataObj.value[i].id);
    }
    var menu = new UI.Menu({
        highlightBackgroundColor: Feature.color('#FFAA00', 'black'),
        sections: [
            {
                title: title,
                items: menuitems
            }
        ]
    });
    
    menu.on('select', function(e) { select_callback(e, ids[e.itemIndex]); });
    menu.show();
}

function showError()
{
    var main = new UI.Card({
        title: 'Error',
        body: JSON.stringify(arguments),
        scrollable: true
    });
    main.show();
}

function showText(text)
{
    var main = new UI.Card({
        body: text,
        scrollable: true
    });
    main.show();
    
}


