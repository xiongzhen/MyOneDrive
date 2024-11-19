let g_nUpdateTableID = 0;

function on_click_driveitem(e) {
    let target = e.target;
    if(target.tagName == 'LABEL') {
        target = target.parentNode;
    }
    if(target.tagName != 'TD') {
        return;
    }
    
    let weburl = target.getAttribute('weburl');
    let shared = weburl ? true : false;

    chrome.storage.local.get(['access_token'], function(items) {
        let access_token = items['access_token'];
        let headers = new Headers({
            'Authorization': 'Bearer ' + access_token
        });

        let isfile = target.getAttribute('isfile');
        let fileid = target.id;
        let url = null;
        if(shared) {
            url = "https://graph.microsoft.com/v1.0/shares/u!";
            weburl64 = btoa(weburl).replaceAll('+','-').replaceAll('/', '_').replaceAll('=', '');
            url += weburl64 + "/driveItem";
        } else {
            url = "https://graph.microsoft.com/v1.0/me/drive/items/";
            url += fileid;
        }
        if (isfile == true || isfile == "true") {
            chrome.storage.local.get(['access_token'], function(items) {
                let access_token = items['access_token'];
                let headers = new Headers({
                    'Authorization': 'Bearer ' + access_token
                });
                if(shared) {
                    headers.append('Prefer', 'redeemSharingLinkIfNecessary');
                }

                fetch(url, {headers: headers})
                .then(r => r.text())
                .then(result => {
                    let item = JSON.parse(result);
                    let downloadUrl = item['@content.downloadUrl'] ?? item['@microsoft.graph.downloadUrl'];
                    if(downloadUrl) {
                        let downloadLink = document.createElement('a');
                        downloadLink.href = downloadUrl;
                        downloadLink.dowload = item.name;
                        document.body.appendChild(downloadLink);
                        downloadLink.click();
                        document.body.removeChild(downloadLink);
                    }
                });
            });
        } else {
            let dynamicStyle = document.getElementById("DynamicStyle");
            let tag = btoa(fileid) + "@" + g_nUpdateTableID;
            g_nUpdateTableID += 1;
            dynamicStyle.innerText = "#DriveItemsTableBody tr[ParentID=\"" + tag + "\"] { display: normal; }";
            dynamicStyle.innerText += "#DriveItemsTableBody tr:not([ParentID=\"" + tag + "\"]) { display: none; }";

            url += "/children";
            let list_header = document.getElementById('list-header');
            let list_header_item = document.createElement("span");
            list_header.appendChild(list_header_item);
            list_header_item.setAttribute('class', 'list-header-item');
            list_header_item.setAttribute('id', fileid);
            if(weburl) {
                list_header_item.setAttribute('weburl', weburl);
            }
            list_header_item.innerText = target.innerText;
            list_header_item.addEventListener('click', navigate_to_folder);

            update_table(url, true, tag, shared);
        }
    });



}

function on_click_upload(e) {

    chrome.storage.local.get(['access_token'], function(items) {
        let access_token = items['access_token'];
        let headers = new Headers({
            'Authorization': 'Bearer ' + access_token,
            'Content-Type': 'application/json'
        });
        let body = JSON.stringify({
            item: {
                '@microsoft.graph.conflictBehavior': 'rename'
            },
            deferCommit: true
        });

        let current_folder = document.querySelector("#list-header .list-header-item:last-child");
        if(!current_folder) {
            return;
        }
        let folder_id = current_folder.id;
        let url = "https://graph.microsoft.com/v1.0/me/drive/items/";
        url += folder_id;
        url += "/test.txt:/createUploadSession";
        fetch(url, {
            method: 'POST',
            headers: headers,
            body: body
        }).then(r => r.text())
        .then(result => {
            console.log(result);
        });
        
    });

}


function navigate_to_folder(e) {
    let target = e.target;
    if(target.tagName != "SPAN") {
        return;
    }
    let list_header = target.parentNode;
    while(true) {
        let next = target.nextSibling;
        if(!next) {
            break;
        }
        list_header.removeChild(next);
    }

    let weburl = target.getAttribute('weburl');

    let id = weburl ?? target.id;
    let tag = btoa(id) + "@" + g_nUpdateTableID;
    g_nUpdateTableID += 1;
    
    let dynamicStyle = document.getElementById("DynamicStyle");
    dynamicStyle.innerText = "#DriveItemsTableBody tr[ParentID=\"" + tag + "\"] { display: normal; }";
    dynamicStyle.innerText += "#DriveItemsTableBody tr:not([ParentID=\"" + tag + "\"]) { display: none; }";
    
    let url = null;
    if(weburl) {
        url = "https://graph.microsoft.com/v1.0/shares/u!"
        url += btoa(weburl).replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');

        let shareroot = target.getAttribute('shareroot');
        if(shareroot == true || shareroot == 'true') {
            url += "/driveItem"
        } else {
            //url += "/driveItem?$expand=children";
            url += "/driveItem/children";
        }
    } else {
        url = "https://graph.microsoft.com/v1.0/me/drive/items/";
        url += id;
        url += "?$expand=children";
    }
    update_table(url, true, tag, weburl ? true : false);
}

function create_one_row(driveitem, tag, shared) {
    let is_folder = driveitem.hasOwnProperty('folder');
    let is_file = driveitem.hasOwnProperty('file');
    let is_bundle = driveitem.hasOwnProperty('bundle');

    let row = document.createElement('tr');
    row.setAttribute('ParentID', tag);

    let cell_name = document.createElement('td');
    let cell_name_label = document.createElement('label');
    cell_name_label.innerText = driveitem.name;
    if(is_folder) {
        cell_name.setAttribute('class', 'folder');
    } else if(is_file) {
        cell_name.setAttribute('class', 'file');
    } else if(is_bundle) {
        cell_name.setAttribute('class', 'bundle');
    }
    cell_name.addEventListener('click', on_click_driveitem);
    cell_name.appendChild(cell_name_label);
    cell_name.setAttribute('isfile', is_file);
    if(shared) {
        cell_name.setAttribute('weburl', driveitem.webUrl);
    } else {
        cell_name.id = driveitem.id;
    }
    row.appendChild(cell_name);

    let cell_size = document.createElement('td');
    let nSize = driveitem.size;
    if(nSize <= 1) {
        cell_size.innerText = nSize + " byte";
    } else if(nSize < 1024) {
        cell_size.innerText = nSize + " bytes";
    } else if(nSize < 1024 * 1024) {
        cell_size.innerText = (nSize / 1024.0).toFixed(2) + " KB";
    } else if(nSize < 1024 * 1024 * 1024) {
        cell_size.innerText = (nSize / 1024.0 / 1024.0).toFixed(2) + " MB";
    } else {
        cell_size.innerText = (nSize / 1024.0 / 1024.0 / 1024.0).toFixed(2) + " GB";
    }
    if(is_folder || is_bundle) {
        let nItems = is_folder ? driveitem.folder.childCount : driveitem.bundle.childCount;
        if(nItems <= 1) {
            cell_size.innerText += " (" + nItems + " item)";
        } else {
            cell_size.innerText += " (" + nItems + " items)";
        }
    }
    row.appendChild(cell_size);

    let cell_type = document.createElement('td');
    if(driveitem.hasOwnProperty('folder')) {
        cell_type.innerText = 'File Folder';
    } else if(driveitem.hasOwnProperty('bundle')) {
        cell_type.innerText = 'Bundle';
    } else if(driveitem.hasOwnProperty('file')) {
        cell_type.innerText = 'File';
    }
    row.appendChild(cell_type);

    let cell_modified = document.createElement('td');
    row.appendChild(cell_modified);
    cell_modified.innerText = driveitem.lastModifiedDateTime;

    let cell_modified_by = document.createElement('td');
    row.appendChild(cell_modified_by);
    cell_modified_by.innerText = driveitem.lastModifiedBy.user.displayName;

    return row;
}

function update_table(url, clear, tag, shared) {
    let tbody = document.getElementById('DriveItemsTableBody');
    if(clear) {
        tbody.innerHTML = "";
    }
    
    chrome.storage.local.get(['access_token'], function(items) {
        let access_token = items['access_token'];
        let headers = new Headers({
            'Authorization': 'Bearer ' + access_token
        });
        if(shared) {
            headers.append("Prefer", "redeemSharingLinkIfNecessary");
        }
        try {
            fetch(url, {headers: headers})
            .then(r => r.text())
            .then(result => {
                let items = JSON.parse(result);

                if(items.hasOwnProperty('error')) {
                    console.log(items);
                    return;
                }

                let values = items['children'];
                if(!values) {
                    values = items['value'];
                }

                if(values) {
                    for (const key in values) {
                        let driveitem = values[key];

                        let row = create_one_row(driveitem, tag, shared);
                        tbody.appendChild(row);
                    }

                    let nextLink = items['@odata.nextLink'] ?? items['children@odata.nextLink'];
                    if(nextLink != null) {
                        update_table(nextLink, false, tag, shared);
                    }
                } else {
                    let row = create_one_row(items, tag, shared);
                    tbody.appendChild(row);
                }
            });
        } catch(err) {
            console.log(err);
        }
    });
}

//function clear_access_token() {
//    chrome.storage.local.set({access_token: null}, function(){
//        console.log("access token cleared");
//    });
//}

document.getElementById('go-sharing-url').addEventListener('click', (e) => {
    let list_header = document.getElementById('list-header');
    list_header.innerHTML = "";
    
    let url = document.getElementById('sharing-url').value;
    if(!url) {
        init();
        return;
    }

    url64 = btoa(url);
    url64 = url64.replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');

    let tag = url64 + "@" + g_nUpdateTableID;
    g_nUpdateTableID += 1;
    let dynamicStyle = document.getElementById("DynamicStyle");
    dynamicStyle.innerText = "#DriveItemsTableBody tr[ParentID=\"" + tag + "\"] { display: normal; }";
    dynamicStyle.innerText += "#DriveItemsTableBody tr:not([ParentID=\"" + tag + "\"]) { display: none; }";

    
    let tbody = document.getElementById('DriveItemsTableBody');
    tbody.innerHTML = "";

    chrome.storage.local.get(['access_token'], function(items) {
        let access_token = items['access_token'];
        let headers = new Headers({
            'Authorization': 'Bearer ' + access_token,
            'Prefer': 'redeemSharingLinkIfNecessary'
        });
        
        fetch("https://graph.microsoft.com/v1.0/shares/u!" + url64 + "/driveItem?$select=*", {headers: headers})
        .then(r => r.text())
        .then(result => {
            let driveitem = JSON.parse(result);
            if(driveitem.hasOwnProperty('error')) {
                let msg = driveitem.error.code;
                msg += '\n\n';
                msg += driveitem.error.message;
                alert(msg);
                return;
            }

            let list_header_item = document.createElement("span");
            list_header.appendChild(list_header_item);
            list_header_item.setAttribute('class', 'list-header-item');
            list_header_item.setAttribute('weburl', url);
            list_header_item.innerText = 'Shared files';
            list_header_item.addEventListener('click', navigate_to_folder);
            list_header_item.setAttribute('shareroot', true);

            //if(driveitem.hasOwnProperty('file') || driveitem.hasOwnProperty('folder')) {
                let row = create_one_row(driveitem, tag, true);
                tbody.appendChild(row);
            //} else {
            //    url = "https://graph.microsoft.com/v1.0/shares/u!" + url64 + "/driveItem/children";
            //    update_table(url, true, tag, true);
            //}
        });
    });

});

function init() {
    chrome.storage.local.get(['access_token'], function(items) {
        let access_token = items['access_token'];
        let headers = new Headers({
            'Authorization': 'Bearer ' + access_token
        });

        fetch('https://graph.microsoft.com/v1.0/me', {headers: headers})
        .then(r => r.text())
        .then(result => {
            let div_owner = document.getElementById('owner');

            let drive = JSON.parse(result);
            if(drive.hasOwnProperty('error')) {
                div_owner.innerText = "";    
                return;
            }
            clearInterval(init_id);
            div_owner.innerText = drive.userPrincipalName;
            
            let list_header = document.getElementById('list-header');
            list_header.innerHTML = "";

            let list_header_item = document.createElement("span");
            list_header.appendChild(list_header_item);
            list_header_item.setAttribute('class', 'list-header-item');
            list_header_item.setAttribute('id', 'root:');
            list_header_item.innerText = 'My files';
            list_header_item.addEventListener('click', navigate_to_folder);

            let tag = "@" + g_nUpdateTableID;
            g_nUpdateTableID += 1;
            let dynamicStyle = document.getElementById("DynamicStyle");
            dynamicStyle.innerText = "#DriveItemsTableBody tr[ParentID=\"" + tag + "\"] { display: normal; }";
            dynamicStyle.innerText += "#DriveItemsTableBody tr:not([ParentID=\"" + tag + "\"]) { display: none; }";

            update_table('https://graph.microsoft.com/v1.0/me/drive/items/root:?$expand=children', true, tag, false);
        });
    });
}

let init_id = setInterval(init, 1000);