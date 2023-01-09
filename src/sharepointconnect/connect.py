import json
import requests

requests.packages.urllib3.disable_warnings()


class SharePoint:
    def __init__(self, **kwargs):
        self.accesskey = None
        self.tenantid = None
        self.sub_site = kwargs["site_name"]
        self.site_url = f"https://{kwargs['domain']}.sharepoint.com/sites/{kwargs['site_name']}"

    def get_access_key(self, **kwargs):
        headers = {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        payload = {
            'grant-type': kwargs['grant_type'],
            'client_id': kwargs['client_id'],
            'client_secret': kwargs['client_secret'],
            'resource': kwargs['resource']
        }
        response = requests.post(url=kwargs['url'], headers=headers, data=payload)
        if response.status_code == 200:
            self.accesskey = response.json().get('access-token')
        else:
            self.accesskey = None


    # Files and Folders retrieval from SharePoint
    def get_file(self, **kwargs):
        url = self.site_url + f"/_api/web/GetFileByServerRelativeUrl('/sites/{self.sub_site}/{kwargs['file_path_remote']}')/$value?binaryStringResponseBody=true"
        headers = {
            'Accept': 'application/json;odata=verbose',
            'Authorization': f'Bearer {self.accesskey}'
        }
        try:
            response = requests.get(url=url, headers=headers)
        except Exception as err:
            return {'status': None, 'content': None}
        else:
            if reponse.status_code == 200:
                open(kwargs['file_path_local'], 'wb').write(response.content)
            return {'status': response.status_code, 'content': response.content}
        return {'status': None, 'content': None}


    def put_file(self, **kwargs):
        if 'folder' not in kwargs.keys():
            kwargs['folder'] = 'Shared%20Documents'
        url = self.site_url + f"/_api/web/GetFolderByServerRelativeUrl('/sites/{self.sub_site}/{kwargs['folder']}')/Files/add(url='{kwargs['filename']}',overwrite=true"
        headers = {
            'Accept': 'application/json;odata=verbose',
            'Authorization': f'Bearer {self.accesskey}'
        }
        with open(kwargs['filepath'], 'rb') as infile:
            try:
                response = requests.post(url=url, datafile=infile, headers=headers)
            except Exception as err:
                return {'status': None, 'content': err}
            else:
                return {'status': response.status_code, 'content': response.content}
        return {'status': None, 'content': None}

    
    # SharePoint lists
    def create_list_item(self, **kwargs):
        url = self.site_url + f"/_api/Web/Lists/getbytitle('{kwargs['list_title']}')/items"
        headers = {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {self.accesskey}'
        }
        try:
            response = request.post(url=url, data=json.dumps(kwargs['payload']), headers=headers)
        except Exceptions as err:
            return {'status': None, 'content': err}
        else:
            return {'status': response.status_code, 'content': response.content}
        return {'status': None, 'content': None}    


    def delete_list_item(self, **kwargs):
        url = self.site_url + f"/_api/Web/Lists/getbytitle('{kwargs['list_title']}')/items/getbyid({kwargs['id']})"
        headers = {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {self.accesskey}',
            'If-Match': '*'
        }
        try:
            response = requests.delete(url=url, headers=headers)
        except Exception as err:
            return {'status': None, 'content': err}
        else:
            return {'status': response.status_code, 'content': response.text}
        return {'status': None, 'content': None}


    def get_list_item(self, **kwargs):
        url = self.site_url + f"/_api/Web/Lists/getbytitle('{kwargs['list_title']}')/items/getbyid({kwargs['id']})"
        headers = {
            'Accept': 'application/json;odata=verbose',
            'Authorization': f'Bearer {self.accesskey}'
        }
        try:
            response = requests.get(url=url, headers=headers)
        except Exception as err:
            return {'status': None, 'content': err}
        else:
            if response.status_code == 200:
                return {'status': response.status_code, 'content': response.json()}
            else:
                return {'status': response.status_code, 'content': response.text}
        return {'status': None, 'content': None}


    def get_list_item_filter(self, **kwargs):
        url = site.site_url + f"/_api/Web/Lists/getbytitle('{kwargs['list_title']}')/items?$filter={kwargs['query']}"
        headers = {
            'Accept': 'application/json;odata=verbose',
            'Authorization': f'Bearer {self.accesskey}'
        }
        try:
            response = requests.get(url=url, headers=headers)
        except Exception as err:
            return {'status': None, 'content': err}
        else:
            if response.status_code == 200:
                return {'status': response.status_code, 'content': response.json()}
            else:
                return {'status': response.status_code, 'content': response.text}
        return {'status': None, 'content': None}


    def put_list_item_file(self, **kwargs):
        url = self.site_url + f"/_api/Web/Lists/getbytitle('{kwargs['list_title']}')/items({kwargs['id']})/AttachmentFiles/add(filename=\'{kwargs['filename']}\')"
        headers = {
            'Accept': 'application/json;odata=verbose',
            'Authorization': f'Bearer {self.accesskey}'
        }
        with open(kwargs['filepath'], 'rb') as infile:
            try:
                response = requests.get(url=url, headers=headers, data=infile)
            except Exception as err:
                return {'status': None, 'content': err}
            else:
                return {'status': response.status_code, 'content': response.text}
        return {'status': None, 'content': None}


    def update_list_item(self, **kwargs):
        url = self.site_url + f"/_api/Web/Lists/getbytitle('{kwargs['list_title']}')/items({kwargs['id']})"
        headers {
            'Accept': 'application/json;odata=verbose',
            'Authorization': f'Bearer {self.accesskey}',
            'Content-Type': 'application/json',
            'If-Match': '*',
            'X-HTTP-Method': 'MERGE'
        }
        try:
            response = requests.post(url=url, data=kwargs['data'], headers=headers)
        except Exception as err:
            return {'status': None, 'content': err}
        else:
            return {'status': response.status_code, 'content': response.text}
        return {'status': None, 'content': None}