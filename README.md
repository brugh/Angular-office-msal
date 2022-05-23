# Getting started

run ```npm install```

## Azure: 

Go to [https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps] and register new Application.  
Go to Authorization and add a platform: Single-Page application  
Enter 'https://localhost:4200' as Redirect URI. 

Check that User.Read is in the default API permissions as DELEGATED permissions. 

## Environment

create a /src/environments/environment.ts file with: 

```javascript
export const environment = {
  production: false,
  clientId: <your Application (client) ID>
  authority: 'https://login.microsoftonline.com/common/',
  redirectUri: window.location.origin,
  scopes: ['User.Read']
};
```

## Manifest

run ```npm run uuid```  
Edit manifest.xml file and add UUID.  
At the bottom, add the Azure Application Client ID  

## Certificates

Run 'npm run mkcert'  
Import (install) CA.crt and CERT.crt into your certificate store  

## Office

Copy manifest.xml to 'C:\programdata\MyAddin' (or other preferred directory).  
Inside an Office program, go to 'File', 'Options', 'Trust Center', 'Trust Center Settings', Select 'Trusted Add-in Catalogs'.  
Enter Catalog Url: ```\\localhost\c$\programdata\MyAddin```  

Restart Office program, then go to 'Insert' tab and select 'My Add-ins'. Go to Shared Folder and select your Add-in. Open Addin from Home tab.  

If your add-in is missing, check the manifest file with 'npm run validate'. 
