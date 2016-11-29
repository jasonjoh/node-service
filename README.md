# Accessing Outlook APIs with Client Credentials flow in Node.js #

**THIS IS A WORK IN PROGRESS**

## Prerequisites ##

The app has the following dependencies:

- 
[Windows Azure Active Directory Authentication Library (ADAL) for Node.js](https://github.com/AzureAD/azure-activedirectory-library-for-nodejs)
- [nconf](https://github.com/indexzero/nconf)

Run `npm install` to install dependencies.

## Register the app ##

We'll start by [registering the app in Azure AD](https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md).

Add your application as a web application.

- Sign-on URL: URL to your "Sign up my org" page. If you are not planning on making this app available outside of your organization, this value isn't used, so any valid URL will work here.

Configure the application.

- Application is multi-tenant: Set to **Yes** if you want this to be used by other Office 365 organizations.
- Permissions to other applications: Add **Office 365 Exchange Online**. Under **Application Permissions**, choose **Read mail in all mailboxes**.

## Configuring the certificate ##

The Outlook APIs require the use of a certificate to acquire tokens when using the client credentials OAuth2 flow. You need to upload the public key from your certificate to your applications Azure AD registration, then use the private key from the same certificate when requesting access tokens. You can use a self-signed certificate for this, which is what we'll do for this sample. The steps outlined here will work on Windows 8.1. For those on different platforms, I will try to provide as much details as I can so you can get the same results.

### Create the certificate ###

Visual Studio 2013 includes a command line tool called [makecert.exe](https://msdn.microsoft.com/en-us/library/bfsktky3(v=vs.110).aspx) which we'll use to make the certificate.

> **NOTE**: As an alternative to using `makecert`, you can use the `New-SelfSignedCertificate` cmdlet in Windows Powershell. To do that, enter the following command in Powershell:
>
> ``` Shell
> New-SelfSignedCertificate -DnsName "MyCompanyName MyAppName Cert" -CertStoreLocation "Cert:\CurrentUser\My"
> ```
>
> Then skip to step 3 below.

The requirements for the certificate are that the private key is marked as exportable and the key length is a **minimum** of 2048 bits. We will create both a public key file (base-64 encoded CER file) and a private key file (PKCS #12 PFX file).

1. Open your **Developer Command Prompt for VS2013**. If you can't find this, look in `%programfiles(x86)%\Microsoft Visual Studio 12.0\Common7\Tools\Shortcuts`.
2. Run the following command: `makecert -r -pe -n "CN=MyCompanyName MyAppName Cert" -b <TODAY'S DATE> -e <SOME DATE IN THE FUTURE> -ss my -len 2048`. Replace `MyCompanyName MyAppName` with something descriptive for your app. Replace `<TODAY'S DATE>` with the current date in the format `mm/dd/yyyy`, and replace `<SOME DATE IN THE FUTURE>` with the date you want this certificate to expire (in the same format).
3. After running the command, the new certificate is installed in your user account's certificate store. Open the Certificate Manager (on Windows 8.1, hit the Windows key, type "Manage user certificates", and choose the first result). Expand **Personal/Certificates**, then locate the certificate by the name you gave in the `-n` parameter of the `makecert` command.
4. Right-click the certificate, choose **All Tasks**, the **Export...**. Click **Next**, select **No, do not export the private key** and click **Next**. Choose **Base-64 encoded X.509 (.CER)** and click **Next**. Provide a file name and click **Next**. For convenience, I recommend saving it as `appcert.cer` in the `./certificates` subfolder in the root of the Node.js project. Click **Finish**.
5. Right-click the certificate, choose **All Tasks**, the **Export...**. Click **Next**, select **Yes, export the private key** and click **Next**. Choose **Personal Information Exchange - PKCS #12 (.PFX)**, select the **Include all certificates in the certification path if possible** option, and click **Next**. Choose **Password**, provide a secure password, and click **Next**. Provide a file name and click **Next**. For convenience, I recommend saving it as `appcert.pfx` in the `./certificates` subfolder in the root of the Node.js project. Click **Finish**.
6. We need to extract the private key from the PFX file into a PEM file for use with the `adal-node` library. For this we'll use [OpenSSL](http://www.openssl.org/). Run the following command. If you did not use the recommended name `appcert.pfx`, change the name of the PFX file accordingly: `openssl pkcs12 -in appcert.pfx -nocerts -out private-key.pem`.
7. Finally, we need to remove the passphrase from the PEM file so `adal-node` can use it. Run the following command: `openssl rsa -in private-key.pem -out private-key-rsa.pem`.

### Upload the certificate to Azure AD ###

The base-64 encoded certificate file is uploaded to Azure AD as part of the application's manifest. The manifest is a JSON file, and the certificate details go in the `keyCredentials` value. In order to generate this value, we need the certificate hash from the .CER file we generated earlier, and the raw base-64 contents of the file.

For Windows systems with Powershell installed, this repo includes the `Get-KeyCredentials.ps1` script to automate getting those values and building the `keyCredentials` value.

Open Windows Powershell and set the current directory to the `./certificates` subfolder in the root of the Node.js project. Run the `Get-KeyCredentials.ps1` script to generate the `keyCredentials.txt` file. The `Get-KeyCredentials.ps1` script takes a single parameter, the name of your .CER file. If you used the recommended name, the Powershell command would be `.\Get-KeyCredentials.ps1 <full path to appcert.cer>`.

Note the output of the script will include a `Thumbprint` value. Copy this value to a safe place, we'll need it later!

Open `keyCredentials.txt` in a text editor. The contents should look similar to this (note the `value` key has been shortened for readability:
	
	"keyCredentials": [
	  {
	    "customKeyIdentifier": "BnZybKxv2c/hvkbHeM2bDy5Dcsg=",
	    "keyId": "bc8c2b35-d5e5-4732-ba0f-414189995db1",
	    "type": "AsymmetricX509Cert",
	    "usage": "Verify",
	    "value": "MIIDIzCC...dhpQcPR2"
	  }
	],

Once you have the `keyCredentials.txt` file, follow these steps to add it to the application's manifest.

1. In the Azure Management Portal, select your app's registration and click **Configure**. Click the **Manage Manifest** button and choose **Download manifest**. Download the manifest to the `./certificates` subfolder in the root of the Node.js project.
2. Open the manifest in a text editor. Locate the `keyCredentials` value (which is currently empty and replace it with the value from the `keyCredentials.txt` file. Save the manifest.
3. In the Azure Management Portal, click the **Manage Manifest** button and choose **Upload manifest**. Browse to the updated manifest and click **OK**.

If you want, go ahead and re-download the app's manifest to confirm that you successfully added the `keyCredentials` value. If you do, don't be alarmed that the `value` key under `keyCredentials` is **null**. Azure won't include the value when you download the manifest for security reasons!

## Configure the app ##

Open the `service-config.json` file and fill in the values as follows:

- `client_id`: The client ID from your app registration in the Azure Management Portal.
- `cert_file`: The relative path to the `private-key-rsa.pem` file.
- `cert_thumprint`: The `Thumbprint` value output by the `Get-KeyCredentials.ps1` script. If you did not use that script, this is the hexadecimal representation of the SHA1 fingerprint of the certificate.
- `tenant`: The domain name for your Office 365 tenant.

### Sample `service-config.json` file

	{
	  "client_id": "5f920d85-823a-4291-87bf-55ce1629186c",
	  "cert_file": "./certificates/private-key-rsa.pem",
	  "cert_thumbprint": "8745398DEA4982B394A13CAC39031FF394EE24F8",
	  "tenant": "contoso.onmicrosoft.com",
	  "users: [
		"allieb@contoso.onmicrosoft.com",
		"alexd@contoso.onmicrosoft.com"
	  ]
	}

Save your changes.

## Run the app ##

From the command line, do `node service.js`. You should see a line that starts with `TOKEN:`, followed by the JSON representation of the token response.