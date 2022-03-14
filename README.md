# pnpjs-starter

A fully implemented starter that will allow you to use the [PnP JS Library](https://pnp.github.io/pnpjs/) to access SharePoint Online from a Nodejs command line application.

## Build

To build the console app

From the root of the project.

```sh
npm install
```

## Configure

The app uses an Azure Active Directory App Registration with a certificate that has been granted access via the Graph API.

Log onto the Azure Tenant you will be using with a shell that supports the Azure CLI.

The default Codespace image has it available already but if not use the [Azure Cloud Shell](https://shell.azure.com).

```sh
# If you need to explicitly authenticate, e.g. using GitHub CodeSpaces
# --use-device-code allows you to log on from GitHub Codespaces without port forwarding issues
az login --use-device-code 

echo "Setting the name of the subscription"
subscriptionName="Visual Studio Enterprise"

echo "Setting the subscription"
az account set --subscription="$subscriptionName"

echo "Getting the current identity's object id"
identityObjectId=$(az ad user show --id $(az account show --query "user.name" --output tsv) \
    --query "objectId" \
    --output tsv)

echo "Settng the location to use"
location=uksouth

echo "Creating a resource group"
az group create \
    --name rg-pnpjs-starter \
    --location $location

echo "Creating a key vault"
az keyvault create \
    --name kv-pnpjs-starter \
    --resource-group rg-pnpjs-starter \
    --enable-rbac-authorization

echo "Assigning access to the key vault"
az role assignment create \
    --role "Key Vault Administrator" \
    --assignee $(az account show \
        --query "user.name" \
        --output tsv) \
    --scope $(az keyvault show \
        --name kv-pnpjs-starter \
        --query "id" \
        --output tsv) 

echo "Creating a certificate"
az keyvault certificate create \
    --name cert-pnpjs-starter \
    --vault-name kv-pnpjs-starter \
    --policy "$(az keyvault certificate get-default-policy)"

echo "Getting if Azure AD App Registration exists"
appPnPJSStarterAppId=$(az ad app list --query "[?displayName=='pnpjs-starter'].appId" --output tsv)
if [[ -z $appPnPJSStarterAppId ]]; then
    echo "Creating the Azure AD App Registration"
    az ad app create --display-name 'pnpjs-starter'
    appPnPJSStarterAppId=$(az ad app list --query "[?displayName=='pnpjs-starter'].appId" --output tsv)
fi

echo "Getting if the Azure AD App Registration's Service Principal exists"
appPnPJSStarterSPLength=$(az ad sp list --spn $appPnPJSStarterAppId | jq ". | length")
if [[ $appPnPJSStarterSPLength -eq "0" ]]; then
    echo "Creating the Azure AD App Registration's Service Principal"
    az ad sp create --id $appPnPJSStarterAppId
fi

echo "Adding a certificate for the Azure AD App Registration"
az ad app update \
    --id $appPnPJSStarterAppId \
    --key-type "AsymmetricX509Cert" \
    --key-usage "Verify" \
    --key-value $(az keyvault certificate show \
        --name cert-pnpjs-starter \
        --vault-name kv-pnpjs-starter \
        --query "cer" \
        --output tsv)

echo "Getting the Microsoft Graph Service Principal"
msGraphSPAppId=$(az ad sp list \
    --query "[?appDisplayName=='Microsoft Graph'].appId" \
    --all \
    --output tsv)

echo "Getting the Microsoft Graph Sites.Selected Role"
msGraphSPSitesSelectedAppRoleId=$(az ad sp show \
    --id $msGraphSPAppId \
    --query "appRoles[?value=='Sites.Selected'].id" \
    --output tsv)

echo "Getting if the Microsoft Graph Sites.Selected Role has the Azure AD App Registration"
appPnPJSStarterAppMSGraphSitesSelectedLength=$(az ad app permission list \
    --id $appPnPJSStarterAppId \
    --query "[?resourceAppId=='$msGraphSPAppId'].resourceAccess | [].{id:id} | [?id=='$msGraphSPSitesSelectedAppRoleId']" |
    jq ". | length")
if [[ $appPnPJSStarterAppMSGraphSitesSelectedLength -eq "0" ]]; then
    echo "Adding the Microsoft Graph Sites.Selected Role to the Azure AD App Registration"
    az ad app permission add --id $appPnPJSStarterAppId --api $msGraphSPAppId --api-permissions $msGraphSPSitesSelectedAppRoleId=Role
fi

echo "Getting an access token for the current session"
accessToken=$(az account get-access-token --resource-type ms-graph --query "accessToken" --output tsv)

echo "Getting the Azure AD App Registraion's Service Principal Object Id"
appPnPJSStarterSPObjectId=$(az ad sp show --id $appPnPJSStarterAppId --query "objectId" --output tsv)

echo "Getting the Microsoft Graph's Service Principal Object Id"
msGraphSPObjectId=$(az ad sp list --query "[?appDisplayName=='Microsoft Graph'].objectId" --all --output tsv)

echo "Getting if the Azure AD Application's Microsoft Graph Sites.Selected Role has been granted"
appPnPJSStarterSPObjectIdappRoleAssignmentsLength=$(curl \
    --header "Authorization: Bearer $accessToken" \
    --request GET \
    --url https://graph.microsoft.com/v1.0/servicePrincipals/$appPnPJSStarterSPObjectId/appRoleAssignments |
    jq '.value | length')

if [[ $appPnPJSStarterSPObjectIdappRoleAssignmentsLength -eq 0 ]]; then
    echo "Granting the Azure AD Application's Microsoft Graph Sites.Selected Role"
    curl \
        --header "Authorization: Bearer $accessToken" \
        --header "Content-Type: application/json" \
        --request POST \
        --url https://graph.microsoft.com/v1.0/servicePrincipals/$appPnPJSStarterSPObjectId/appRoleAssignments \
        --data "{\"principalId\":\"$appPnPJSStarterSPObjectId\", \"resourceId\":\"$msGraphSPObjectId\", \"appRoleId\":\"$msGraphSPSitesSelectedAppRoleId\"}"
fi

echo "Getting the certificate as a secret"
derBase64=$(az keyvault secret show \
        --name cert-pnpjs-starter \
        --vault-name kv-pnpjs-starter \
        --query "value" \
        --output tsv)

mkdir temp
cd temp
openssl req -x509 -newkey rsa:2048 -keyout keytmp.pem -out cert.pem -days 365 -passout pass:HereIsMySuperPass -subj '/C=US/ST=Washington/L=Seattle'
openssl rsa -in keytmp.pem -out key.pem -passin pass:HereIsMySuperPass

az ad app credential reset --id d8f467b5-1067-4cc8-ba19-c10414087520 --cert "@temp/cert.pem"
```
