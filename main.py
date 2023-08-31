#!/usr/bin/env python3
import sys
import asyncio

DEBUG = False

global ad_users
ad_users = {}
error_encountered=0

AUA_SUBSCRIPTION="ae3575a6-395a-452f-87d8-7985bf239cf4"
AUA="A-NLWE-AUA-P1"
AUA_RG="A-NLWE-RGP-AUA-P1"

SECRETS={}

SECRET_KEYS= {
    "ad_to_sfmc_var_SFMC_ACCOUNT_ID",
    "ad_to_sfmc_var_SFMC_CLIENT_ID",
    "ad_to_sfmc_var_SFMC_CLIENT_SECRET",
    "ad_to_sfmc_var_AD_CLIENT_ID",
    "ad_to_sfmc_var_AD_CLIENT_SECRET"
}

SMFC_DATA_EXTENSION_CUSTOMER_KEY="58A8681E-615D-408C-B860-77B514BF048C"
SMFC_DATA_EXTENSION_NAME="NAE_Internal_Communications_DE"

SFMC_API_URL = "https://mc7mvw1rlkn1gvxwkvdk4zfs81h4.rest.marketingcloudapis.com/hub/v1/"
SFMC_SOAP_API_URL = "https://mc7mvw1rlkn1gvxwkvdk4zfs81h4.soap.marketingcloudapis.com/Service.asmx"
SFMC_SOAP_WSDL = " https://mc7mvw1rlkn1gvxwkvdk4zfs81h4.soap.marketingcloudapis.com/etframework.wsdl"
SFMC_API_AUTH_URL = "https://mcf5m6qnqkc44dvvmrzz6jxs74m8.auth.marketingcloudapis.com"


AD_TENANTS = {"naecentral","naeeuro","naecn","nordanglia","naena"}
NAE_CENTRAL_TENANT_AD_GROUP_IDS = {'NAE All Office Staff' : 'fc519dc2-f38d-4148-a7d3-1577f158906c'}


# To allow local debug, comment out the contents of this function and uncomment the print statement.
# and update:
# SECRETS={
#    "ad_to_sfmc_var_AD_CLIENT_ID" : "<SFMC-Automation AD App Registration/AUA Variable>",
#    "ad_to_sfmc_var_AD_CLIENT_SECRET" : "<SFMC-Automation AD App Registration/AUA Variable>",
#    "ad_to_sfmc_var_SFMC_ACCOUNT_ID" : "<On page 16 Pracedo Docs/AUA Variable>",
#    "ad_to_sfmc_var_SFMC_CLIENT_ID" : "<On page 16 Pracedo Docs>",
#    "ad_to_sfmc_var_SFMC_CLIENT_SECRET" : "<On page 16 Pracedo Docs>"
# }
def get_automation_variable():
    #print("\tSkipping authentication and using hard-coded values\n")
    try:
        client = AutomationClient(
            credential=DefaultAzureCredential(),
            subscription_id=AUA_SUBSCRIPTION,
        )
        for key in SECRET_KEYS:
            print("\tGetting variable '{}'".format(key))
            try:
                response = client.variable.get(
                    resource_group_name=AUA_RG,
                    automation_account_name=AUA,
                    variable_name=key,
                )
                value = response.value
                SECRETS[key] = value
                print("\Value variable begins: '{}'".format(value[:5]))
            except Exception as e:
                print("\t\tError getting variable: {}\n".format(e))
    except Exception as e:
        print("\t\tError authenticating: {}\n".format(e))


def __map_attr9_users(tenant, users):
    for user in users.value:
        if None != user.on_premises_extension_attributes:
            if None != user.on_premises_extension_attributes.extension_attribute_9:
                if user.on_premises_extension_attributes.extension_attribute_9.startswith("Administration") or user.extensions.extension_attribute_9.startswith("Teacher"):
                    if user.user_principal_name not in ad_users:
                        ad_users[user.user_principal_name] = {
                            "givenname": user.given_name,
                            "surname": user.surname,
                            "office" : user.office_location,
                            "department" : user.department,
                            "country" : user.country,
                            "street" : user.street_address,
                            "city" : user.city,         
                            "region" : tenant,
                            "company" : user.company_name
                        }
                    else:
                        print("\t\t\tDuplicate user found: {}".format(user.user_principal_name))
    return users

def __map_group_users(tenant, users):
    for user in users.value:

        if user.user_principal_name not in ad_users:
            ad_users[user.user_principal_name] = {
                "givenname": user.given_name,
                "surname": user.surname,
                "office" : user.office_location,
                "department" : user.department,
                "country" : user.country,
                "street" : user.street_address,
                "city" : user.city,         
                "region" : tenant,
                "company" : user.company_name
            }
        else:
            print("\t\t\tDuplicate user found: {}".format(user.user_principal_name))
    return users

async def get_ad_users():

    


    for tenant in AD_TENANTS:
    
        if "naecn" == tenant:
            tenant_id = "244cd12d-b25b-456f-b283-cfe4997f00f5"
            extension_attribute_9 = "extension_87ce92dfb96b4ac689f8f53836bffaa6_extensionAttribute9"
        elif "nordanglia" == tenant:
            tenant_id = "7fc9707b-9d8f-4f6a-856b-aeb848fef103"
            extension_attribute_9 = "extension_9afa365be4354169950f0b2b9c1d45fb_extensionAttribute9"
        elif "naeeuro" == tenant:
            tenant_id = "96d805bc-5a6b-4181-89b6-5eb1ca3ed4da"
            extension_attribute_9 = "extension_69c7512044b44dcebfeb06607100c4b0_extensionAttribute9"
        elif "naena" == tenant:
            tenant_id = "0e7c664a-d905-41dd-b08f-6ce92461eb0b"
            extension_attribute_9 = "extension_093768be288447f6a94d1dd4c5eef3c4_extensionAttribute9"
        elif "naecentral" == tenant:
            tenant_id = "189ae708-15d7-447e-9278-38b19d37390b"



        # Get Graph API client
        tenant_id = "common"
        credential = ClientSecretCredential("189ae708-15d7-447e-9278-38b19d37390b",
                                        SECRETS['ad_to_sfmc_var_AD_CLIENT_ID'],
                                        SECRETS['ad_to_sfmc_var_AD_CLIENT_SECRET'], additionally_allowed_tenants=[tenant_id])
        scopes = ['https://graph.microsoft.com/.default'] 
        client = GraphServiceClient(credentials=credential, scopes=scopes)
        auth_provider = AzureIdentityAuthenticationProvider(credential, scopes=scopes)


        # Query all tenants except NAE Central
        if "naecentral" != tenant:
            print("\tQuerying '{}' tenant users (extension attribute '9')".format(tenant))
            query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(select = ["user_principal_name","givenname","surname","office","department","country","street","city","region","company","onPremisesExtensionAttributes"],filter = "UserType eq 'Member'", top = 999)
            request_configuration = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(query_parameters = query_params,)
            request_adapter = GraphRequestAdapter(auth_provider)
            try:
                print("\t\tGetting initial response...")
                users = await client.users.get(request_configuration = request_configuration)
                
                # function to map users to a global ad_users variable
                __map_attr9_users(tenant, users)

                # Paginate through the remaining users
                count = 1
                while users.odata_next_link:
                        request_info = client.users.to_get_request_information(request_configuration=request_configuration)
                        request_info.url = users.odata_next_link
                        print("\t\tGetting paged response {}...".format(count))
                        users = await request_adapter.send_async(request_info, UserCollectionResponse, {})
                        # function to map users to a global ad_users variable
                        users = __map_attr9_users(tenant, users)
                        count = count + 1
                print("\t\tFinished.".format(count))
            except APIError as e:
                print("\tUnable to query AD due to error:\n\t\tMS Graph Client Request Id: {}\n\t\tMS Graph Request Id: {}\n\t\tHTTP Status Code: {}\n\t\tError Details: {}\n\t\tError Message: {}\n".format(e.error.inner_error.client_request_id, e.error.inner_error.request_id, e.response_status_code, e.error.code, e.error.message))
                error_encountered=1
            except Exception as e:
                print("\tUnable to query AD due to error: {}".format(e))
                error_encountered=1

        # For NAE Central tenant, query specific groups
        else:
            print("\tQuerying '{}' tenant users (members of specific groups)".format(tenant))
            for group_name in NAE_CENTRAL_TENANT_AD_GROUP_IDS:
                print("\t\tGroup: '{}'".format(group_name))
                try:

                    query_params = MembersRequestBuilder.MembersRequestBuilderGetQueryParameters(select = ["user_principal_name","givenname","surname","office","department","country","street","city","region","company"])
                    request_configuration = MembersRequestBuilder.MembersRequestBuilderGetRequestConfiguration(query_parameters = query_params,)
                    request_adapter = GraphRequestAdapter(auth_provider)
                    print("\t\tGetting initial response...")
                    members = await client.groups.by_group_id(NAE_CENTRAL_TENANT_AD_GROUP_IDS[group_name]).members.get(request_configuration = request_configuration)
                    
                    #next_link = members.odata_next_link
                    # while "" != next_link:

                    #client.groups.by_group_id(NAE_CENTRAL_TENANT_AD_GROUP_IDS[group_name]).get()
                    __map_group_users(tenant, members)

                    # Paginate through the remaining users
                    count = 1
                    while members.odata_next_link:
                            request_info = client.groups.to_get_request_information(request_configuration=request_configuration)
                            request_info.url = members.odata_next_link
                            print("\t\tGetting paged response {}...".format(count))
                            members = await request_adapter.send_async(request_info, GroupCollectionResponse, {})
                            # function to map users to a global ad_users variable
                            members = __map_group_users(tenant, members)
                            count = count + 1
                    print("\t\tFinished.".format(count))
                except Exception as e:
                    print("\tUnable to query AD due to error:\n\t\t{}\n".format(e))
                    error_encountered=1

    print("\tFound {} users in all tenants\n".format(len(ad_users)))
    return ad_users

def get_smfc_users(smfc_client):
    print ("\tGetting {} from SFMC API".format(SMFC_DATA_EXTENSION_NAME))
    row = ET_Client.ET_DataExtension_Row()
    row.CustomerKey = SMFC_DATA_EXTENSION_CUSTOMER_KEY
    row.auth_stub = smfc_client
    row.props = ["mail","givenname","surname","region", "company", "department", "country", "street", "city", "office"]
    getResponse = row.get()
    rows = getResponse.results
    while getResponse.more_results: 
        getResponse = row.getMoreResults()
        rows.append(getResponse.results)

    users = {}
    for row in rows:
        properties = row['Properties']
        for prop in properties:
            for keypair in prop[1]:
                if 'mail' == keypair['Name']:
                    mail = keypair['Value']
                elif 'givenname' == keypair['Name']:
                    givenname = keypair['Value']
                elif 'surname' == keypair['Name']:
                    surname = keypair['Value']
                elif 'region' == keypair['Name']:
                    region = keypair['Value']
                elif 'company' == keypair['Name']:
                    company = keypair['Value']
                elif 'office' == keypair['Name']:
                    office = keypair['Value']
                elif 'department' == keypair['Name']:
                    department = keypair['Value']
                elif 'country' == keypair['Name']:
                    country = keypair['Value']
                elif 'street' == keypair['Name']:
                    street = keypair['Value']
                elif 'city' == keypair['Name']:
                    city = keypair['Value']
        users[mail] = {
          "givenname" : givenname, 
          "surname" : surname,
          "department" : department,
          "country" : country,
          "street" : street,
          "city" : city,
          "office" : office,
          "region"  : region, 
          "company"  : company 
        }
    print ("\t{} contains {} users\n".format(SMFC_DATA_EXTENSION_NAME, len(users))) 
    return users



def add_smfc_users(smfc_client, users):
    row = ET_Client.ET_DataExtension_Row()
    row.auth_stub = smfc_client 
    row.CustomerKey = SMFC_DATA_EXTENSION_CUSTOMER_KEY
    row.Name = SMFC_DATA_EXTENSION_NAME
   
    for user in users:
        values = users[user]
        row.props = { "mail" : user, 
          "givenname" : values['givenname'], 
          "surname" : values['surname'],
          "department" : values['department'],
          "country" : values['country'],
          "street" : values['street'],
          "city" : values['city'],
          "office" : values['office'],
          "region"  : values['region'], 
          "company"  : values['company'] 
        }
        postResponse = row.post()
        print("\t\t{} for user {}".format(str(postResponse.results[0]['StatusMessage']), user))
    print("")

def remove_smfc_users(smfc_client, users):
    row = ET_Client.ET_DataExtension_Row()
    row.auth_stub = smfc_client
    row.CustomerKey = SMFC_DATA_EXTENSION_CUSTOMER_KEY
    row.Name = SMFC_DATA_EXTENSION_NAME

    for user in users:
        values = users[user]
        row.props = { "mail" : user, 
          "givenname" : values['givenname'], 
          "surname" : values['surname'],
          "department" : values['department'],
          "country" : values['country'],
          "street" : values['street'],
          "city" : values['city'],
          "office" : values['office'],
          "region"  : values['region'], 
          "company"  : values['company'] 
        }
        deleteResponse = row.delete()
        print("\t\t{} for user {}".format(str(deleteResponse.results[0]['StatusMessage']), user))
    print("")

def get_sfmc_users_to_remove(sfmc_users, ad_users):
    keys = sfmc_users.keys() - ad_users.keys()
    return {key: sfmc_users[key] for key in keys}

def get_sfmc_users_to_add(sfmc_users, ad_users):
    keys = ad_users.keys() - sfmc_users.keys()
    return {key: ad_users[key] for key in keys}

# Main entry point for script
if __name__ == '__main__':

    try:
        import ET_Client
        from azure.identity import DefaultAzureCredential
        from azure.identity.aio import ClientSecretCredential
        from azure.mgmt.automation import AutomationClient
        from msgraph import GraphServiceClient
        from msgraph import GraphRequestAdapter
        from msgraph.generated.models.user_collection_response import UserCollectionResponse
        from msgraph.generated.models.group_collection_response import GroupCollectionResponse
        from msgraph.generated.users.users_request_builder import UsersRequestBuilder
        from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder
        from msgraph.generated.groups.item.members.members_request_builder import MembersRequestBuilder
        from kiota_abstractions.api_error import APIError
        from kiota_authentication_azure.azure_identity_authentication_provider import AzureIdentityAuthenticationProvider

    except ImportError as e:
        print("Error: {}\nPlease run \'pip install -r requirements.txt\' to install required Python modules".format(e))
        sys.exit(1)

    print("Getting variables from automation account...\n")
    get_automation_variable()


    print("Getting all NAE staff email addresses from AD...\n")
    # Get Azure Automation Run-As Credentials
    ad_users = asyncio.run(get_ad_users())

    print("AD query done\n")
    print("Updating SFMC list based on AD users...\n")
    
    
    # Authenticate with Salesforce Marketing Cloud    
    sfmc_params = { 'clientid': SECRETS['ad_to_sfmc_var_SFMC_CLIENT_ID'],
            'clientsecret': SECRETS['ad_to_sfmc_var_SFMC_CLIENT_SECRET'],
            'authenticationurl': SFMC_API_AUTH_URL,
            'soapendpoint': SFMC_SOAP_API_URL,
            'defaultwsdl': SFMC_SOAP_WSDL,
            'baseapiurl': SFMC_API_URL,
            'applicationType': 'server',
            'accountId': SECRETS['ad_to_sfmc_var_SFMC_ACCOUNT_ID'],
            'useOAuth2Authentication': 'True'}
    sfmc_client = ET_Client.ET_Client(False, DEBUG, sfmc_params)

    # Retrieve data extension that contains users
    sfmc_users = get_smfc_users(sfmc_client)

    # Compare AD vs SFMC users to work out who to add and remove
    users_to_remove = get_sfmc_users_to_remove(sfmc_users, ad_users)
    users_to_add = get_sfmc_users_to_add(sfmc_users, ad_users)


    # Call SFMC API to add and remove users
    if len(users_to_remove) > 0:
        print("\t{} disabled/removed AD users to remove from {}:".format(len(users_to_remove), SMFC_DATA_EXTENSION_NAME))
        remove_smfc_users(sfmc_client, users_to_remove)
    else:
        print("\tNo disabled/removed AD users to remove from {}\n".format(SMFC_DATA_EXTENSION_NAME))

    if len(users_to_add) > 0:
        print("\t{} enabled/new AD users to add to {}:".format(len(users_to_add), SMFC_DATA_EXTENSION_NAME))
        add_smfc_users(sfmc_client, users_to_add)
    else:
        print("\tNo enabled/new AD users to add to {}\n".format(SMFC_DATA_EXTENSION_NAME))   
    
    print("SFMC update done")
    sys.exit(error_encountered)