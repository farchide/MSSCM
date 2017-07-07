Clear

#Add references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
#Getting the current path of CSOM dll
$currentpath = (Get-Item -Path ".\" -Verbose).FullName+'\'

#Add references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Write-Host "Load CSOM libraries" -foregroundcolor black -backgroundcolor yellow

Set-Location $PSScriptRoot
$dll1 = "$currentpath" + "Microsoft.SharePoint.Client.dll"
$dll2 = "$currentpath" + "Microsoft.SharePoint.Client.Runtime.dll"
$dll3 = "$currentpath" + "Microsoft.SharePoint.Client.Taxonomy.dll"
Add-Type -Path $dll1
Add-Type -Path $dll2
Add-Type -Path $dll3
Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 

#Credentials to connect to office 365 site collection url 
$url =Read-Host -Prompt "Please enter site collection url"
$username =Read-Host -Prompt "Please enter User Name"
$Password = Read-Host -Prompt "Please enter your password" -AsSecureString

## The following four lines only need to be declared once in your script.
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Description."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Description."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
## Use the following each time your want to prompt the use
$title = "Title" 
$message = "Is it for SharePoint Online?"
$result = $host.ui.PromptForChoice($title, $message, $options, 1)
switch ($result) {
    0{
        $result= "Yes"
    }
    1{
        $result= "No"
    }
}

#Loading the Site context
Write-Host "Authenticate to SharePoint Online site collection $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow  
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
if($result -eq "Yes"){
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
$Context.Credentials = $credentials 
}
else{
    $credentials = New-Object System.Net.NetworkCredential($username, $password)
    $Context.Credentials = $credentials 
}

#Local Variables
$arraySiteColumns = @()
$bSitecolumns = @()
$ListLookupField=$null
$fields=$null
$id=$null
 
#OutPut Files to store the Site Columsn and Content Types in the form of JSON
$xmlFilePath1 = "C:\Install\SiteColumnsSchema.json"
$xmlFilePath = "C:\Install\ContentTypeSchema.json"
New-Item $xmlFilePath -type file -force

#Accessing the web  and all Fields in the Web Objects 
$web = $Context.Web 
$Context.Load($web)
$fields=$web.Fields 
$Context.Load($fields)
$Context.ExecuteQuery()

echo "Generating File for Site Columns..." 
ForEach ($id in $fields) 
{ 
    $item = New-Object PSObject
    $item | Add-Member -type NoteProperty -Name 'Field' -Value $($id.SchemaXml)
    $item | Add-Member -type NoteProperty -Name 'Name' -Value $($id.InternalName)
    $item | Add-Member -type NoteProperty -Name 'Title' -Value $($id.Title)
    $item | Add-Member -type NoteProperty -Name 'Type' -Value $($id.FieldTypeKind.ToString())
    #getting the current site column Field Type
    $ifLookup=$id.FieldTypeKind.ToString();
    #getting the if current site column consists of Lookup or not
    if($ifLookup -eq "Lookup")
    {
        #checking condition Lookup list is empty and length of guid is 38
        if($id.LookupList -ne "" -and $id.LookupList.Length -ge '38')
        {
            # Get the lookup List
            $lookupList=$id.LookupList
            $lookupListFieldID=$id.LookupField 
            $IsMultipleValues=$id.AllowMultipleValues      
            $item | Add-Member -type NoteProperty -Name 'List' -Value $($lookupList)            
            $listID=$lookupList.Replace("{","").Replace("}","")
            $list = $web.Lists.GetById("$listID")
            $Context.Load($list)
            $Context.ExecuteQuery()

            $ListInternalName=$list.EntityTypeName.replace('_x0020_', ' ')
            if($list.BaseType -eq "GenericList")
            {
                $ListInternalName=$ListInternalName.Substring(0,$ListInternalName.Length-4) 
            }
            #Add lookup list details to the PS Object of source field details
            $item | Add-Member -type NoteProperty -Name 'ListInternalName' -Value $($ListInternalName)
            $item | Add-Member -type NoteProperty -Name 'ListBaseType' -Value $($list.BaseType.ToString())
            $item | Add-Member -type NoteProperty -Name 'ListTitle' -Value $($list.Title)
            $item | Add-Member -type NoteProperty -Name 'IsMultipleValues' -Value $($IsMultipleValues)
            $item | Add-Member -type NoteProperty -Name 'ListField' -Value $($lookupListFieldID)            
            $item | Add-Member -type NoteProperty -Name 'Group' -Value $($id.Group)
            $item | Add-Member -type NoteProperty -Name 'Description' -Value $($id.Description)
            $item | Add-Member -type NoteProperty -Name 'ReadOnlyField' -Value $($id.ReadOnlyField)
            $item | Add-Member -type NoteProperty -Name 'Hidden' -Value $($id.Hidden)
            $item | Add-Member -type NoteProperty -Name 'Required' -Value $($id.Required)
            $item | Add-Member -type NoteProperty -Name 'EnforceUniqueValues' -Value $($id.EnforceUniqueValues)
            $item | Add-Member -type NoteProperty -Name 'WebID' -Value $($web.ID)
            $ListLookupField=$list.Fields
            $Context.Load($ListLookupField)
            $Context.ExecuteQuery()
            $item | Add-Member -type NoteProperty -Name 'ListLookupField' -Value $($ListLookupField.SchemaXml)
            #write-host "Created site column for Lookup" $id.InternalName "and" $lookupList "and " $lookupListFieldID "and"  $IsMultipleValues "and"  $ListInternalName 
         }
    }
    $arraySiteColumns += $item
} 

$bSitecolumns=$arraySiteColumns | ConvertTo-Json
Add-Content $xmlFilePath1 $bSitecolumns 

#Operation+++++++++++++++++++++++++++++++++++++++++++++++++++Start 
echo "CSV file generated successfully for the Site Columns, please check the below given path" 
echo "File created at :" $xmlFilePath1

########################################################################
## Generate Content Type Schema Details from Source ##
########################################################################
$ContentTypeJson= @()
$arrayContentType = @()
$array =@()
$cc=""; 
$field=$null

$cts = $web.ContentTypes
$Context.Load($cts)
$context.ExecuteQuery()

New-Item $xmlFilePath -type file -force

# Looping through source content types
ForEach ($cc in $Context.Web.ContentTypes)
{
        # Creating PS Object for each content type from source
        $item = New-Object PSObject
        $item | Add-Member -type NoteProperty -Name 'Name' -Value $($cc.Name)
        $item | Add-Member -type NoteProperty -Name 'ID' -Value $($cc.Id.ToString())
        $item | Add-Member -type NoteProperty -Name 'Description' -Value $($cc.Description)
        $item | Add-Member -type NoteProperty -Name 'Group' -Value $($cc.Group)
        $item1 = New-Object PSObject
		$contentTypeObj = @()
        $context.load($cc.fields)
        $context.ExecuteQuery()
        # Looping through Content Type columns and adding to content type schema details (for Json)
          ForEach ($field in $cc.Fields)
          {
            $array = @{ 'FieldName' = $field.Title
                         'FieldRequired' = $field.Required
                         'FieldHidden' = $field.Hidden
                         'FieldReadOnly' = $field.ReadOnlyField
                         'DisplayName' = $field.Title
                         'InternalName' = $field.InternalName
                        } 			
            $contentTypeObj+=$array
            $array=@()
          }
          $contentTypeObj = $contentTypeObj | ConvertTo-Json
          $item | Add-Member -type NoteProperty -Name 'Fields' -Value $($contentTypeObj) 
          $ContentTypeJson += $item
}
$arrayContentType = $ContentTypeJson | ConvertTo-Json
Add-Content $xmlFilePath $arrayContentType 
#Operation+++++++++++++++++++++++++++++++++++++++++++++++++++Start 
echo "CSV file generated successfully for the Content Types, please check the below given path" 
echo "File created at :" $xmlFilePath


