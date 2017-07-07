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

$arraySiteColumns = @()
$array =@()
$SPOLists="" 
$listCT=""
$allContentTypes=""
$xmlFilePath = "C:\Install\ListSchema.json"
#Create Export Files
New-Item $xmlFilePath -type file -force

$web=$Context.Web
$SPOLists=$web.Lists
$Context.Load($SPOLists)
$Context.ExecuteQuery()

##OOB SharePoint Lists to exclude from Export 
$ListsNameToExclude  = "_catalogs,Lists/ContentTypeSyncLog,/IWConvertedForms,FormServerTemplates,Lists/PublishedFeed,/ProjectPolicyItemList,/SiteAssets,/SitePages,/Style Library,Lists/TaxonomyHiddenList"   
$ArrayListtoExcl = $ListsNameToExclude.Split(",");  

foreach($list in $SPOLists) 
{  
  if($list.Hidden -ne $true ) 
  { 
    $item = New-Object PSObject
    Write-Host "List Name :"$list.Title
    $list = $web.Lists.GetByTitle($list.Title)
    $Context.Load($list)
    $LitsLookupField=$list.Fields
    $Context.Load($LitsLookupField)
    $Context.ExecuteQuery()
    $LitsInternalName=$list.EntityTypeName.replace('_x0020_',' ')
    if($list.BaseType -eq "GenericList")
    {
        $LitsInternalName=$LitsInternalName.Substring(0,$LitsInternalName.Length-4) 
    }
    
    $item | Add-Member -type NoteProperty -Name 'SListInternalName' -Value $($LitsInternalName)
    $item | Add-Member -type NoteProperty -Name 'SListType' -Value $($list.BaseType.ToString())
    $item | Add-Member -type NoteProperty -Name 'SListTitle' -Value $($list.Title)
    $item | Add-Member -type NoteProperty -Name 'SListDescription' -Value $($list.Description)
    $item | Add-Member -type NoteProperty -Name 'SListHidden' -Value $($list.Hidden)    
    $item | Add-Member -type NoteProperty -Name 'SListLookupField' -Value $($LitsLookupField.SchemaXml)
    if($list.ContentTypesEnabled -eq $true)
    {
       $defaultCT=""
       $allContentTypes=""
       $listCT=$list.ContentTypes
       $Context.Load($listCT)
       $Context.ExecuteQuery()
       $allContentTypes=""
       $defaultCT=$listCT[0].Name
       $contentTypeObj = @()
       foreach($scontentType in $listCT) 
       { 
             $listContentType=$scontentType.Fields
             $context.load($listContentType)
             $context.ExecuteQuery()
             $aryCTFields=@()
             $contentTypeObjFields=@()
             ForEach ($field in $listContentType)
              {
                $aryCTFields = @{ 'FieldName' = $field.Title
                             'FieldRequired' = $field.Required
                             'FieldHidden' = $field.Hidden
                             'FieldReadOnly' = $field.ReadOnlyField
                             'DisplayName' = $field.Title
                             'InternalName' = $field.InternalName
                            } 			
                $contentTypeObjFields+=$aryCTFields
                $aryCTFields=@()
              }
              $contentTypeObjFields= $contentTypeObjFields | ConvertTo-Json
              
            #$allContentTypes += $scontentType.Name +";"
            $array = @{ 'ContentTypeName' = $scontentType.Name
                         'ContentTypeFields' = $contentTypeObjFields

                        } 			
            $contentTypeObj+=$array
            $array=@()
       }
       $item | Add-Member -type NoteProperty -Name 'SListContentTypesEnabled' -Value $($list.ContentTypesEnabled)
       $item | Add-Member -type NoteProperty -Name 'SDefaultContentType' -Value $($defaultCT)
       
       $contentTypeObj = $contentTypeObj | ConvertTo-Json
       $item | Add-Member -type NoteProperty -Name 'SAllContentTypeNames' -Value $($contentTypeObj)
    }
    $arraySiteColumns += $item
  } 
  
 }

$bSitecolumns = @()

$bSitecolumns=$arraySiteColumns | ConvertTo-Json

Add-Content $xmlFilePath $bSitecolumns 