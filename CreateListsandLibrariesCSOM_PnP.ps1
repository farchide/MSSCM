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

$SPOLists=""
$dCTS=$null
$ctFields=""
$ccFields=$null
$contentTypeFields=$null

$destWeb=$Context.Web
$Context.Load($destWeb)
$SPOLists=$destWeb.Lists
$dCTS=$Context.Web.ContentTypes
$fields = $destWeb.Fields
$Context.Load($fields)
$Context.Load($SPOLists)
$Context.Load($dCTS)
$Context.ExecuteQuery()
foreach ($List in $SPOLists)
{ 
    Write-Host "Lists:" $List.EntityTypeName
}
#Get exported XML file
$fieldsXML=(Get-Content -Raw 'C:\Install\ListSchema.json' | ConvertFrom-Json)


#loop through each entry and create the columnGroup
For ($i=0; $i -lt $fieldsXML.Length; $i++) 
{
    $temp="false";
    $fieldXML = $fieldsXML[$i]
    $ListExists="false"    
    #$spList=$destWeb.Lists.TryGetList($fieldsXML[$i].LitsTitle.replace(' ', '%20'
    foreach ($List in $SPOLists)
    { 
        $LitsInternalName=$fieldsXML[$i].SListInternalName.replace(' ', '_x0020_')        
        Write-Host "Source List name:" $LitsInternalName 
        if($List.BaseType -eq "GenericList")
        {
            $DLitsInternalName=$List.EntityTypeName.Substring(0,$List.EntityTypeName.Length-4) 
        } 
        else
        {
            $DLitsInternalName=$List.EntityTypeName
        }
        Write-Host "Destination List name:" $DLitsInternalName      
        if($LitsInternalName -eq $DLitsInternalName)
        {            
            $ListExists="true";
            break;
        }
    }
    if($ListExists -eq "true")
    {
        write-host -f green $listName "exists in the site"
        $spList = $destWeb.Lists.GetByTitle($fieldsXML[$i].SListTitle)
        $Context.Load($spList)
        $Context.ExecuteQuery()
    }
    else
    {
        write-host -f yellow $fieldsXML[$i].SListTitle ":does not exist in the site"
        $spoListCreationInformation=New-Object Microsoft.SharePoint.Client.ListCreationInformation 
        $spoListCreationInformation.Title=$fieldsXML[$i].SListInternalName
        $spoListCreationInformation.Description=$fieldsXML[$i].SListDescription 
        #https://msdn.microsoft.com/EN-US/library/office/microsoft.sharepoint.client.listtemplatetype.aspx 
        $templateName=$fieldsXML[$i].SListType
        $spoListCreationInformation.TemplateType=[int][Microsoft.SharePoint.Client.ListTemplatetype]::$templateName
        $spoList=$destWeb.Lists.Add($spoListCreationInformation) 
        $spoList.Update()
        $destWeb.Update()
        #$spoList.Description=$sListDescription 
        $Context.Load($spoList)
        $Context.ExecuteQuery() 
 
        Write-Host "----------------------------------------------------------------------------"  -foregroundcolor Green 
        Write-Host "List" $fieldsXML[$i].SListTitle "created in "$destWeb" !!" -ForegroundColor Green 
        Write-Host "----------------------------------------------------------------------------"  -foregroundcolor Green  
                 
        #Get the Lists
        write-host -f green $fieldsXML[$i].SListTitle "exists in the site"
        $spList = $Context.Web.Lists.GetByTitle($fieldsXML[$i].SListInternalName)
        $Context.Load($spList)
        #$Context.ExecuteQuery()
        $spList.Title=$fieldsXML[$i].SListTitle
        #$spList.Title="Desc List"
        $Context.Load($spList)
        $spList.Update()
        $Context.ExecuteQuery()                          
    }
    if($spList.Hidden -ne $true)
    {
        if($fieldsXML[$i].SListContentTypesEnabled -eq $true)
        {

            if($spList.ContentTypesEnabled -eq $false){   
                $spList.ContentTypesEnabled = $true
            }
            $spList.update()
            $Context.ExecuteQuery()
            $customContentTypeArray = $fieldsXML[$i].SAllContentTypeNames | ConvertFrom-Json
            for($k=0; $k -lt $customContentTypeArray.Length; $k++)
            {
                $ctExists="false"
                $ctFields=""                
                $ccFields=$spList.ContentTypes
                $Context.Load($ccFields)
                $Context.ExecuteQuery()
               foreach($cc in $ccFields) 
                { 
                    if($customContentTypeArray[$k].ContentTypeName -eq $cc.Name) 
                    {
                        $ctExists="true";
                        $ctFields=$cc
                        break;
                    }
                }
                if ($ctExists -eq "true"){
                    Write-Host "--- ContentType exists in the List"
                    $newContentTypes=$spList.ContentTypes
                    $Context.Load($newContentTypes)
                    $Context.ExecuteQuery()

                    $objCT=$newContentTypes | Where {$_.Name -eq $customContentTypeArray[$k].ContentTypeName}
                    $SPContentType=$spList.ContentTypes.GetById($objCT.Id.StringValue) 
                    $Context.Load($SPContentType)
                     
                    $contentTypeFields=$spList.ContentTypes.GetById($objCT.Id.StringValue).Fields 
                    $Context.Load($contentTypeFields)                  
                    $Context.ExecuteQuery()
                } 
                else
                {
                    Write-Host "--- ContentType does not exists in the List"
                    Write-Host "--- adding Content type  in the List"
                    $objCT=$context.Web.ContentTypes | Where {$_.Name -eq $customContentTypeArray[$k].ContentTypeName}
                    $addExistContentType=$spList.ContentTypes.AddExistingContentType($objCT)
                    $spList.update()
                    $Context.ExecuteQuery()
                    $newContentTypes=$spList.ContentTypes
                    $Context.Load($newContentTypes)
                    $Context.ExecuteQuery()

                    $objCT=$newContentTypes | Where {$_.Name -eq $customContentTypeArray[$k].ContentTypeName}
                    $SPContentType=$spList.ContentTypes.GetById($objCT.Id.StringValue) 
                    $Context.Load($SPContentType)
                    $contentTypeFields=$spList.ContentTypes.GetById($objCT.Id.StringValue).Fields 
                    $Context.Load($contentTypeFields)                                     
                    $Context.ExecuteQuery()
                }  
               <##Field to be verified for each content type                
                $contentTypeFields=$newContentTypes.Fields
                $Context.Load($contentTypeFields)  
                $Context.ExecuteQuery()
                <## Fields to be checked starts here ##>
                $ctfieldsXML=($customContentTypeArray[$k].ContentTypeFields | ConvertFrom-Json)
                for($j=0; $j -lt $ctfieldsXML.Length; $j++) {
                    $ctFieldIsExist="false"                     
                    foreach($ctf in $contentTypeFields) {
                       if($ctfieldsXML[$j].InternalName -eq $ctf.InternalName) 
                       {
                          $ctFieldIsExist="true";
                           break;
                       } 
                   } 
                    if ($ctFieldIsExist -eq "true"){
                        Write-Host "--- ContentType Field [$($ctfieldsXML[$j].InternalName)] -> Match [$ctFieldIsExist]" -ForegroundColor Green
                    }
                    else
                    {
                        Write-Host "--- ContentType Field [$($ctfieldsXML[$j].InternalName)] -> Match [$ctFieldIsExist]" -ForegroundColor White
                        #Add Fields Reference to the New content type
                        if(!$SPContentType.FieldLinks[$ctfieldsXML[$j].InternalName])
                        {
                            $field = $fields.GetByTitle($ctfieldsXML[$j].DisplayName) 
                            Write-Host "Site column" $Name "exist" -foregroundcolor black -backgroundcolor Green

                            #Create a field link for the Content Type by getting an existing column
                            $flci = new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
           
                            $flci.Field = $field

                            #Check to see if column is Optional, Required or Hidden
                            if($flci.Field.InternalName -ne "Title")
                             {
                                #Add column to Content Type
                                $SPContentType.FieldLinks.Add($flci)
                                $SPContentType.Update($false)
                                $Context.Load($SPContentType)
                                $Context.ExecuteQuery()
                                <##$SPContentType=$Context.web.contenttypes.getbyid($ctci.Id.StringValue)                        
                                $Context.Load($SPContentType)
                                $Context.ExecuteQuery()#>
                                $Field=$SPContentType.Fields.GetByTitle($ctfieldsXML[$j].DisplayName)
                                $Context.Load($Field)
                                $Context.ExecuteQuery()
                                If($fieldsXML[$j].FieldRequired -eq "True") 
                                { 
                                    $SPContentType.FieldLinks.GetById($Field.Id.Guid).Required = $true;
                                } 
                                If($fieldsXML[$j].FieldHidden -eq "True")
                                { 
                                    $SPContentType.FieldLinks.GetById($Field.Id.Guid).Hidden = $true;
                                }
                                $SPContentType.Update($false)
                                $Context.Load($SPContentType)
                                $Context.ExecuteQuery()
                                  
                            }
                        }  
                    }                            
                } 
            <## Fields to be Checked ends here ##>
                                                
            }  

            ##Set the Default content Type in the Lists
            $newContents=$spList.ContentTypes
            $Context.Load($newContents)
            $Context.ExecuteQuery()
            $newCTO=@()
            foreach($cc in $newContents) {
                if($fieldsXML[$i].SDefaultContentType -eq $cc.Name) 
                {
                    $ctList = New-Object System.Collections.Generic.List[Microsoft.SharePoint.Client.ContentTypeId]
                    $ctList.Add($cc.Id)
                    $spList.RootFolder.UniqueContentTypeOrder=$ctList
                    $spList.Update()
                    $Context.Load($spList)
                    $Context.ExecuteQuery() 
                    break;
                }
            }                  

        }
        #Read read Json Array and create the List custom fields only
        $templateXml = [XML]($fieldsXML[$i].SListLookupField) 
        $LitsLookupFields=$spList.Fields
        $Context.Load($LitsLookupFields)
        $spList.Update()
        $Context.ExecuteQuery()
        #looping the List Field with Json Array
        ForEach($field in $templateXml.Fields.Field)  {
            if($spList.Title -eq $LitsInternalName)
            {
                if($field.Name -ne "ID" -or $field.Name -ne "Title" -and $field.Name -ne "" -and $field.Name -ne "Checkmark")
                {
                    $tempField="false";               
                    foreach ($LitsLookupField in $LitsLookupFields)
                    {
                        if($LitsLookupFields -ne $true)
                        {
                            if($LitsLookupField.InternalName -eq $field.Name)
                            {            
                                $tempField="true";
                                break;
                            }
                        }
                    }
               
                    if($tempField -eq "true")
                    {
                        write-host "Yes, Given Column do Exists in the List:" $fieldsXML[$i].SListTitle
                    }
                    else
                    {
                        if($field.Hidden -ne "TRUE")
                        {
                            Write-Host "List Name:" $fieldsXML[$i].SListTitle  "and  Field Name:" $field.Name
		                    $spList.Fields.AddFieldAsXml($field.OuterXml,$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                            $spList.Update()
                            $Context.ExecuteQuery()
                        }
	                }
                }
            }
        }
        
        
    }

}