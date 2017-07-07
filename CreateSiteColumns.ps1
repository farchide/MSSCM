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

$site = $Context.Site 
$Context.Load($site)
$destWeb = $site.RootWeb
$Context.Load($destWeb)
$fields=$destWeb.Fields 
$Context.Load($fields)
$lists=$destWeb.Lists
$Context.Load($lists)
$Context.ExecuteQuery()

################################################################
# For SiteColumns
#Get exported XML file
$fieldsXML=(Get-Content -Raw 'C:\Install\SiteColumnsSchema.json' | ConvertFrom-Json)

#loop through each entry and create the columnGroup
For ($i=0; $i -lt $fieldsXML.Length; $i++) 
{
    $temp="false";
    $fieldXML = $fieldsXML[$i].Field.Replace("@{","")
    foreach ($Field in $fields)
    {
        if($fieldsXML[$i].Name -eq $Field.InternalName)
        {           
            $temp="true";
            break;
        }
    }
    if($temp -eq "true")
    {
       write-host "Yes, Given Site Column do Exists!"
    }
    else
    {
        if($fieldsXML[$i].Type -eq "Lookup")
        {  
            $ListExists="false"    
            foreach ($List in $lists)
            { 
                $ListInternalName=$fieldsXML[$i].ListInternalName.replace(' ', '_x0020_')
                if($List.BaseType -eq "GenericList")
                {
                    $dList=$List.EntityTypeName
                    $dList=$dList.Substring(0,$dList.Length-4) 
                }
                if($ListInternalName -eq $dList)
                {            
                    $ListExists="true";
                    break;
                }
            }
            if($ListExists -eq "true")
            {
               write-host -f green $listName "exists in the site"
               $spList = $destWeb.Lists.GetByTitle($fieldsXML[$i].ListTitle)
               $Context.Load($spList)
               $ListLookupFields=$spList.Fields
               $Context.Load($ListLookupFields)
               $spList.Update()
               $Context.ExecuteQuery()
            }
            else
            {
                write-host -f yellow $listName "List does not exist in the site"
                $spoListCreationInformation=New-Object Microsoft.SharePoint.Client.ListCreationInformation 
                $spoListCreationInformation.Title=$fieldsXML[$i].ListInternalName 
                #https://msdn.microsoft.com/EN-US/library/office/microsoft.sharepoint.client.listtemplatetype.aspx 
                $templateName=$fieldsXML[$i].ListBaseType
                $spoListCreationInformation.TemplateType=[int][Microsoft.SharePoint.Client.ListTemplatetype]::$templateName
                $spoList=$destWeb.Lists.Add($spoListCreationInformation) 
                $spoList.Update()
                $destWeb.Update()
                $Context.Load($spoList)
                $Context.ExecuteQuery() 
 
                Write-Host "Created List for Lookup"  -foregroundcolor Green 

                $spList = $Context.Web.Lists.GetByTitle($fieldsXML[$i].ListInternalName)
                $Context.Load($spList)
                $spList.Title=$fieldsXML[$i].ListTitle
                $Context.Load($spList)
                $ListLookupFields=$spList.Fields
                $Context.Load($ListLookupFields)
                $spList.Update()
                $Context.ExecuteQuery()                          
            }
            #Read read Json Array
            $templateXml = [XML]($fieldsXML[$i].ListLookupField) 
            #looping the List Field with Json Array
                ForEach($field in $templateXml.Fields.Field)  {	
                    if($field.Name -ne "ID" -or $field.Name -ne "Title" -and $field.Name -ne "")
                    {
                        $tempField="false";
                        foreach ($ListLookupField in $ListLookupFields)
                        {
                            if($ListLookupField.InternalName -eq $field.Name)
                            {            
                                $tempField="true";
                                break;
                            }
                        }
                        if($tempField -eq "true")
                        {
                           write-host "Yes, Given Site Column do Exists!"
                        }
                        else
                        {
                            if($field.Hidden -ne "TRUE")
                            {
		                        $spList.Fields.AddFieldAsXml($field.OuterXml,$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                                $spList.Update()
                                $Context.ExecuteQuery()
                            }
	                    }
                    }
                }
            #Create the Site Columns    
            $ListID=$spList.ID.Guid
            $ListID="{"+$ListID+"}"
            $webId=$destWeb.ID.Guid
            $webId=$webId
            $id="{"+[guid]::NewGuid()+"}"
            $fieldXML=$fieldXML.Replace($fieldsXML[$i].WebID,$destWeb.ID.Guid).Replace($fieldsXML[$i].List,$spList.ID.Guid)


            $fieldOption = [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue
            Write-Host "creating site columns"  -foregroundcolor black -backgroundcolor yellow 
            #Create column on the site
            if($fieldXML.Contains("Version")){
                $fieldXML=[XML]$fieldXML
                $fieldXML.Field.ParentNode.Field.RemoveAttribute("Version")
                $fieldXML=$fieldXML.Field.ParentNode.Field.OuterXml
            }            
            $field = $fields.AddFieldAsXml($fieldXML, $true, $fieldOption)
            $Context.Load($field)
            write-host "Created site column" $fieldsXML[$i].Name "on" $destWeb.Url
            $Context.ExecuteQuery()
	                 
        }
        else
        {  
        write-host $fieldXML     
        $fieldOption = [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue
        Write-Host "creating site columns"  -foregroundcolor black -backgroundcolor yellow 
        #Create column on the site
        if($fieldXML.Contains("Version")){
            $fieldXML=[XML]$fieldXML
            $fieldXML.Field.ParentNode.Field.RemoveAttribute("Version")
            $fieldXML=$fieldXML.Field.ParentNode.Field.OuterXml
        }        
        $field = $fields.AddFieldAsXml($fieldXML, $true, $fieldOption)
        $Context.Load($field)
        write-host "Created site column" $fieldsXML[$i].Name "on" $destWeb.Url
        }
    }   
}
try
{
  $Context.ExecuteQuery()
  Write-Host "Site columns created successfully" -foregroundcolor black -backgroundcolor Green 
}
catch
{
  Write-Host "Error while creating site columns" $itemField $_.Exception.Message -foregroundcolor black -backgroundcolor Red 
  return
}





