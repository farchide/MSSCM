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

#Local variables
$fieldCollection =$null
$site=$null
$rootWeb=$null
$fields=$null
$webCT=$null
$contentTypes=$null
$ctFields=""

$site = $Context.Site 
$Context.Load($site)
$rootWeb = $site.RootWeb
$Context.Load($rootWeb)
$fields=$rootWeb.Fields
$Context.Load($fields)
$webCT = $Context.Site.RootWeb.ContentTypes
$Context.Load($webCT)
$contentTypes = $Context.Web.ContentTypes
$Context.Load($contentTypes)
$Context.ExecuteQuery()

#Get exported XML file
$CTXML=(Get-Content -Raw 'C:\Install\ContentTypeSchema.json' | ConvertFrom-Json)
for($i=0; $i -lt $CTXML.Length; $i++) {
    $ccIsExist="false";
    foreach( $cc in $Context.Web.ContentTypes) {
        if($CTXML[$i].Name -eq $cc.Name) 
        {
            $ctFields= $cc.Fields
            $Context.Load($ctFields)
            $Context.ExecuteQuery()
            $ccIsExist="true";
            break;
        }                
    }
    if ($ccIsExist -eq "true"){
        Write-Host "--- ContentType [$($CTXML[$i].Name)] -> Match [$ccIsExist]" -ForegroundColor Green 
        $fieldsXML=($CTXML[$i].Fields | ConvertFrom-Json)
        for($j=0; $j -lt $fieldsXML.Length; $j++) {
            $ctFieldIsExist="false"
            foreach($ctf in $ctFields) {
                if($fieldsXML[$j].InternalName -eq $ctf.InternalName) 
                {
                    $ctFieldIsExist="true";
                    break;
                } 
            } 
            if ($ctFieldIsExist -eq "true"){
                Write-Host "--- ContentType Field [$($fieldsXML[$j].InternalName)] -> Match [$ctFieldIsExist]" -ForegroundColor Green
            }
            else
            {
                Write-Host "--- ContentType Field [$($fieldsXML[$j].InternalName)] -> Match [$ctFieldIsExist]" -ForegroundColor White
                $SPContentType = $Context.web.contenttypes.getbyid($cc.Id.StringValue)
                $Context.Load($SPContentType) 
                $Context.ExecuteQuery()
                #Add Fields Reference to the New content type
                if( !$SPContentType.FieldLinks[$fieldsXML[$j].InternalName])
                {
                    $field = $Context.Site.RootWeb.Fields.GetByInternalNameOrTitle($fieldsXML[$j].InternalName)
                    $Context.Load($field) 
                    $Context.ExecuteQuery()
                    #Create a field link for the Content Type by getting an existing column
                    $flci = new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
           
                    $flci.Field = $field
                    #Check to see if column is Optional, Required or Hidden
                    if($flci.Field.InternalName -ne "Title")
                    {
                        #Add column to Content Type
                        $SPContentType.FieldLinks.Add($flci)
                        $SPContentType.Update($true)
                        $Context.Load($SPContentType)
                        $Context.ExecuteQuery()
                        $SPContentType=$Context.web.contenttypes.getbyid($cc.Id.StringValue)                        
                        $Context.Load($SPContentType)
                        $Context.ExecuteQuery()
                        $Field=$SPContentType.Fields.GetByInternalNameOrTitle($fieldsXML[$j].InternalName)
                        $Context.Load($Field)                        
                        $Context.ExecuteQuery()                       
                        If($fieldsXML[$j].FieldRequired -eq "True") 
                        { 
                           $SPContentType.FieldLinks.GetById($Field.Id.Guid).Required = $true;
                        } 
                        If($fieldsXML[$j].FieldHidden -eq "True")
                        { 
                            $SPContentType.FieldLinks.GetById($Field.Id.Guid).Hidden=$true;
                        }   
                        $SPContentType.Update($true)
                        $Context.Load($SPContentType)
                        $Context.ExecuteQuery()
                    }                   
                    
                }  
               
            }
        }                                                    
    }
    else{
        Write-Host "--- ContentType [$($CTXML[$i].Name)] -> Match [$ccIsExist]" -ForegroundColor White
        # create Content Type using ContentTypeCreationInformation object (ctci)
    $ctci = new-object Microsoft.SharePoint.Client.ContentTypeCreationInformation
    $ctci.name = $CTXML[$i].Name
    $ctci.Id=$CTXML[$i].ID
    $ctci.Description=$CTXML[$i].Description
    $ctci.group = $CTXML[$i].Group
    $ctci = $contentTypes.add($ctci)
    $Context.load($ctci)
    $Context.executeQuery()

    # get the new content type object
    $SPContentType = $Context.web.contenttypes.getbyid($ctci.id)
    $newCTFields= $SPContentType.Fields
    $Context.Load($newCTFields)
    $Context.ExecuteQuery()
                
    $fieldsXML=($CTXML[$i].Fields | ConvertFrom-Json)
    for($j=0; $j -lt $fieldsXML.Length; $j++) {
        $ctFieldIsExist="false"
        foreach($ctf in $newCTFields) {
           if($fieldsXML[$j].InternalName -eq $ctf.InternalName) 
           {
              $ctFieldIsExist="true";
               break;
           } 
       } 
        if ($ctFieldIsExist -eq "true"){
            Write-Host "--- ContentType Field [$($fieldsXML[$j].InternalName)] -> Match [$ctFieldIsExist]" -ForegroundColor Green
        }
        else
        {
            Write-Host "--- ContentType Field [$($fieldsXML[$j].InternalName)] -> Match [$ctFieldIsExist]" -ForegroundColor White
            #Add Fields Reference to the New content type
            if(!$SPContentType.FieldLinks[$fieldsXML[$j].InternalName])
            {
                $field = $Context.Site.RootWeb.Fields.GetByInternalNameOrTitle($fieldsXML[$j].InternalName)
                $Context.Load($field) 
                $Context.ExecuteQuery()
                #Create a field link for the Content Type by getting an existing column
                $flci = new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
           
                $flci.Field = $field
                #Check to see if column is Optional, Required or Hidden
                if($flci.Field.InternalName -ne "Title")
                 {
                    #Add column to Content Type
                    $SPContentType.FieldLinks.Add($flci)
                    $SPContentType.Update($true)
                    $Context.Load($SPContentType)
                    $Context.ExecuteQuery()
                    $SPContentType=$Context.web.contenttypes.getbyid($ctci.Id.StringValue)                        
                    $Context.Load($SPContentType)
                    $Context.ExecuteQuery()
                    $Field=$SPContentType.Fields.GetByInternalNameOrTitle($fieldsXML[$j].InternalName)
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
                    $SPContentType.Update($true)
                    $Context.Load($SPContentType)
                    $Context.ExecuteQuery()  
                }
            }  
        }                            
    }    
}
}