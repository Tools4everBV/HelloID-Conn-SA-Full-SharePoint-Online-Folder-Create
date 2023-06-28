# script
$body = @{
    "client_id"=$AADAppId
    "scope"="https://graph.microsoft.com/.default"
    "client_secret"=$AADAppSecret
    "grant_type"="client_credentials"
}
$siteid = $form.dropDownSites.id
$sitename = $form.dropDownSites.name
$newfoldername = $form.folderName
$readPermissions = $form.dualListRead.Right
$writePermissions = $form.dualListWrite.Right

$tokenquery = Invoke-RestMethod -uri https://login.microsoftonline.com/$($AADtenantID)/oauth2/v2.0/token -body $body -Method Post -ContentType 'application/x-www-form-urlencoded'

$baseGraphUri = "https://graph.microsoft.com/"
$headers = @{
    "content-type" = "Application/Json"
    "authorization" = "Bearer $($tokenquery.access_token)"
}
	
$created = $false
try {
    $bodypost = @{
        "name" = $newfoldername
        "folder" = @{}
    }
    $a = Invoke-RestMethod -uri "$baseGraphUri/v1.0/sites/$siteid/drive/root/children" -Method POST -body ($bodypost | ConvertTo-Json) -Headers $headers
    Write-Information "Folder created: $newfoldername"

    $Log = @{
            Action            = "CreateGroup" # optional. ENUM (undefined = default) 
            System            = "Sharepoint" # optional (free format text) 
            Message           = "Successfully created folder:  [$($newfoldername)]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $sitename # optional (free format text) 
            TargetIdentifier  = $newfoldername # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
    
    $mailnick = $sitename + "_" + $newfoldername
    $mailnick = $mailnick -replace " ", "_"
    $bodygroupread = @{
        "description" = "$($a.id) - READ"
        "displayName" = "Read Group for Site $sitename and folder $newfoldername"
        "mailEnabled" = $false
        "mailNickName" = $mailnick + "_read"
        "securityEnabled" = $true
    }
    $aread = Invoke-RestMethod -uri "$baseGraphUri/v1.0/groups" -Method POST -body ($bodygroupread | ConvertTo-Json) -Headers $headers
    Write-Information "Read Group created for folder"

    $Log = @{
            Action            = "CreateGroup" # optional. ENUM (undefined = default) 
            System            = "Sharepoint" # optional (free format text) 
            Message           = "Successfully created read group for folder:  [$($newfoldername)]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $sitename # optional (free format text) 
            TargetIdentifier  = $newfoldername # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
    
    $bodygroupwrite = @{
        "description" = "$($a.id) - WRITE"
        "displayName" = "Write Group for Site $sitename and folder $newfoldername"
        "mailEnabled" = $false
        "mailNickName" = $mailnick + "_write"
        "securityEnabled" = $true
    }
    $awrite = Invoke-RestMethod -uri "$baseGraphUri/v1.0/groups" -Method POST -body ($bodygroupwrite | ConvertTo-Json) -Headers $headers
    Write-Information "Write Group created for folder"

    $Log = @{
            Action            = "CreateGroup" # optional. ENUM (undefined = default) 
            System            = "Sharepoint" # optional (free format text) 
            Message           = "Successfully created write group for folder:  [$($newfoldername)]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $sitename # optional (free format text) 
            TargetIdentifier  = $newfoldername # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log

    $created = $true
}
catch {
    Write-Error "Error occured while creating folder. Error $_"

    $Log = @{
            Action            = "CreateGroup" # optional. ENUM (undefined = default) 
            System            = "Sharepoint" # optional (free format text) 
            Message           = "Error occured while creating folder: [$($newfoldername)]. Error $_" # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $sitename # optional (free format text) 
            TargetIdentifier  = $newfoldername # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}

if ($created) {
    for ($i = 0; $i -lt 20; $i++)
    {  
        try {            
            $bodyinviteread = @{
                "requireSignIn" = $true
                "sendInvitation" = $false
                "roles" = @("read")
                "recipients" =  @(@{"objectId" = $aread.id})
            }
            $ainviteread = Invoke-RestMethod -uri "$baseGraphUri/v1.0/sites/$siteid/drive/items/$($a.id)/invite" -Method POST -body ($bodyinviteread | ConvertTo-Json) -Headers $headers
            Write-Information "Read Group invited to folder"

            $Log = @{
                Action            = "CreateGroup" # optional. ENUM (undefined = default) 
                System            = "Sharepoint" # optional (free format text) 
                Message           = "Read Group invited to folder:  [$($newfoldername)]" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $sitename # optional (free format text) 
                TargetIdentifier  = $newfoldername # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        
            $bodyinvitewrite = @{
                "requireSignIn" = $true
                "sendInvitation" = $false
                "roles" = @("write")
                "recipients" =  @(@{"objectId" = $awrite.id})
            }
            $ainviteread = Invoke-RestMethod -uri "$baseGraphUri/v1.0/sites/$siteid/drive/items/$($a.id)/invite" -Method POST -body ($bodyinvitewrite | ConvertTo-Json) -Headers $headers
            
            Write-Information "Write Group invited to folder"
            $Log = @{
                Action            = "CreateGroup" # optional. ENUM (undefined = default) 
                System            = "Sharepoint" # optional (free format text) 
                Message           = "Write Group invited to folder:  [$($newfoldername)]" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $sitename # optional (free format text) 
                TargetIdentifier  = $newfoldername # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
            break
        }
        catch {
            Start-Sleep -Seconds 20
        }
    }
        
    if($readPermissions -ne $null){
        try {
            foreach($user in $readPermissions){
                $addGroupMembershipUri = $baseGraphUri + "v1.0/groups/$($aread.id)/members" + '/$ref'
                $body = @{ "@odata.id"= "$baseGraphUri/v1.0/users/$($user.id)" } | ConvertTo-Json -Depth 10
    
                $response = Invoke-RestMethod -Method POST -Uri $addGroupMembershipUri -Body $body -Headers $headers -Verbose:$false
            }
    
            Write-Information "Finished adding AzureAD users [$($readPermissions | ConvertTo-Json)] to AzureAD group [$($bodygroupread.displayName)]"

            $Log = @{
                Action            = "AddMembers" # optional. ENUM (undefined = default) 
                System            = "Sharepoint" # optional (free format text) 
                Message           = "Finished adding AzureAD users [$($readPermissions | ConvertTo-Json)] to AzureAD group [$($bodygroupread.displayName)]" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $sitename # optional (free format text) 
                TargetIdentifier  = $newfoldername # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        } catch {
            Write-Error "Could not add AzureAD users [$($readPermissions | ConvertTo-Json)] to AzureAD group [$($bodygroupread.displayName)]. Error: $($_.Exception.Message)"

            $Log = @{
                Action            = "AddMembers" # optional. ENUM (undefined = default) 
                System            = "Sharepoint" # optional (free format text) 
                Message           = "Could not add AzureAD users [$($readPermissions | ConvertTo-Json)] to AzureAD group [$($bodygroupread.displayName)]. Error: $($_.Exception.Message)" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $sitename # optional (free format text) 
                TargetIdentifier  = $newfoldername # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        }
    }
    if($writePermissions -ne $null){
        try {
            foreach($user in $writePermissions){                
                $addGroupMembershipUri = $baseGraphUri + "v1.0/groups/$($awrite.id)/members" + '/$ref'
                $body = @{ "@odata.id"= "$baseGraphUri/v1.0/users/$($user.id)" } | ConvertTo-Json -Depth 10
    
                $response = Invoke-RestMethod -Method POST -Uri $addGroupMembershipUri -Body $body -Headers $headers -Verbose:$false
            }
    
            Write-Information "Finished adding AzureAD users [$($writePermissions | ConvertTo-Json)] to AzureAD group [$($bodygroupwrite.displayName)]"

            $Log = @{
                Action            = "AddMembers" # optional. ENUM (undefined = default) 
                System            = "Sharepoint" # optional (free format text) 
                Message           = "Finished adding AzureAD users [$($writePermissions | ConvertTo-Json)] to AzureAD group [$($bodygroupwrite.displayName)]" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $sitename # optional (free format text) 
                TargetIdentifier  = $newfoldername # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log

        } catch {
            Write-Error "Could not add AzureAD users [$($writePermissions | ConvertTo-Json)] to AzureAD group [$($bodygroupwrite.displayName)]. Error: $($_.Exception.Message)"

            $Log = @{
                Action            = "AddMembers" # optional. ENUM (undefined = default) 
                System            = "Sharepoint" # optional (free format text) 
                Message           = "Could not add AzureAD users [$($writePermissions | ConvertTo-Json)] to AzureAD group [$($bodygroupwrite.displayName)]. Error: $($_.Exception.Message)" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $sitename # optional (free format text) 
                TargetIdentifier  = $newfoldername # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        }
    }
}
