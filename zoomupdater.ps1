Param(
    [string]$curlPath = "C:\windows\system32",
    [string]$apitoken = ,
    [string]$apiuri = "https://api.zoom.us/v2/users",
    [string]$res = @(),
    [string]$stat = "&status=",
    [string]$pagenumber = "?page_number=",
    [string]$pagesize = "&page_size=",
    [int]$pagenumber_val = 1,
    [int]$pagesize_val = 300,
    [string]$pagestat_val = "status",
    [string]$alluserspath = "$PSScriptRoot\Allzoomusers.csv",
    [string]$usernumbercsv = "$PSScriptRoot\Allusersnumbers.csv",
    [string]$exceptioncsv = "$PSScriptRoot\exceptionlist.csv",
    [string]$xmlConfigFile = "$PSScriptRoot\loggerConfig.xml"
)

Import-Module New-Log4NetLogger -EA Stop
$logger = New-Log4NetLogger -XmlConfigPath $xmlConfigFile -loggerName "zoomupdater.log"
$logger.Info("Starting Script")
if (Test-Path $alluserspath ) {
    Remove-Item $alluserspath -Force
}
if (Test-Path $usernumbercsv) {
    Remove-Item $usernumbercsv -Force

}
#Setting exceptions to never process
Try {
    $exceptionlist = Import-Csv -Path $exceptioncsv | select Email
    $logger.Info("Setting exceptions")
    # remove trailing backslash from $curPath if present
    $curlPath = $curlPath.trimEnd("\")
    # setting paging options for api request 
    $pagenumber += $pagenumber_val
    $pagesize += $pagesize_val
    $stat += $pagestat_val
    $apipaging = $pagenumber + $pagesize + $stat
    # Api URL generation
    $apiuri += $apipaging

    # JWT authentication string
    $auth = 'Authorization: Bearer' + ' ' + $apitoken
    # First Api Call
    $response = & $curlPath\curl.exe --request GET --url $apiuri -H 'content-type: application/json' -H "$auth"
    $logger.Info("Doing get request to $($apiuri)")
    # Getting the number of pages to loop through
    $pagecounter = $(($response | ConvertFrom-Json).page_count)
    # Parsing Users from Json response
    $userresp = $(($response | ConvertFrom-Json).users)
    # Storing users in CSV
    $userresp | select "email", "id" | Export-Csv -Append -NoTypeInformation -Path $alluserspath
    $logger.Info("Creating csv of all users")

    do {
        $pagenumber_val++
        $pagenumber = "?page_number="
        $pagenumber += $pagenumber_val
        $apipaging = $pagenumber + $pagesize + $stat
        $apiuri = "https://api.zoom.us/v2/users"

        $logger.Info("Creating api url to curl: $($apiuri)")
        $apiuri += $apipaging

        $logger.Info("Executing api curl request to: $($apiuri)")
        $response = & $curlPath\curl.exe --request GET --url $apiuri -H 'content-type: application/json' -H "$auth"

        $logger.Info("Executing api curl request to: $($apiuri)")
        $userresp = $(($response | ConvertFrom-Json).users)

        $logger.Info("Saving users to $($alluserspath)")
        $userresp | select "email", "id" | Export-Csv -Append -NoTypeInformation -Path $alluserspath 

    }while (($null -ne $response) -and !($pagenumber_val -le $pagecounter))
}

catch {
    $logger.Error("Unable to GET users from Zoom API")  
    $logger.Error($_ -as [string]) 

}

try {
    $logger.Info("Loading users from $($alluserspath)")
    $usersimport = Import-Csv -Path $alluserspath | select email, id
    foreach ($userids in $usersimport.id) {
        if (!$exceptionlist.Email.Contains($($userids.Email))) {
            [string]$numberapiurl = "https://api.zoom.us/v2/phone/users/"
            $numberapiurl += $userids
            $logger.Info("Querying zoom phone user api: $($numberapiurl)")
            $userresponse = & $curlPath\curl.exe --request GET --url $numberapiurl -H 'content-type: application/json' -H "$auth"

            $phoneuserresp = $(($userresponse | ConvertFrom-Json))

            $newphoneuser = [PSCustomObject]@{
                Email       = $phoneuserresp.email
                Extension   = $phoneuserresp.extension_number
                PhoneNumber = $phoneuserresp.phone_numbers.number
            }
            $logger.Info("Saving $($newphonuser.Email) user information in $($usernumbercsv)")
            if ($null -ne $newphoneuser.Email) {
                $newphoneuser | select Email, Extension, PhoneNumber | Export-Csv -Append -NoTypeInformation -Path $usernumbercsv
            }
        }
    }
}
Catch {
    $logger.Error("Unable to get the user's phone number and extension.")
    $logger.Error($_ -as [string]) 

}

try {
    $logger.Info("Importing user information from $($usernumbercsv)")
    $updateDusers = Import-Csv -Path $usernumbercsv | select Email, Extension, PhoneNumber

    foreach ($newuser in $updateDusers) {
        if (!$exceptionlist.Email.Contains($($newuser.Email))) {
    
            $logger.Info("Checking for $($newuser.Email) in AD")
            $aduserobj = Get-ADUser -Filter "EmailAddress -eq '$($newuser.Email)'"
            
            if ($null -eq $aduserobj) {
                $logger.Info("Could not find $($newuser.Email) in AD")
            }
            elseif ($newuser.PhoneNumber) {
    
                Get-ADUser -Filter "EmailAddress -eq '$($newuser.Email)'" | Set-ADUser -OfficePhone $($newuser.PhoneNumber.TrimStart("+1") + " ext: " + $newuser.Extension)
                $logger.Info("Setting phone number: $($newuser.PhoneNumber.TrimStart("+1")) and extension: $($newuser.Extension) for $($newuser.Email) in AD")
    
            }
            else {
                $logger.Info("Setting extension: $($newuser.Extension) for $($newuser.Email) in AD")
                Get-ADUser -Filter "EmailAddress -eq '$($newuser.Email)'" | Set-ADUser -OfficePhone $("ext: " + $newuser.Extension)
            }
        }
    }
}
catch {
    $logger.Error("Unable to update AD user")
    $logger.Error($_ -as [string])
}