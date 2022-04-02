# MAIN PIECES OF LOGIC USED
# New-Mailbox -Shared -Name "$Email" -DisplayName "$DisplayName" -PrimarySmtPAddress "$Email" -Alias $EmailFront
# Set-Mailbox -Identity $Email -EmailAddress @{Add= "$Alias"}
# Add-MailboxPermission -Identity $Email -User $UPN -AccessRights FullAccess -InheritanceType All 
# Add-RecipientPermission "$Email" -AccessRights SendAs -Trustee "$UPN" -Confirm:$false

# ADDITIONAL LOGIC THAT MIGHT BE USEFULL IF THE ADDRESS IS ALREADY TAKEN
# try {
#     Remove-Mailbox $Email -Confirm:$false
#     Write-Host -ForegroundColor Red "_____________________________" 
#     Write-Host -ForegroundColor Red "Removing $Email" 
#     Write-Host -ForegroundColor Red "-----------------------------" 
# }
# catch {
#     Write-Host "Didn't exist, creating new "
# }
# try {
#     Remove-DistributionGroup -Identity $Email -Confirm:$false
#     Write-Host -ForegroundColor Red "_____________________________" 
#     Write-Host -ForegroundColor Red "Removing $Email" 
#     Write-Host -ForegroundColor Red "-----------------------------" 
# }
# catch {
#     Write-Host "Didn't exist, creating new "
# }

cls

$IsItDone_Pos = 1
$DisplayName_Pos = 2
$Email_Pos = 3
$User_Or_Alias_Pos = 4

$Path = $args[0]
$CurrentWorksheet = $args[1]

Write-Host "Creating Excel App Object" 
$excel = new-object -comobject Excel.Application 
$excel.visible = $true 
$excel.DisplayAlerts = $false 
$excel.WindowState = "xlMaximized"
Write-Host "Opening Workbook"
Write-Host "____________________"

try {
    $workbook = $excel.workbooks.open($path)
}
catch {
    Write-Host $_.Exception.Message
    Write-Host "Closing Excel"
    Start-Sleep -Seconds 5
    $excel.Quit()
    throw $_
}
try {
    $Worksheet = $workbook.Worksheets.item($CurrentWorksheet)
}
catch {
    Write-Output $_.Exception.Message
    Write-Output "Closing Workbook"
    $workbook.Close()
    Write-Output "Closing Excel"
    Start-Sleep -Seconds 5
    $excel.Quit()
    throw $_
}

$verticalCount = (($Worksheet.UsedRange.Rows).count - 1 )
$horizontalCount = ($Worksheet.UsedRange.Columns).count - 3
$mailboxCount = $verticalCount / 2
Write-Host -ForegroundColor DarkGreen "Mailbox Count: $mailboxCount"
Write-Host "If this makes sense press any key, if not CTRL + C"
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host ""

Write-Host "
Which authentication method would you like to use:
--------------------------------------------------
1. Autoauthentication via UPN
2. Autoauthentication via prompt
3. Type UPN manually
4. Hardcoded Authentication in the file
5. Exchange module YOLO cause sometimes this works"

$number = Read-Host "Time to choose Dr. Freeman"
switch ($number){
1 {
    $UPN = whoami /UPN
    Connect-ExchangeOnline -UserPrincipalName $UPN
}
2 {
    Connect-ExchangeOnline
}
3 {
    $UPN = Read-Host
    Connect-ExchangeOnline -UserPrincipalName $UPN
}
4 {
    $UPN = "INSERT YOUR UPN HERE"
    Connect-ExchangeOnline -UserPrincipalName $UPN
}
5 {}
}

$keepGoing = $true
while ($keepGoing) {
    for ($i = 1; $i -lt $verticalCount + 1; $i++) { 
        if (($Worksheet.Cells.Item($i + 1, $IsItDone_Pos)).Text -eq "OK" -or ($Worksheet.Cells.Item($i + 1, $IsItDone_Pos)).Text -eq "SKIP") {
            continue
        }
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos) = "In Progress"
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos).Interior.ColorIndex = 44
        $DisplayName = $Worksheet.Cells.Item($i + 1, $DisplayName_Pos).Text 
        $Email = $Worksheet.Cells.Item($i + 1, $Email_Pos).Text 
        $User = $Worksheet.Cells.Item($i + 1, $User_Or_Alias_Pos).Text 
        $EmailSplit = ($Email).split("@") 
        $TextInfo = (Get-Culture).TextInfo 
        $EmailFront = $TextInfo.ToTitleCase($EmailSplit[0]) 

        Write-Host -ForegroundColor DarkGreen "_____________________________" 
        Write-Host -ForegroundColor DarkGreen "$DisplayName" 
        Write-Host -ForegroundColor DarkGreen "$Email" 
        Write-Host -ForegroundColor DarkGreen "$User" 
        Write-Host -ForegroundColor DarkGreen "-----------------------------" 

        New-Mailbox -Shared -Name "$Email" -DisplayName "$DisplayName" -PrimarySmtPAddress "$Email" -Alias $EmailFront 

        for ($j = 0; $j -lt $horizontalCount; $j++) {
            if (($Worksheet.Cells.Item($i + 1, $j + 4).Text -eq "")) {
                continue
            }
            $User = $Worksheet.Cells.Item($i + 1, $j + 4).Text 
            
            Write-Host -ForegroundColor Blue "_____________________________" 
            Write-Host -ForegroundColor Blue "Adding Read Access to Shared Mailbox $DisplayName for User with the Email $User"
            Write-Host -ForegroundColor Blue "--------------"

            Add-MailboxPermission -Identity $Email -User $User -AccessRights FullAccess -InheritanceType All 

            Write-Host -ForegroundColor Blue "_____________________________" 
            Write-Host -ForegroundColor Blue "Adding Send As Access to Shared Mailbox $DisplayName for User with the Email $User"
            Write-Host -ForegroundColor Blue "--------------"

            Add-RecipientPermission "$Email" -AccessRights SendAs -Trustee "$User" -Confirm:$false

            Write-Host -ForegroundColor Red "GIVING EXCHANGE TIME TO PROCESS THE CHANGES"
            Start-Sleep -Seconds 5
        }
        Write-Host "--------------"
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos) = "OK"
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos).Interior.ColorIndex = 43
        $i++

        for ($j = 0; $j -lt $horizontalCount; $j++) {
            if (($Worksheet.Cells.Item($i + 1, $j + 4).Text -eq "")) {
                continue
            }
            $Alias = $Worksheet.Cells.Item($i + 1, $j + 4).Text 
            
            Write-Host -ForegroundColor Blue "_____________________________" 
            Write-Host -ForegroundColor Blue "Adding Alias: $Alias to Shared Mailbox with the Email $Email"
            Write-Host -ForegroundColor Blue "--------------"

            Set-Mailbox -Identity $Email -EmailAddress @{Add= "$Alias"}

            Write-Host -ForegroundColor Red "GIVING EXCHANGE TIME TO PROCESS THE CHANGES"
            Start-Sleep -Seconds 5
        }
        Write-Host "--------------"

        Start-Sleep -Seconds 5
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos) = "OK"
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos).Interior.ColorIndex = 43

    }
    $keepGoing = $false
}

$workbook.Save()
Write-Host "All done, closing workbook"
Start-Sleep -Seconds 5
$excel.Quit()
Disconnect-ExchangeOnline -Confirm:$false