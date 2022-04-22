# MAIN PIECES OF LOGIC USED
# New-Mailbox -Shared -Name "$Email" -DisplayName "$DisplayName" -PrimarySmtPAddress "$Email" -Alias $EmailFront
# Set-Mailbox -Identity $Email -EmailAddress @{Add= "$Alias"}
# Add-MailboxPermission -Identity $Email -User $UPN -AccessRights FullAccess -InheritanceType All 
# Add-RecipientPermission "$Email" -AccessRights SendAs -Trustee "$UPN" -Confirm:$false

# ADDITIONAL LOGIC THAT MIGHT BE USEFULL IF THE ADDRESS IS ALREADY TAKEN
# try {
#     Remove-Mailbox $Email -Confirm:$false
#     Write-Host -ForegroundColor Red "_____________________________" 
#     Write-Host -ForegroundColor Red "Removing $Email Mailbox" 
#     Write-Host -ForegroundColor Red "-----------------------------" 
# }
# catch {
#     Write-Host "Didn't exist, creating new "
# }
# try {
#     Remove-DistributionGroup -Identity $Email -Confirm:$false
#     Write-Host -ForegroundColor Red "_____________________________" 
#     Write-Host -ForegroundColor Red "Removing $Email Distribution Group" 
#     Write-Host -ForegroundColor Red "-----------------------------" 
# }
# catch {
#     Write-Host "Didn't exist, creating new "
# }

# USAGE

# .\MailboxCreator.ps1	"C:\Users\StefanKubisa\Documents\Scripts\SharedMailboxCreation.xlsx"	Sheet1

# Adding mailboxes, users and aliases

# Status	    Display Name                Mailbox	                        User 1 / Alias 1	                    User 2 / Alias 2	                    User 3 / Alias 3
#               Automation Test Mailbox 1	test.automailbox.1@domain.com	name.surname@domain.com	                name2.surname2@domain.com	            name3.surname3@domain.com
#               Automation Test Mailbox 1	test.automailbox.1@domain.com	test.automailbox.1.alias1@domain.com	test.automailbox.1.alias2@domain.com	test.automailbox.1.alias3@domain.com
#               Automation Test Mailbox 2	test.automailbox.2@domain.com	name.surname@domain.com	                name2.surname2@domain.com	            name3.surname3@domain.com
#               Automation Test Mailbox 2	test.automailbox.2@domain.com	test.automailbox.2.alias1@domain.com	test.automailbox.2.alias2@domain.com	test.automailbox.2.alias3@domain.com

# Adding users to mailboxes

# Status	    Display Name                Mailbox	                        User 1 	                                User 2 	                                User 3 
#               Automation Test Mailbox 1	test.automailbox.1@domain.com	name.surname@domain.com	                name2.surname2@domain.com	            name3.surname3@domain.com
#               Automation Test Mailbox 2	test.automailbox.2@domain.com	name.surname@domain.com	                name2.surname2@domain.com	            name3.surname3@domain.com

# Adding aliases to mailboxes

# Status	    Display Name                Mailbox	                        Alias 1	                                Alias 2	                                Alias 3
#               Automation Test Mailbox 1	test.automailbox.1@domain.com	test.automailbox.1.alias1@domain.com	test.automailbox.1.alias2@domain.com	test.automailbox.1.alias3@domain.com
#               Automation Test Mailbox 2	test.automailbox.2@domain.com	test.automailbox.2.alias1@domain.com	test.automailbox.2.alias2@domain.com	test.automailbox.2.alias3@domain.com

# ----------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------
# !!!!!!!!!!!! THE 5 SECONDS OF SLEEP ARE NEEDED FOR EXCHANGE TO RELIABLY SET THE ENTRIES !!!!!!!!!!!!
# ----------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------
cls

# Set-ExecutionPolicy Bypass -Scope Process

# Installing nuget and exchang online goes here 

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
$mailboxCount = $verticalCount
Write-Host -ForegroundColor DarkGreen "Mailbox Count: $mailboxCount"
Write-Host "If this makes sense press any key, if not CTRL + C"
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host ""

Write-Host "
Which authentication method would you like to use:
--------------------------------------------------
1. Autoauthentication via UPN (User Principal Name)
2. Autoauthentication via prompt - it might not spring in front of you and be below on the taskbar 
3. Type UPN manually into the prompt 
4. Hardcoded Authentication in the file
5. Already have an Exchange session going"

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
    $UPN = "stefan.kubisa@sellerx.com"
    Connect-ExchangeOnline -UserPrincipalName $UPN
}
5 {}
}

Write-Host "
What would you like to do:
--------------------------------------------------
1. Create Mailboxes, add users and add aliases
2. Create Mailboxes and add users
3. Create Mailboxes and add aliases 
4. Create Mailboxes only 
5. Add users to mailboxes only
6. Add aliases to mailboxes only
7. Add users and aliases to a mailbox" 

$case = Read-Host "It's time to choose"

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
        Write-Host -ForegroundColor DarkGreen "-----------------------------" 

        if($case -eq 1 -or $case -eq 2 -or $case -eq 3 -or $case -eq 4) { 
            New-Mailbox -Shared -Name "$Email" -DisplayName "$DisplayName" -PrimarySmtPAddress "$Email" -Alias $EmailFront 
        } 

        if($case -eq 1 -or $case -eq 2 -or $case -eq 5 -or $case -eq 7) {
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
                $Worksheet.Cells.Item($i + 1, $j + 4).Interior.ColorIndex = 43
                Start-Sleep -Seconds 5
            }
            Write-Host "--------------"
            $Worksheet.Cells.Item($i + 1, $IsItDone_Pos) = "OK"
            $Worksheet.Cells.Item($i + 1, $IsItDone_Pos).Interior.ColorIndex = 43
        }

        if($case -eq 1 -or $case -eq 7) {
            $i++
        }

        if($case -eq 1 -or $case -eq 3 -or $case -eq 6 -or $case -eq 7) {
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
                $Worksheet.Cells.Item($i + 1, $j + 4).Interior.ColorIndex = 43
                Start-Sleep -Seconds 5
            }
            Write-Host "--------------"

            Start-Sleep -Seconds 5
        }

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