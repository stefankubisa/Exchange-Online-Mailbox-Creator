# Mailbox-Creator
Creates Shared Mailboxes in Exchange Online Adding Users and Aliases

## HowTo ##
1) Open the script by exacuting it with the following 3 parameters:

| Script File | Absolute Path to Spreadsheet | Sheet you want to use |
|----------|----------|----------|
| .\MailboxCreator.ps1 | "C:\Users\StefanKubisa\Documents\Scripts\SharedMailboxCreation.xlsx" | Sheet1 | 

Like so: 

.\MailboxCreator.ps1 "C:\Users\StefanKubisa\Documents\Scripts\SharedMailboxCreation.xlsx" Sheet1

Use cases are

1. Create Mailboxes, add users and add aliases
2. Create Mailboxes and add users
3. Create Mailboxes and add aliases 
4. Create Mailboxes only 
5. Add users to mailboxes only
6. Add aliases to mailboxes only
7. Add users and aliases to a mailbox

2) Make sure the static values match your worksheet's colum

Case 1 and 7

| Status | Display Name | Mailbox | User 1 / Alias 1 | User 2 / Alias 2 | User 3 / Alias 3 |
|----------|----------|----------|----------|----------|----------|
|  | Automation Test Mailbox 1 | `test.automailbox.1@domain.com` | name.surname@domain.com | name2.surname2@domain.com | name3.surname3@domain.com |
|  | Automation Test Mailbox 1 | `test.automailbox.1@domain.com` | `test.automailbox.1.alias1@domain.com` | `test.automailbox.1.alias2@domain.com` | `test.automailbox.1.alias3@domain.com` |
|  | Automation Test Mailbox 2 | `test.automailbox.2@domain.com` | name.surname@domain.com | name2.surname2@domain.com | name3.surname3@domain.com |
|  | Automation Test Mailbox 2 | `test.automailbox.2@domain.com` | `test.automailbox.2.alias1@domain.com` | `test.automailbox.2.alias2@domain.com` | `test.automailbox.2.alias3@domain.com` |

Case 2, 4 and 5 

| Status | Display Name | Mailbox | User 1 | User 2 | User 3 |
|----------|----------|----------|----------|----------|----------|
|  | Automation Test Mailbox 1 | `test.automailbox.1@domain.com` | name.surname@domain.com | name2.surname2@domain.com | name3.surname3@domain.com |
|  | Automation Test Mailbox 2 | `test.automailbox.2@domain.com` | name.surname@domain.com | name2.surname2@domain.com | name3.surname3@domain.com |

Case 3, 4 and 6

| Status | Display Name | Mailbox | Alias 1 | Alias 2 | Alias 3 |
|----------|----------|----------|----------|----------|----------|
|  | Automation Test Mailbox 1 | `test.automailbox.1@domain.com` | `test.automailbox.1.alias1@domain.com` | `test.automailbox.1.alias2@domain.com` | `test.automailbox.1.alias3@domain.com` |
|  | Automation Test Mailbox 2 | `test.automailbox.2@domain.com` | `test.automailbox.2.alias1@domain.com` | `test.automailbox.2.alias2@domain.com` | `test.automailbox.2.alias3@domain.com` |