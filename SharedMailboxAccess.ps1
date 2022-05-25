$Shared = Read-Host -Prompt "Enter Shared Mailbox Email"
$user = Read-Host -Prompt "Enter User Email"
$status = Get-Mailbox $Shared -ErrorAction SilentlyContinue
if($status -eq $null)
{
    Write-Host "`n$Shared could not be found OR does not exist" -ForegroundColor Yellow
}
else
{
    $users = Get-MailboxPermission $Shared | select -ExpandProperty user
    if($users -contains $user)
    {
        Write-Host "`n$user already has access to $Shared`n" -ForegroundColor Green
    }
    else
    {
        Write-Host "`n$user does not have access to $Shared`n" -ForegroundColor Yellow
        Write-Host "1. Grant Access"
        Write-Host "2. Exit"
        $input = Read-Host -Prompt "`nSelect action from above (1 or 2)"
        switch($input)
        {
            1 {Add-MailboxPermission $Shared -User $user -AccessRights fullaccess
               Add-RecipientPermission $Shared -Trustee $user -AccessRights sendas -Confirm:$false;break}
            2 {;break}
        }
    }
}

