#START

$out = @()
$result = "C:\Temp\MailboxSizeinMB $(Get-Date -Format dd-MMM-yyyy` HH:mm).csv"
$users = Get-EXOMailbox -ResultSize Unlimited
foreach($user in $users)
{
#Get-EXOMailboxStatistics $user.primarysmtpaddress | select @{name="TotalItemSize";E={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}
$Size = Get-EXOMailboxStatistics $user.primarysmtpaddress | select TotalItemSize,TotalDeletedItemSize,DisplayName
$SizeinMB = [math]::Round(($size.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)
$DSizeinMB = [math]::Round(($size.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)
$data = [PSCustomObject]@{	DisplayName = $Size.DisplayName
				PrimarySMTPAddress = $user.PrimarySMTPAddress
				MailboxType = $user.RecipientTypeDetails
				SizeinMB = $SizeinMB
				DeletedSizeinMB = $DSizeinMB
				TotalSize = $SizeinMB + $DSizeinMB
			 }
$out += $data
}

$out | Export-Csv $result -NoTypeInformation -Encoding UTF8

#END