#START

$result = "C:\temp\Room_Details.csv"
$out = @()

$rooms = Get-EXOMailbox -RecipientTypeDetails RoomMailbox | select DisplayName,PrimarySMTPAddress,RecipientTypeDetails

foreach($room in $rooms)
{
	$cal=Get-CalendarProcessing $room.primarysmtpaddress | select AutomateProcessing,AllowConflicts,AllowRecurringMeetings,ConflictPercentageAllowed,MaximumConflictInstances,EnableResponseDetails
	
	$property = [PSCustomObject]@{	DisplayName=$room.DisplayName
					Email=$room.primarysmtpaddress
					RecipientType=$room.recipienttypedetails
					AutomateProcessing=$cal.automateprocessing
					AllowConflicts=$cal.AllowConflicts
					AllowRecurringMeetings=$cal.AllowRecurringMeetings
					ConflictPercentageAllowed=$cal.ConflictPercentageAllowed
					MaximumConflictInstances=$cal.MaximumConflictInstances
					EnableResponseDetails=$cal.EnableResponseDetails
				     }
	$out += $property
}

$out | Export-Csv $result -NoTypeInformation -Encoding UTF8

#END