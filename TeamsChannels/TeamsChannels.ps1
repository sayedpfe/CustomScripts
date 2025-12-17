Connect-mgGraph -Scopes "ChannelSettings.Read.All, Channel.ReadBasic.All, ChannelSettings.ReadWrite.All, ChannelSettings.Read.Group"
Get-MgAllTeamChannel -TeamId 38d37704-57f3-4d1e-b696-6b3ee897540d
$team = Get-Team -GroupId 38d37704-57f3-4d1e-b696-6b3ee897540d
Get-TeamChannel -GroupId $team.GroupId |
>>   Select DisplayName, MembershipType, Id