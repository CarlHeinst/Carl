function Convertto-Exceldate
{
    [CmdletBinding()][Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(ValueFromPipelineByPropertyName=$true,Position=0)]$date=(Get-Date)
    )

    Begin
    {
    }
    Process
    {
    $now = Get-Date $date
    $base = get-date 01/01/1900
    $output = ( ( $now - $base ) | foreach {$_.days} ) + 2
    return $output
    }
    End
    {
    }
}

Function Convert-FromUnixdate ($UnixDate) {
   [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').`
   AddSeconds($UnixDate))
}


function Convert-UnixDatetoExcelDate
{
    [CmdletBinding()][Alias()]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)]$date
    )

    Begin
    {
    }
    Process
    {
    $now = Convert-FromUnixdate -UnixDate $date
    $base = Get-Date 01/01/1900
    $output = ( ( $now - $base ) | foreach {$_.days} ) + 2
    return $output
    }
    End
    {
    }
}