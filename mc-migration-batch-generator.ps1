#Name: Exchange Migration batch generator
#Autor: Viktor Ahorner
#Webside: blog.mccloud.cloud
#Version: 0.1 public

# SCRIPT COnfiguration
$global:exportlocation = 'C:\temp\' #Store-Location for Export and Import
[long]$global:dayliimit = 500000000000 #Data-Transferlimit per batchjob in bytes > Actual size is 500GB per batchjob, it requires upload speed of around 35MBit/s
#----------------------
function MCIntro()
{
Write-Host "                                          MCCloud.cloud solution                                    "
Write-Host "                                                  presents                                          "
Start-Sleep -Seconds 5
}
function title()
{
Write-Host '                                        Exchange Migrationbatch                                     ' -BackgroundColor DarkGreen
Write-Host '                                              GENERATOR                                             ' -BackgroundColor DarkGreen
}
function Load-ExchangeModules()
{
Write-Host 'Looking for Exchange-Modules' -ForegroundColor DarkGray
    #Add Exchange snapin if not already loaded
    if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
    {
	    Write-Verbose "Loading the Exchange 2010 snapin"
	    try
	    {
		    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	    }
	    catch
	    {
		    #Snapin not loaded
		    Write-Warning $_.Exception.Message
		    EXIT
	    }
	    . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	    Connect-ExchangeServer -auto -AllowClobber
    }
}

function Load-MailboxStatistics()
{
#--- Export Mailboxstatistics
Write-Host 'Starting to analyze mailbox statistics' -ForegroundColor DarkGray
$mailboxstatistics = @()
$stats = @()

#--- ALL Exchange servers START
#---------- If you want to schedule all your exchange servers, use this block#

$exchServers = Get-ExchangeServer
#---- All Exchange servers END

#---- Specific exchange servers START
#---------- If you want to schedule specific exchange servers, use this block#
#$exchangeservers =@('EXCH01','EXCH02','EXCH03') #--- List of ExchangeServers separated by coma
#$exchServers=@()
#foreach($server in $exchangeservers)
#{
#$exchServers += (get-exchangeserver -identity $server)
#}

#-----Specific exchange servers END
    foreach ($Exch in $exchServers)
    {
        $mailboxstatistics =@()
        $stats = Get-Mailbox -Server $Exch.FQDN -ResultSize unlimited
        foreach($stat in $stats)
        { 
        $mailboxstatistic=@()
        $statistics = Get-MailboxStatistics -Identity $Stat.Identity
                $object = New-Object -TypeName PSObject
                $object | Add-Member -MemberType Noteproperty -Name DisplayName $stat.Displayname
                $object | Add-Member -MemberType Noteproperty -Name TotalItemSIze $statistics.TotalItemSIze
                $object | Add-Member -MemberType Noteproperty -Name ItemCount $statistics.ItemCount
                $object | Add-Member -MemberType Noteproperty -Name PrimarySmtpAddress $stat.PrimarySmtpAddress.Address
                $mailboxstatistic += $object
        if($stat.displayname -inotlike 'Discovery Search Mailbox*')
        {
        $mailboxstatistics += $mailboxstatistic
        }
        }
        $stats = Get-Mailbox -Server $Exch.FQDN -RecipientTypeDetails SharedMailbox -ResultSize unlimited
        foreach($stat in $stats)
        { 
        $mailboxstatistic=@()
        $statistics = Get-MailboxStatistics -Identity $Stat.Identity
                $object = New-Object -TypeName PSObject
                $object | Add-Member -MemberType Noteproperty -Name DisplayName $stat.Displayname
                $object | Add-Member -MemberType Noteproperty -Name TotalItemSIze $statistics.TotalItemSIze
                $object | Add-Member -MemberType Noteproperty -Name ItemCount $statistics.ItemCount
                $object | Add-Member -MemberType Noteproperty -Name PrimarySmtpAddress $stat.PrimarySmtpAddress.Address
                $mailboxstatistic += $object
        if($stat.displayname -inotlike 'Discovery Search Mailbox*')
        {
        $mailboxstatistics += $mailboxstatistic
        }        }
        $stats = Get-Mailbox -Server $Exch.FQDN -RecipientTypeDetails RoomMailbox -ResultSize unlimited
        foreach($stat in $stats)
        { 
        $mailboxstatistic=@()
        $statistics = Get-MailboxStatistics -Identity $Stat.Identity
                $object = New-Object -TypeName PSObject
                $object | Add-Member -MemberType Noteproperty -Name DisplayName $stat.Displayname
                $object | Add-Member -MemberType Noteproperty -Name TotalItemSIze $statistics.TotalItemSIze
                $object | Add-Member -MemberType Noteproperty -Name ItemCount $statistics.ItemCount
                $object | Add-Member -MemberType Noteproperty -Name PrimarySmtpAddress $stat.PrimarySmtpAddress.Address
                $mailboxstatistic += $object
        if($stat.displayname -inotlike 'Discovery Search Mailbox*')
        {
        $mailboxstatistics += $mailboxstatistic
        }        }
        $stats = Get-Mailbox -Server $Exch.FQDN -RecipientTypeDetails LinkedMailbox -ResultSize unlimited
        foreach($stat in $stats)
        { 
        $mailboxstatistic=@()
        $statistics = Get-MailboxStatistics -Identity $Stat.Identity
                $object = New-Object -TypeName PSObject
                $object | Add-Member -MemberType Noteproperty -Name DisplayName $stat.Displayname
                $object | Add-Member -MemberType Noteproperty -Name TotalItemSIze $statistics.TotalItemSIze
                $object | Add-Member -MemberType Noteproperty -Name ItemCount $statistics.ItemCount
                $object | Add-Member -MemberType Noteproperty -Name PrimarySmtpAddress $stat.PrimarySmtpAddress.Address
                $mailboxstatistic += $object
        if($stat.displayname -inotlike 'Discovery Search Mailbox*')
        {
        $mailboxstatistics += $mailboxstatistic
        }        }
        $stats = Get-Mailbox -Server $Exch.FQDN -RecipientTypeDetails LinkedRoomMailbox -ResultSize unlimited
        foreach($stat in $stats)
        {
        $mailboxstatistic=@()
        $statistics = Get-MailboxStatistics -Identity $Stat.Identity
                $object = New-Object -TypeName PSObject
                $object | Add-Member -MemberType Noteproperty -Name DisplayName $stat.Displayname
                $object | Add-Member -MemberType Noteproperty -Name TotalItemSIze $statistics.TotalItemSIze
                $object | Add-Member -MemberType Noteproperty -Name ItemCount $statistics.ItemCount
                $object | Add-Member -MemberType Noteproperty -Name PrimarySmtpAddress $stat.PrimarySmtpAddress.Address
                $mailboxstatistic += $object
        if($stat.displayname -inotlike 'Discovery Search Mailbox*')
        {
        $mailboxstatistics += $mailboxstatistic
        }        }
                $stats = Get-Mailbox -Server $Exch.FQDN -RecipientTypeDetails EquipmentMailbox -ResultSize unlimited
        foreach($stat in $stats)
        { 
        $mailboxstatistic=@()
        $statistics = Get-MailboxStatistics -Identity $Stat.Identity
                $object = New-Object -TypeName PSObject
                $object | Add-Member -MemberType Noteproperty -Name DisplayName $stat.Displayname
                $object | Add-Member -MemberType Noteproperty -Name TotalItemSIze $statistics.TotalItemSIze
                $object | Add-Member -MemberType Noteproperty -Name ItemCount $statistics.ItemCount
                $object | Add-Member -MemberType Noteproperty -Name PrimarySmtpAddress $stat.PrimarySmtpAddress.Address
                $mailboxstatistic += $object
        if($stat.displayname -inotlike 'Discovery Search Mailbox*')
        {
        $mailboxstatistics += $mailboxstatistic
        }        }
    }
$mailboxstatistics | sort -Descending TotalItemSIze -unique
$mailboxstatistics | ConvertTo-Csv -Delimiter ';' | Out-File -FilePath ($exportlocation+'mailbosstatistics-export.csv')
return $mailboxstatistics
}

function Schedule-Batches
{
$filepath = ($global:exportlocation+'mailbosstatistics-export.csv')
[System.Collections.ArrayList]$mailbox_schedule = @()
try
{
$mailboxes = Get-Content -Path $filepath -ErrorAction stop
$mailboxes = $mailboxes -replace ('"','')
}
catch
{
Write-Host 'Export file not found'-BackgroundColor DarkCyan
Write-Host 'Please validate if '$filepath' is valid' -BackgroundColor DarkCyan
return $null
}
[string]$totalitemsize ='' 
[long]$totalsum=0
foreach ($counter in (2..($mailboxes.Length-1)))
{
$raw = $mailboxes[$counter]
$totalitemsize = ($raw.split(';')[1]).Split('(')[1]
$totalitemsize = $totalitemsize.Split(')')[0]
$totalitemsize = $totalitemsize.Split(' ')[0]
$totalitemsize = $totalitemsize.Replace(',','')
[long]$intotal = $totalitemsize
                $object = New-Object -TypeName PSObject
                $object | Add-Member -MemberType Noteproperty -Name DisplayName $raw.split(';')[0] #Add Displayname to the migration batches lists
                $object | Add-Member -MemberType Noteproperty -Name TotalItemSIze $intotal #Add Totalitemsize to the migration batches lists
                $object | Add-Member -MemberType Noteproperty -Name ItemCount $raw.split(';')[2] #Add ItemCound to the migration batches lists
                $object | Add-Member -MemberType Noteproperty -Name PrimarySmtpAddress $raw.split(';')[3] #Add PrimarySMTPAddress to the migration batches lists

                $mailbox_schedule += $object
$totalsum += $intotal
}
[int]$migrationbatches = ([math]::Round(($totalsum / $dayliimit))+1)
$batchtotal = 0
$Mailbox =0
[System.Collections.ArrayList]$migrationbatch=@()
if($mailbox_schedule.Count -gt 0)
{
foreach ($batchnr in (0..($migrationbatches-1)))
{
$batchtotal = 0                                            #Set total batchsize to 0
[System.Collections.ArrayList]$tempbatch =@()              #Define temporare batchjob, this is each single batch job
    while(($batchtotal+$subsum) -lt $dayliimit)            #start to loop as long total data amount smaller data limit each week
    { 
     Write-Host 'Scheduling '$($mailbox_schedule[$Mailbox]) -ForegroundColor Cyan

    $tempbatchsize = $tempbatch.add($mailbox_schedule[$Mailbox])            #Scheduling current biggest mailbox 
        $batchtotal+= $($mailbox_schedule[$Mailbox].TotalitemSize) #adding current biggest mailbox size to batchtotal size
        #Start-Sleep -Seconds 5                            #debugging line
        $subsum = 0                                        #setting up current biggest mailbox size as limit for smaller batches
                [int]$counter = 1
                while($subsum -le $mailbox_schedule[$Mailbox].TotalitemSize -and ($batchtotal+$subsum) -lt $dayliimit)
                {
                    #$counter +=1
                    $subsum += $mailbox_schedule[$mailbox_schedule.count - 1].TotalItemSIze
                    Write-Host 'Scheduling '$($mailbox_schedule[$mailbox_schedule.count - $counter]) -ForegroundColor Cyan
                    $tempbatch +=$mailbox_schedule[$mailbox_schedule.count - 1]
                    Write-Host 'Remove '$mailbox_schedule[$mailbox_schedule.count - 1]' from collection' -ForegroundColor DarkYellow
                    try
                    {

                   $mailbox_schedule.Removeat($mailbox_schedule.count - 1)
                  }
                   catch
                   {

                   }
                }
        try
        {
        $mailbox_schedule.Remove($mailbox_schedule[$Mailbox])
        }
        catch
        {
        }

        $batchtotal += $subsum
        if($mailbox_schedule.Count -gt 0)
        {
        }
        else
        {
        break
        }
    }
        $batchamount = $migrationbatch.Add($tempbatch)
 
}
}
else
{
    Write-Host 'ALL ITEMS ARE SCHEDULED' -BackgroundColor DarkBlue -ForegroundColor White
}

if($migrationbatches -gt 1)
{
    $forend = ($migrationbatches-1)
}
    foreach($singlebatch in (0..$forend))
    {
        $migrationbatch[$singlebatch]| ConvertTo-Csv -Delimiter ';' | Out-File ($global:exportlocation+'Migrationbatch_'+($singlebatch+1)+'.csv')
    }
Write-Host 'All '$migrationbatches' migration batches have been exported into '$global:exportlocation -BackgroundColor DarkGreen -ForegroundColor White
}

MCIntro
title
Load-ExchangeModules
Load-MailboxStatistics
Schedule-Batches
