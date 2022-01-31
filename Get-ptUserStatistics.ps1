function Get-ptUserStatistics {
    <#
        .SYNOPSIS
            Вывод данных пользователя по параметру samaccountname
        .DESCRIPTION
            Возвращает таблицу [hashtable] сведений о указанных пользователях
        .EXAMPLE
            "bruce.lee","bruce.willis" | Get-ptUserStatistics
        .EXAMPLE
            Get-ptUserStatistics -SamAccountName bruce.lee,bruce.willis
        .EXAMPLE
            Get-ADGroupMember wg_test | Get-ptUserStatistics
        .PARAMETER SamAccountName
            Samaccountname
        .NOTES
            Author: lybomir_dobrynin
            Creation Date: 21.01.2019
    #>

    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [System.String[]]$SamAccountName
    )

    begin {

        $Out  = @()

        $psconfig = @{
            MailServer = 'mail.domain.local'
            PDC = $env:USERDNSDOMAIN
        }

        $TestSmtpServer = Test-Connection $psconfig.MailServer -Quiet -ErrorAction SilentlyContinue
        if (!$TestSmtpServer -or !$?) {
            throw "__pterror: Ошибка при подключении к почтовому серверу"
        }

        $TestPDCServer = Test-Connection $psconfig.PDC -Quiet -ErrorAction SilentlyContinue
        if (!$TestPDCServer -or !$?) {
            throw "__pterror: Ошибка при подключении к PDC-серверу"
        }

        $CommandList = "Get-Mailbox","Get-MailboxStatistics","Get-MailboxPermission"
        $SessionProperties = @{
            ConfigurationName = 'Microsoft.Exchange'
            ConnectionUri     = "http://$($psconfig.MailServer)/powershell"
            Authentication    = 'Kerberos'
        }

        try{
            $Session = New-PSSession @SessionProperties -ErrorAction Stop
            Import-PSSession $Session -CommandName $CommandList -AllowClobber | Out-Null 
        }
        catch{
            throw "__pterror: Ошибка создания сессии"
        }
    }

    process{
        foreach ($SAN in $SamAccountName){
            $Data = New-Object PSObject -Property @{
                SamAccountName            = $SAN
                DisplayName               = $null
                TotalItemSize             = $null
                LitigationHoldEnabled     = $null
                ExchangeGuid              = $null
                ForwardingAddress         = $null
                Database                  = $null
                FullAccess                = $null
                PrimarySMTPADdress        = $null
                SipAddress                = $null
                SingleItemRecoveryEnabled = $null
                UseDatabaseQuotaDefaults  = $null
            }

            $Searcher = [ADSISearcher]"(sAMAccountName=$SAN)"
            $Searcher.SearchRoot = "LDAP://$($psconfig.PDC)/DC=domain,DC=local"
            $Searcher.PropertiesToLoad.Add('DisplayName') | Out-Null
            $Results = $Searcher.FindOne()

            if ($Results){

                $filter = {samaccountname -eq "{0}"} -f $SAN

                if (-not ($UserMbx = Get-Mailbox -filter $filter -ErrorAction SilentlyContinue -DomainController $psconfig.PDC)) {
                    $Out += $Data
                    continue
                }

                $Data.DisplayName               = $Results.Properties.Item("DisplayName")
                $Data.Database                  = $UserMbx.Database
                $Data.ExchangeGuid              = $UserMbx.ExchangeGuid
                $Data.PrimarySMTPADdress        = $UserMbx.PrimarySMTPADdress
                $Data.UseDatabaseQuotaDefaults  = $UserMbx.UseDatabaseQuotaDefaults
                $Data.SipAddress                = $UserMbx.emailaddresses | Where-Object {$_ -match "^SIP:"} | ForEach-Object {$_ -replace "sip:"}
                $Data.TotalItemSize             = ($UserMbx | Get-MailboxStatistics -DomainController $psconfig.PDC).totalitemsize.Value  -replace '^.+\((.+\))','$1' -replace '\D' -as [long]
                $Data.LitigationHoldEnabled     = $UserMbx.LitigationHoldEnabled
                $Data.ForwardingAddress         = $UserMbx.ForwardingAddress 
                $Data.SingleItemRecoveryEnabled = $UserMbx.SingleItemRecoveryEnabled
                $Data.FullAccess                = $UserMbx | Get-MailboxPermission -DomainController $psconfig.PDC | Where-Object {
                    ($_.user -notlike "nt authority*") -and ($_.user -notlike "S-1-*") -and ($_.isinherited -eq $false) } | Select-Object -ExpandProperty User

            } else {
                continue
            }

            $Out += $Data
        }
    }

    end{
        Remove-PSSession $Session
        return $Out | Select-Object SamAccountName,PrimarySMTPADdress,DisplayName,ExchangeGuid,ForwardingAddress,FullAccess,LitigationHoldEnabled,SingleItemRecoveryEnabled,UseDatabaseQuotaDefaults,TotalItemSize,Database,SipAddress
    }
 
}