# This script adds the following to your Office 365 environment:
# - Outbound connector
# - Inbound Connector
# - Transport rule disabled by default
#
# After running the script, you will have to manually check and activate the transport rule.
#
# Web: https://sendteq.com
# Contact: support@sendteq.com
# Use at own risk!

# Connect to Office 365 console
Set-ExecutionPolicy RemoteSigned -Force
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

# Let's set some variables:
$Instance = Read-Host -Prompt "Enter SENDTEQ instance FQDN"
$Instance_IP = [System.Net.Dns]::GetHostAddresses($Instance) | foreach {echo $_.IPAddressToString }

# Add Inbound Connector - Accept processed e-mail from DOQEX Instance
New-InboundConnector -Name “SENDTEQ to Office 365” -Comment “Accept e-mail send back from SENDTEQ instance - Added by SENDTEQ” -Enabled $true -ConnectorType OnPremises -SenderDomains * -RestrictDomainsToIPAddresses $true -RequireTls $true -SenderIPAddresses $Instance_IP

# Add Outbound Connector - Send e-mail to DOQEX for processing
New-OutboundConnector -name "Office 365 to SENDTEQ" -Comment "Send e-mail to SENDTEQ for processing - Added by SENDTEQ" -Enabled $true -ConnectorType Partner -IsTransportRuleScoped $true -UseMXRecord $false -SmartHosts $Instance -TlsSettings EncryptionOnly

# Add the transport rule
New-TransportRule -Name "SENDTEQ Message Transport" -Comments "Send outbound e-mail to SENDTEQ for compliancy - Added by SENDTEQ" -Enabled $false -FromScope InOrganization -SentToScope NotInOrganization -ExceptIfHeaderContainsMessageHeader X-DOQEX-NOTIFY -ExceptIfHeaderContainsWords DOQEX -ExceptIfHeaderMatchesMessageHeader X-DOQEX-LMID -ExceptIfHeaderMatchesPatterns "@" -RouteMessageOutboundConnector "Office 365 to SENDTEQ"