# This script creates the Office 365 config for DOQEX
# Author: Roel van den Bussche (roel@moreteq.com)
# Use at own risk

# Connect to Office 365 console
Set-ExecutionPolicy RemoteSigned -Force
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

# Let's set some variables:
$dq_instance = Read-Host -Prompt "Enter DOQEX instance FQDN"
$dq_node_ip = Read-Host -Prompt "Enter DOQEX IP address"

# Add Inbound Connector - Accept processed e-mail from DOQEX Instance
New-InboundConnector -Name “DOQEX to Office 365 - by moreteq” -Comment “Accept e-mail send from DOQEX instance” -Enabled $true -ConnectorType OnPremises -SenderDomains * -RestrictDomainsToIPAddresses $true -RequireTls $true -SenderIPAddresses $dq_node_ip

# Add Outbound Connector - Send e-mail to DOQEX for processing
New-OutboundConnector -name "Office 365 to DOQEX2 - by moreteq" -Comment "Send e-mail to DOQEX for processing" -Enabled $true -ConnectorType Partner -IsTransportRuleScoped $true -UseMXRecord $false -SmartHosts $dq_instance -TlsSettings EncryptionOnly

# Add the transport rule
New-TransportRule -Name "DOQEX Message Transport - by moreteq" -Comments "Send outbound e-mail via DOQEX for compliancy" -Enabled $false -FromScope InOrganization -SentToScope NotInOrganization -ExceptIfHeaderContainsMessageHeader X-DOQEX-NOTIFY -ExceptIfHeaderContainsWords DOQEX -ExceptIfHeaderMatchesMessageHeader X-DOQEX-LMID -ExceptIfHeaderMatchesPatterns "@" -RouteMessageOutboundConnector "Office 365 to DOQEX2 - by moreteq"