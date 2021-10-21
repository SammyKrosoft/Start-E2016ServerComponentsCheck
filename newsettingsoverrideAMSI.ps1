New-SettingOverride -Name “DisablingAMSIScan” -Component Cafe -Section HttpRequestFiltering -Parameters (“Enabled=False”) -Reason “Testing”

Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh

Restart-Service -Name W3SVC, WAS -Force




