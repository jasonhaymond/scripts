Get-AppxProvisionedPackage -Online | Where-Object -FilterScript {$_.DisplayName -eq "Microsoft.Office.Desktop"} | Remove-AppxProvisionedPackage -Online
Get-AppxPackage -Name Microsoft.Office.Desktop -AllUsers | Remove-AppxPackage -AllUsers
