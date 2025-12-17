
Get-AzResourceProvider -ProviderNamespace Microsoft.Network | 
  Where-Object { $_.ResourceTypes.ResourceTypeName -eq "virtualNetworks" } | 
  Select-Object -ExpandProperty Locations
