Function Get-ComStatus
{
    Function Test-ComPlus
    {
        try
        {
            $comCatalog = New-Object -ComObject COMAdmin.COMAdminCatalog
            $appColl = $comCatalog.GetCollection("Applications")
            $appColl.Populate()
            return $appColl.Count
        } catch {
            return $Error[0].Exception.Message
        }
    }
    
    $COMHealth = Test-ComPlus
    
    if ($COMHealth -is [int])
    {
        Write-Output "COM+ is responding. Found: $COMHealth COM+ applications."
    } else {
        Write-Output "Unable to query COM+. Error: $COMHealth"
    }
}

Get-ComStatus
