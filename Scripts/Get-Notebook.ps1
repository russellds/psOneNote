function Get-Notebook {
    
    [xml]$xml = $null
    
    $oneNote.GetHierarchy($null, 
        [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages, 
        [ref]$xml)   
    
    Select-Xml -Xml $xml -Namespace $namespace -XPath '/one:Notebooks/one:Notebook/@name' | 
        foreach { $_.Node.Value }
}