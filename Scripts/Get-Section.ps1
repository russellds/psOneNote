function Get-Section {
    
    [xml]$xml = $null
    
    $oneNote.GetHierarchy($null, 
        [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages, 
        [ref]$xml)   
    
    $xml.Notebooks.Notebook.Section |
        Select-Object -ExpandProperty name |
        where { $_ -ne 'OneNote_RecycleBin' } |
        Sort-Object
}