[Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote") | Out-Null

$oneNote = New-Object -ComObject OneNote.Application

$namespace = @{
    one = 'http://schemas.microsoft.com/office/onenote/2010/onenote'
}

$scripts = Get-ChildItem $(Join-Path -Path $psScriptRoot -ChildPath "Scripts") -Recurse

foreach( $script in $scripts) {
	if(-not $script.PSIsContainer) {
    	. $script.Fullname
	}
}
