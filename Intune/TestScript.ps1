Try {
    $folderContent = Get-ChildItem - Path 'C:\IDontExist' -ErrorAction SilentlyContinue
}
Catch {
    #Dont tell a soul
}
If (-not $folderContent) {
    Try {
        New-Item -Path 'C:\IDontExist' -ItemType 'Directory' -Force
    }
    Catch {
        #I failed miserably
    }
}   



$newShell = New-Object -ComObject Shell.Application
$newShell.open("intunemanagementextension://syncapp")
