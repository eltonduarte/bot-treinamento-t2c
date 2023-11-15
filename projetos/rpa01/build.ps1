$exclude = @("venv", "rpa01.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "rpa01.zip" -Force