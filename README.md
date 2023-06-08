# UpdateADPhotos
Updates Active Directory Thumbnail Photos

If you'd like to insert your own logo, from Terminal/PowerShell:
```
[convert]::ToBase64String((Get-Content '<PATH TO IMAGE>' -Encoding Byte))
```

Uncomment this line and paste that output:
```
$Global:Logo = [Convert]::FromBase64String("<OUTPUT HERE>")
```

Uncomment 
Line 133:
```
#### Load Logo Onto Form
Load-Logo
```
