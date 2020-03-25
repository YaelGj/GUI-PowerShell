$Path = "C:\Users\ygrosjean\Desktop\powe\testdir"
$acl = Get-Acl $Path
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("FR_IT_SACRED",,"Write","ContainerInherit,ObjectInherit","InheritOnly","Allow")
$acl.AddAccessRule($accessRule)
$acl | Set-Acl $Path