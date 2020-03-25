$PersonnelPathRoot = "C:\Partage\Personnel"
$UsersList = ("utilisateur01","utilisateur02")

foreach($User in $UsersList){

    # Creation du dossier personnel
    New-Item -ItemType Directory -Path "$PersonnelPathRoot\$User"

    # Desactiver l'heritage tout en copiant les autorisations NTFS héritées
    Get-Item "$PersonnelPathRoot\$User" | Disable-NTFSAccessInheritance

    # Ajout des autorisations NTFS
    Add-NTFSAccess –Path "$PersonnelPathRoot\$User"  –Account "$User@it-connect.local" –AccessRights FullControl
    
    # Modifier le proprietaire sur le dossier
    Set-NTFSOwner -Path "$PersonnelPathRoot\$User" -Account "$User@it-connect.local"
    
    # Supprimer des autorisations NTFS
    Remove-NTFSAccess –Path "$PersonnelPathRoot\$User"  –Account "Utilisateurs" -AccessRights FullControl

} # foreach($User in $UsersList)