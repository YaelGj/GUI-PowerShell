#Prérequis
cls
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
Import-Module active directory

$ButtonSize = "200,40"

#Création de la forme principale
$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = "240,400"
$Form.text                       = "Form"
$Form.TopMost                    = $false
$Form.AutoSizeMode               = "GrowAndShrink"
$Form.StartPosition              = "CenterScreen"
$Form.FormBorderStyle            = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$Form.MaximizeBox                = $False
$Form.MinimizeBox                = $False
$Form.Text                       = "Gestion Users"
$Form.BackColor                  = "White"

#Logo sacred
$PictureBox                      = New-Object system.Windows.Forms.PictureBox
$PictureBox.width                = 200
$PictureBox.height               = 100
$PictureBox.location             = New-Object System.Drawing.Point(20,20)
$PictureBox.imageLocation        = ".\logo.png"
$PictureBox.SizeMode             = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Form.Controls.Add($PictureBox)

#Bouton rafraichir
$ButtonRefresh                   = New-Object system.Windows.Forms.Button
$ButtonRefresh.text              = "Afficher contenu CSV"
$ButtonRefresh.width             = 200
$ButtonRefresh.height            = 40
$ButtonRefresh.location          = New-Object System.Drawing.Point(20,150)
$ButtonRefresh.Font              = "Calibri,13"
$Form.Controls.Add($ButtonRefresh)

#Bouton importer
$ButtonImport                    = New-Object system.Windows.Forms.Button
$ButtonImport.text               = "Importer CSV"
$ButtonImport.location           = New-Object System.Drawing.Point(20,200)
$ButtonImport.Size               = new-object System.Drawing.Size($ButtonSize)
$ButtonImport.Font               = "Calibri,13"
$Form.Controls.Add($ButtonImport)

#Indication réinitialiser
$LabelUsername                   = New-Object system.Windows.Forms.Label
$LabelUsername.text              = "Entrez login pour reset"
$LabelUsername.AutoSize          = $true
$LabelUsername.location          = New-Object System.Drawing.Point(45,270)
$LabelUsername.Size              = new-object System.Drawing.Size($ButtonSize)
$LabelUsername.Font              = "Calibri,13"
$Form.Controls.Add($LabelUsername)

#Saisie réinitialiser
$TextBoxReset                    = New-Object system.Windows.Forms.TextBox
$TextBoxReset.multiline          = $false
$TextBoxReset.width              = 200
$TextBoxReset.height             = 40
$TextBoxReset.location           = New-Object System.Drawing.Point(20,300)
$TextBoxReset.Font               = "Calibri,13"
$Form.Controls.Add($TextBoxReset)

#Bouton réinitialiser
$ButtonReset                     = New-Object system.Windows.Forms.Button
$ButtonReset.text                = "Reset mot de passe"
$ButtonReset.width               = 200
$ButtonReset.height              = 40
$ButtonReset.location            = New-Object System.Drawing.Point(20,350)
$ButtonReset.Font                = "Calibri,13"
$Form.Controls.Add($ButtonReset)

#Déclaration variable
$EmplacementCSV = ".\CSV\CSVADUsers.csv"

#Déclaration de la fonction de génération de mot de passe
function PasswordGenerator($SamAccountName)
{
    #tirer 10 nombres aléatoire dans la range correspondant dans la table ASCII au majuscule, minuscule, caractère spéciaux et chiffre
    $Password = -join ((33..125) | Get-Random -Count 10 | % {[char]$_})
   
    #Définition du mot de passe ainsi générer
    Set-ADAccountPassword -Identity $SamAccountName -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Reset
    [System.Windows.MessageBox]::Show("Le mot de passe de $SamAccountName est $Password")
}


$ButtonRefresh.Add_Click(
{
    #Importer le CSV puis afficher son contenu dans un tableau popup
    Import-Csv -delimiter ";" -Path $EmplacementCSV | Out-GridView –Title Get-CsvData
})

$ButtonImport.Add_Click(
{
    #Importer le CSV "CSVADUsers.csv" dans la variable "$Users"
    $Users = Import-csv -delimiter ";" -path $EmplacementCSV

    #Executer ce script pour chaque utilisateur présent dans le CSV 
    foreach ($User in $Users)
    {
	    #Extraire chaque info de l'utilisateur et stocker dans les variables 
   Correspondantes
            $SamAccountName    = $User.User
            $Surname           = $User.Nom    
            $GivenName         = $User.Prenom
            $EmailAddress      = $User.Email
            $User_Group        = $User.Groupe_User
            $Work_Group        = $User.Groupe_Metier
            $Title             = $User.Metier
            $Description       = $User.Description
            $Country           = $User.Pays
            $PostalCode        = $User.CodePostal
            $City              = $User.Ville
            $StreetAddress     = $User.Adresse
            $Department        = $User.Service
            $HomePhone         = $User.teldomicile
            $MobilePhone       = $User.telmobile
            $OfficePhone       = $User.TelBureau
            $HomePage          = $User.SiteWeb


	#Vérifier si l'utilisateur existe déjà dans l'AD
	if (Get-ADUser -Filter {sAMAccountName -eq $SamAccountName})
    {
	    #Si l’utilisateur existe alors afficher une erreur
        [System.Windows.MessageBox]::Show("Un compte utilisateur se nommant $SamAccountName existe déjà dans l'Active Directory.")		 
        Break

	}
	else{
        #L'utilisateur n'existe pas dans l'AD donc création de l’utilisateur
        if ($Country -eq "france")
        {
            #Si le pays est "france" alors placer l'utilisateur dans l'OU correspondant à sa 
 	        #ville dans l'OU France et définir le bon code pays
            $OU = "OU=Users,OU=" + "$City" + ",OU=France,OU=SACRED,DC=sacredlubin,DC=local"
            $Country = "FR"
        }
        elseif ($Country -eq "China")
        {
            #Sinon si le pays est "China" alors placer l'utilisateur dans l'OU correspondant à son pays et définir le bon code pays
            $OU = "OU=Users,OU=" + "$Country" +",OU=SACRED,DC=sacredlubin,DC=local"
            $Country = "CN"
        }
        elseif ($Country -eq "Mexico")
        {
            #Sinon si le pays est "Mexico" alors placer l'utilisateur dans l'OU correspondant à son pays et définir le bon code pays
            $OU = "OU=Users,OU=" + "$Country" +",OU=SACRED,DC=sacredlubin,DC=local"
            $Country = "MX"
        }
        elseif ($Country -eq "Morocco")
        {
   #Sinon si le pays est "Morocco" alors placer l'utilisateur dans l'OU correspondant à son pays et définir le bon code pays
            $OU = "OU=Users,OU=" + "$Country" +",OU=SACRED,DC=sacredlubin,DC=local"
            $Country = "MA"
        }
            #Créer l'utilisateur avec les variables correspondantes
    New-ADUser `
            -Name "$GivenName $Surname" `
            -SamAccountName $SamAccountName `
            -Surname $Surname `
            -GivenName $GivenName `
            -EmailAddress $EmailAddress `
            -UserPrincipalName $EmailAddress `
            -Title $Title `
            -Description $Description `
            -Country $Country `
            -PostalCode $PostalCode `
            -City $City `
            -StreetAddress $StreetAddress `
            -Department $Department `
            -HomePhone $HomePhone `
            -MobilePhone $MobilePhone `
            -OfficePhone $OfficePhone `
            -HomePage $HomePage `
            -DisplayName "$Surname" `
            -Office $Ville `
            -Path $OU `
            -Company "Sacred"
            #Ajouter l'utilisateur au groupe correspondant à sont métier
            Add-ADGroupMember -Identity $Work_Group -Members $SamAccountName
            #Générer le mot de passe de l'utilisateur
            PasswordGenerator($SamAccountName)

        #Si l'utilisateur fait partie du groupe "FR_USERS" ou "FR_NOMADS" alors le définir pour générer sont dossier personnel
        # dans le dossier correspondant puis enlever la définition car ce paramètre est géré par GPO dans l'entreprise
        if ($User_Group -eq "FR_USERS")
        {
            Set-ADUser -Identity $SamAccountName -HomeDirectory "\\slbfs028\fr_users$\$SamAccountName" -HomeDrive H
            Set-ADUser -Identity $SamAccountName -HomeDirectory $null -HomeDrive $null
        }
        elseif ($User_Group -eq "FR_NOMADS")
        {
            Set-ADUser -Identity $SamAccountName -HomeDirectory "\\slbfs028\fr_nomad$\$SamAccountName" -HomeDrive H
            Set-ADUser -Identity $SamAccountName -HomeDirectory $null -HomeDrive $null
        }
    }
}


})

$ButtonReset.Add_Click(
{
    #Récupérer le contenu du textbox
    $SamAccountName = $TextBoxReset.Text
    #Si le textbox était vide alors afficher un message d'erreur
    if ($SamAccountName.Length -eq 0)
    {
    [System.Windows.MessageBox]::Show("Erreur : il n'y a pas de valeur")
    }
    else
    {   
	#Si l’utilisateur existe alors réinitialiser sont mot de passe sinon afficher un message d’erreur
        if (Get-ADUser -Filter {sAMAccountName -eq $SamAccountName})
        {
        PasswordGenerator($SamAccountName)

	    }
	else
    {
        [System.Windows.MessageBox]::Show("Erreur : impossible de trouver le compte $SamAccountName dans l'Active Directory.")		 
    }
    }
})
$Form.ShowDialog()
