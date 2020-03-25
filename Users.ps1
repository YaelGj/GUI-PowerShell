cls
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
Import-Module activedirectory

$ButtonSize = "200,40"

#Création forme principal
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
$Form.BackColor = "GhostWhite"

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
$ButtonImport.ForeColor          = ""
$Form.Controls.Add($ButtonImport)

#Indication réinitialiser
$LabelUsername                   = New-Object system.Windows.Forms.Label
$LabelUsername.text              = "\/ Reset Password \/"
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

#Déclaration emplacement du CSV
$EmplacementCSV = ".\CSV\CSVADUsers.csv"

#Déclaration de la fonction de génération de mot de passe
function PasswordGenerator($SamAccountName)
{
    $Password = -join ((33..125) | Get-Random -Count 10 | % {[char]$_})
    $ViewPassword = $Password
    Set-ADAccountPassword -Identity $SamAccountName -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Reset
  
}

$ButtonRefresh.Add_Click(
{

     Import-Csv -delimiter ";" -Path $EmplacementCSV | Out-GridView –Title Get-CsvData
     
})

$ButtonImport.Add_Click(
{

    #Importer le CSV "CSVADUsers.csv" dans la variable "$Users"
    $Users = Import-csv -delimiter ";" -path $EmplacementCSV

    #Executer ce script pour chaque utilisateurs présent dans le CSV 
    foreach ($User in $Users){

	        #Extraire chaques infos de l"utilisateur et stocker dans les variables correspondantes
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


	#Vérifier si l'utilisateur existe déja dans l'AD
	if (Get-ADUser -Filter {sAMAccountName -eq $SamAccountName}){

	    #Si l"utilisateur existe alors afficher une erreur
        [System.Windows.MessageBox]::Show("Un compte utilisateur se nomant $SamAccountName existe déja dans l'Active Directory.")		 
        Break

	}
	else{
        
        #L'utilisateur n'éxiste pas dans l'AD donc création de l"utilisateur
        
        if ($Country -eq "france"){
        
            $OU = "OU=Users,OU=" + "$City" + ",OU=France,OU=SACRED,DC=sacredlubin,DC=local"
            $Country = "FR"
        }
        elseif ($Country -eq "China"){

            $OU = "OU=Users,OU=" + "$Country" +",OU=SACRED,DC=sacredlubin,DC=local"
            $Country = "CN"

        }
        elseif ($Country -eq "Mexico"){

            $OU = "OU=Users,OU=" + "$Country" +",OU=SACRED,DC=sacredlubin,DC=local"
            $Country = "MX"

        }
        elseif ($Country -eq "Morocco"){

            $OU = "OU=Users,OU=" + "$Country" +",OU=SACRED,DC=sacredlubin,DC=local"
            $Country = "MA"

        }


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
            -Path $OU 
            Add-ADGroupMember -Identity $Work_Group -Members $SamAccountName

            PasswordGenerator($SamAccountName)
            $AllPassword = @{ Login = $SamAccountName; Password = $ViewPassword}

            
        if ($User_Group -eq "FR_USERS"){

            Set-ADUser -Identity $SamAccountName -HomeDirectory "\\slbfs028\fr_users$\$SamAccountName" -HomeDrive H
            Set-ADUser -Identity $SamAccountName -HomeDirectory $null -HomeDrive $null
        }
        elseif ($User_Group -eq "FR_NOMADS"){

            Set-ADUser -Identity $SamAccountName -HomeDirectory "\\slbfs028\fr_nomad$\$SamAccountName" -HomeDrive H
            Set-ADUser -Identity $SamAccountName -HomeDirectory $null -HomeDrive $null
        }
    }
}

$AllPassword | Out-GridView –Title Password

})

$ButtonReset.Add_Click(
{
    $SamAccountName = $TextBoxReset.Text

    if ($SamAccountName.Length -eq 0)
    {
    [System.Windows.MessageBox]::Show("Erreur : il n'y à pas de valeur")
    }
    else
    {    
        if (Get-ADUser -Filter {sAMAccountName -eq $SamAccountName})
        {
        PasswordGenerator($SamAccountName)
        [System.Windows.MessageBox]::Show("Le mot de passe de $SamAccountName est $ViewPassword")
        write-host $ViewPassword
        write-host $Password
	    }
	else
    {
        [System.Windows.MessageBox]::Show("Erreur : impossible de trouver le compte $SamAccountName dans l'Active Directory.")		 
        
    }
    }
    

})

$Form.ShowDialog()