cls
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
Import-Module activedirectory

$ButtonSize = "200,40"

#forme principal
$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = "240,400"
$Form.text                       = "Form"
$Form.TopMost                    = $false
$Form.AutoSizeMode               = "GrowAndShrink"
$Form.StartPosition              = "CenterScreen"
$Form.FormBorderStyle            = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$Form.MaximizeBox                = $False
$Form.MinimizeBox                = $False
$Form.Text                       = "Gestion Dossier"
$Form.BackColor                  = "White"

#logo sacred
$PictureBox                      = New-Object system.Windows.Forms.PictureBox
$PictureBox.width                = 200
$PictureBox.height               = 100
$PictureBox.location             = New-Object System.Drawing.Point(20,20)
$PictureBox.imageLocation        = ".\logo.png"
$PictureBox.SizeMode             = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Form.Controls.Add($PictureBox)

$LabelDir                        = New-Object system.Windows.Forms.Label
$LabelDir.text                   = "Emplacement"
$LabelDir.location               = New-Object System.Drawing.Point(70,130)
$LabelDir.Size                   = new-object System.Drawing.Size(200,20)
$LabelDir.Font                   = "Calibri,13"
$Form.Controls.Add($LabelDir)

$TextBoxPath                     = New-Object system.Windows.Forms.TextBox
$TextBoxPath.multiline           = $false
$TextBoxPath.location            = New-Object System.Drawing.Point(20,150)
$TextBoxPath.Size                = new-object System.Drawing.Size($ButtonSize)
$TextBoxPath.Font                = "Calibri,13"
$TextBoxPath.Text                = "\\SLBFS028\FR_METIERS"
$Form.Controls.Add($TextBoxPath)

$LabelDir                        = New-Object system.Windows.Forms.Label
$LabelDir.text                   = "Nom Dossier"
$LabelDir.location               = New-Object System.Drawing.Point(70,180)
$LabelDir.Size                   = new-object System.Drawing.Size(200,20)
$LabelDir.Font                   = "Calibri,13"
$Form.Controls.Add($LabelDir)

$TextBoxDir                      = New-Object system.Windows.Forms.TextBox
$TextBoxDir.multiline            = $false
$TextBoxDir.location             = New-Object System.Drawing.Point(20,200)
$TextBoxDir.Size                 = new-object System.Drawing.Size($ButtonSize)
$TextBoxDir.Font                 = "Calibri,13"
$TextBoxDir.Text                 = "Dossier"
$Form.Controls.Add($TextBoxDir)

$LabelGroup                      = New-Object system.Windows.Forms.Label
$LabelGroup.text                 = "Nom Groupe"
$LabelGroup.location             = New-Object System.Drawing.Point(70,230)
$LabelGroup.Size                 = new-object System.Drawing.Size(200,20)
$LabelGroup.Font                 = "Calibri,13"
$Form.Controls.Add($LabelGroup)

$TextBoxGroup                    = New-Object system.Windows.Forms.TextBox
$TextBoxGroup.multiline          = $false
$TextBoxGroup.location           = New-Object System.Drawing.Point(20,250)
$TextBoxGroup.Size               = new-object System.Drawing.Size($ButtonSize)
$TextBoxGroup.Font               = "Calibri,13"
$TextBoxGroup.Text               = "SAC_FR"
$Form.Controls.Add($TextBoxGroup)

$ButtonVisu                      = New-Object system.Windows.Forms.Button
$ButtonVisu.text                 = "Aperçu emplacement"
$ButtonVisu.location             = New-Object System.Drawing.Point(20,300)
$ButtonVisu.Size                 = new-object System.Drawing.Size($ButtonSize)
$ButtonVisu.Font                 = "Calibri,13"
$Form.Controls.Add($ButtonVisu)

$ButtonCreate                    = New-Object system.Windows.Forms.Button
$ButtonCreate.text               = "Executer"
$ButtonCreate.location           = New-Object System.Drawing.Point(20,350)
$ButtonCreate.Size               = new-object System.Drawing.Size($ButtonSize)
$ButtonCreate.Font               = "Calibri,13"
$Form.Controls.Add($ButtonCreate)

$ButtonVisu.Add_Click(
{
     $Path    = $TextBoxPath.Text
     Get-ChildItem -Path "$Path" | Select -Property name,lastwritetime |Out-GridView –Title aperçus

})

$ButtonCreate.Add_Click(
{
    $Path     = $TextBoxPath.Text
    $Group    = $TextBoxGroup.Text
    $Dir      = $TextBoxDir.Text
    

    if(($Path.Length -ne 0) -or ($Group.Length -ne 0) -or ($Dir.Length -ne 0))
    { 
    
        if((Get-ADGroup -LDAPFilter "(SAMAccountName=$Group)") -eq $null)
        {
        New-ADGroup -Name "$Group" -GroupScope "DomainLocal" -GroupCategory "Security" -DisplayName "$Group" `
        -Path "OU=Groups METIERS,OU=_Groups,OU=France,OU=SACRED,DC=sacredlubin,DC=local"
        }
        if($Group -Match "READ")
        {
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule @("sacredlubin\$Group", "ReadAndExecute", "None" , "none", "Allow")
        }
        else
        {
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule @("sacredlubin\$Group", "Modify", "containerinherit", "none", "Allow")
        }

        if((Test-Path "$Path\$Dir") -ne $true)
        {
        New-Item -Name "$Dir" -ItemType directory -Path "$path"
 
        }
        $Path = "$Path\$Dir"
        $Acl = Get-Acl $Path     
        $AccessRule
        $Acl.AddAccessRule($accessRule)
        $Acl | Set-Acl $Path
        [System.Windows.MessageBox]::Show("Succès")
   }
   else
   {
       [System.Windows.MessageBox]::Show("Valeurs manquantes")
   }
   
}
)

$Form.ShowDialog()