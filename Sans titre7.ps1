#Vidage de la console, chargement des pré-requis
cls
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
Import-Module activedirectory

#definition de la taille de tous les boutons
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
$Form.Text                       = "Menu"
$Form.BackColor                  = "White"

#logo sacred
$PictureBox                      = New-Object system.Windows.Forms.PictureBox
$PictureBox.location             = New-Object System.Drawing.Point(20,20)
$PictureBox.Size                 = new-object System.Drawing.Size(200,100)
$PictureBox.imageLocation        = ".\logo.png"
$PictureBox.SizeMode             = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Form.Controls.Add($PictureBox)

#Bouton script Gestion Users
$ButtonUsers                     = New-Object system.Windows.Forms.Button
$ButtonUsers.text                = "Gestion users"
$ButtonUsers.location            = New-Object System.Drawing.Point(20,150)
$ButtonUsers.Size                = new-object System.Drawing.Size($ButtonSize)
$ButtonUsers.Font                = "Calibri,13"
$ButtonUsers.Add_Click({& ".\Users.ps1" })
$Form.Controls.Add($ButtonUsers)

#bouton script Gestion RDP
$ButtonRDP                       = New-Object system.Windows.Forms.Button
$ButtonRDP.text                  = "Nettoyer Sessions RDP"
$ButtonRDP.location              = New-Object System.Drawing.Point(20,200)
$ButtonRDP.Size                  = new-object System.Drawing.Size($ButtonSize)
$ButtonRDP.Font                  = "Calibri,13"
$Form.Controls.Add($ButtonRDP)

#bouton script Gestion dossier
$ButtonDirs                      = New-Object system.Windows.Forms.Button
$ButtonDirs.text                 = "Gestion Dossiers"
$ButtonDirs.location             = New-Object System.Drawing.Point(20,250)
$ButtonDirs.Size                 = new-object System.Drawing.Size($ButtonSize)
$ButtonDirs.Font                 = "Calibri,13"
$ButtonDirs.Add_Click({& ".\Dirs.ps1" })
$Form.Controls.Add($ButtonDirs)

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.text                    = "Button"
$Button4.location                = New-Object System.Drawing.Point(20,300)
$Button4.Size                    = new-object System.Drawing.Size($ButtonSize)
$Button4.Font                    = "Calibri,13"
$Form.Controls.Add($Button4)


$ButtonRDP.Add_Click(
{
    
    $Server1   = "slbts040"
    $Server2   = "slbts041"
    $Session = qwinsta /server:$Server1 | select-string -notmatch "Actif" | select-string -notmatch "services"

    if ($Session)
    {
        $Session| % { 
        logoff ($_.tostring() -split ' +')[2] /Server:$Server1
        }
    }

    $Session = qwinsta /server:$Server2 | select-string "D" | select-string -notmatch "services"
    if ($Session)
    {
        $Session| % { 
        logoff ($_.tostring() -split ' +')[2] /Server:$Server2
        }
    }
    [System.Windows.MessageBox]::Show("Sessions fermées")
}
)

$Form.ShowDialog()