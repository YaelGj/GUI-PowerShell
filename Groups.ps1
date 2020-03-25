cls
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
Import-Module activedirectory

$ButtonSize = "200,40"

#forme principal
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
$Form.Text                       = "Gestion Users"
$Form.BackColor = "GhostWhite"

$PictureBox                      = New-Object system.Windows.Forms.PictureBox
$PictureBox.width                = 200
$PictureBox.height               = 100
$PictureBox.location             = New-Object System.Drawing.Point(20,20)
$PictureBox.imageLocation        = ".\logo.png"
$PictureBox.SizeMode             = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Form.Controls.Add($PictureBox)

$TextBoxFile                     = New-Object system.Windows.Forms.TextBox
$TextBoxFile.multiline           = $false
$TextBoxFile.location            = New-Object System.Drawing.Point(20,290)
$TextBoxFile.Size                = new-object System.Drawing.Size($ButtonSize)
$TextBoxFile.Font                = "Calibri,13"
$Form.Controls.Add($TextBoxFile)

$ButtonRefresh                   = New-Object system.Windows.Forms.Button
$ButtonRefresh.text              = "Afficher contenu CSV"
$ButtonRefresh.location          = New-Object System.Drawing.Point(20,150)
$ButtonRefresh.Size              = new-object System.Drawing.Size($ButtonSize)
$ButtonRefresh.Font              = "Calibri,13"
$Form.Controls.Add($ButtonRefresh)

$ButtonImport                    = New-Object system.Windows.Forms.Button
$ButtonImport.text               = "Importer CSV"
$ButtonImport.location           = New-Object System.Drawing.Point(20,200)
$ButtonImport.Size               = new-object System.Drawing.Size($ButtonSize)
$ButtonImport.Font               = "Calibri,13"
$ButtonImport.ForeColor          = ""
$Form.Controls.Add($ButtonImport)

$LabelUsername                   = New-Object system.Windows.Forms.Label
$LabelUsername.text              = "\/ Reset Password \/"
$LabelUsername.AutoSize          = $true
$LabelUsername.location          = New-Object System.Drawing.Point(45,270)
$LabelUsername.Size              = new-object System.Drawing.Size($ButtonSize)
$LabelUsername.Font              = "Calibri,13"
$Form.Controls.Add($LabelUsername)


$ButtonReset                     = New-Object system.Windows.Forms.Button
$ButtonReset.text                = "Reset mot de passe"
$ButtonReset.location            = New-Object System.Drawing.Point(20,330)
$ButtonReset.Size                = new-object System.Drawing.Size($ButtonSize)
$ButtonReset.Font                = "Calibri,13"
$Form.Controls.Add($ButtonReset)

$Form.ShowDialog()