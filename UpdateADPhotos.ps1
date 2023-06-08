Add-Type -assembly System.Windows.Forms

#### Define Global Variables
$Global:Users = $null;
$Global:Bytes = $null;
$Global:Login = $null;
$Global:lbl_StatusMsg = New-Object System.Windows.Forms.Label
$Global:cbx_UserSelection = New-Object System.Windows.Forms.ComboBox
$Global:UArray = @()
$Global:NIF = [Convert]::FromBase64String("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")
#$Global:Logo = [Convert]::FromBase64String("")

#### Define Functions
function Load-Logo {
    $img_Logo = New-Object System.Windows.Forms.PictureBox
    $img_Logo.Location = New-Object System.Drawing.Size(27,10)
    $ms = New-Object System.IO.MemoryStream(,$Logo)
    $Banner = [System.Drawing.Image]::FromStream($ms)
    $img_Logo.Image = $Banner
    $img_Logo.AutoSize = $true
    $main_form.Controls.Add($img_Logo)
}

function Get-Users() {
    $Global:UArray = @()
    $Users = get-aduser -filter * -Properties sAMAccountName,telephoneNumber,givenName,sn,userAccountControl,thumbnailPhoto
    Foreach ($User in $Users)
    {
        if ($User.telephoneNumber -and ($User.userAccountControl -eq "512")) {
            $Global:UArray += [PSCustomObject]@{
                Name = $User.sn+", "+$User.givenName
                Login = $User.sAMAccountName
                Photo = $User.thumbnailPhoto
            }
        }
    }
    $Global:UArray = $Global:UArray | Sort-Object -Property Name
}

function Populate-Dropdown() {
    Get-Users
    $Global:UArray | ForEach-Object {[void] $cbx_UserSelection.Items.Add($_.Name)}
}

function Open-File(){
    $main_form.Controls.Remove($img_NewPhoto)
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter = 'Pictures (*.jpg;*.png;*.bmp)|*.jpg;*.png;*.bmp'
    }
    $FileBrowser.ShowDialog() |  Out-Null
    $UName = ($UArray |? {$_.Name -eq $cbx_UserSelection.Text}).Login
    $NewImage = Resize-Image $FileBrowser.FileName $UName
    $img_NewPhoto.Image = $NewImage
    $img_NewPhoto.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::CenterImage
    $main_form.Controls.Add($img_NewPhoto)
    $btn_Submit.Enabled = $true
} 

function Resize-Image ([string]$SelectedImage, [string]$UName) {
    $NewPhoto = [System.Drawing.Image]::FromFile($SelectedImage)
    $dimensions = @((96 / $NewPhoto.Width), (96 / $NewPhoto.Height))
    $ratio = $dimensions | Sort-Object -Descending | Select -Last 1
    [int]$Width = $NewPhoto.Width * $ratio
    [int]$Height = $NewPhoto.Height * $ratio
    $Bitmap = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Width, $Height
    $NewImage = [System.Drawing.Graphics]::FromImage($Bitmap)
    $NewImage.DrawImage($NewPhoto, $(New-Object -TypeName System.Drawing.Rectangle -ArgumentList 0, 0, $Width, $Height))
    $memory = New-Object System.IO.MemoryStream
    $Bitmap.Save($memory, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $Global:Bytes = $memory.ToArray()
    $memory.Flush()
    $memory = New-Object System.IO.MemoryStream(,$Bytes)

    return [System.Drawing.Image]::FromStream($memory)
}

function Update-Photo([string] $operation){
    Try {
        if ($operation -eq "Replace") {
            Set-ADUser $Login -Replace @{thumbnailPhoto=$Bytes}
            Get-Message "Successfully replaced photo." "Success"
        } else {
            Set-ADUser $Login -Clear thumbnailPhoto
            Get-Message "Successfully cleared photo." "Success"
        }
    } Catch {
        Get-Message $Error[0] "Error"
    }
    Get-Users
    Get-Thumbnail
}

function Get-Thumbnail(){
    $main_form.Controls.Remove($img_NewPhoto)
    $img_NewPhoto.Image = $NewImage
    $main_form.Controls.Add($img_NewPhoto)
    $Photo = ($Global:UArray |? {$_.Name -eq $cbx_UserSelection.Text}).Photo
    if ($Photo) {
        $ms = New-Object System.IO.MemoryStream(,$Photo)
        $btn_Clear.Enabled = $true
    } else {
        $ms = New-Object System.IO.MemoryStream(,$NIF)
        $btn_Clear.Enabled = $false
    }
    $CurrentPhoto = [System.Drawing.Image]::FromStream($ms)
    $main_form.Controls.Remove($img_CurrentPhoto)
    $img_CurrentPhoto.Image = $CurrentPhoto
    $img_CurrentPhoto.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::CenterImage
    $main_form.Controls.Add($img_CurrentPhoto)
    $btn_GetPic.Enabled = $true
}

function Get-Message([string] $Msg, [string] $Type) {
    $main_form.Controls.Remove($lbl_StatusMsg)
    if ($Type -eq "Error") {
        $lbl_StatusMsg.ForeColor = "Red"
    } else {
        $lbl_StatusMsg.ForeColor = "Black"
    }
    $lbl_StatusMsg.Text = $Msg
    $main_form.Controls.Add($lbl_StatusMsg)
}

#### Set Basic Form Attributes
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Update Active Directory Thumbnail Photo'
$main_form.Width = 600
$main_form.Height = 400
$main_form.AutoSize = $false

#### Load Bank Logo Onto Form
#Load-Logo

#### Define all text based fields
$lbl_UserSelect = New-Object System.Windows.Forms.Label
$lbl_UserSelect.Text = "User:"
$lbl_UserSelect.Location  = New-Object System.Drawing.Point(85,80)
$lbl_UserSelect.AutoSize = $true
$lbl_CurrentPhoto = New-Object System.Windows.Forms.Label
$lbl_CurrentPhoto.Text = "Current Photo"
$lbl_CurrentPhoto.Location  = New-Object System.Drawing.Point(160,120)
$lbl_CurrentPhoto.AutoSize = $true
$lbl_NewPhoto = New-Object System.Windows.Forms.Label
$lbl_NewPhoto.Text = "New Photo"
$lbl_NewPhoto.Location  = New-Object System.Drawing.Point(378,120)
$lbl_NewPhoto.AutoSize = $true
$lbl_StatusMsg.Location  = New-Object System.Drawing.Point(0,320)
$lbl_StatusMsg.Width = 600
$lbl_StatusMsg.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

#Create dropdown menu for User selection
$cbx_UserSelection.Width = 150
$cbx_UserSelection.Location  = New-Object System.Drawing.Point(125,80)
Populate-Dropdown

#### Define all buttons
$btn_GetPic = New-Object System.Windows.Forms.Button
$btn_GetPic.Location = New-Object System.Drawing.Point(369,80)
$btn_GetPic.AutoSize = $true
$btn_GetPic.Text = "Select"
$btn_GetPic.Enabled = $false
$btn_GetPic.Add_Click({Open-File})
$btn_Submit = New-Object System.Windows.Forms.Button
$btn_Submit.Location = New-Object System.Drawing.Point(369,280)
$btn_Submit.AutoSize = $true
$btn_Submit.Text = "Replace"
$btn_Submit.Enabled = $false
$btn_Submit.Add_Click({Update-Photo "Replace"})
$btn_Clear = New-Object System.Windows.Forms.Button
$btn_Clear.Location = New-Object System.Drawing.Point(158,280)
$btn_Clear.AutoSize = $true
$btn_Clear.Text = "Remove"
$btn_Clear.Enabled = $false
$btn_Clear.Add_Click({Update-Photo "Clear"})

#### Create Photo Field Boxes
$img_CurrentPhoto = New-Object System.Windows.Forms.PictureBox
$img_CurrentPhoto.Location = New-Object System.Drawing.Size(147,160)
$img_CurrentPhoto.BorderStyle = "FixedSingle"
$img_CurrentPhoto.Width = 96
$img_CurrentPhoto.Height = 96
$img_NewPhoto = New-Object System.Windows.Forms.PictureBox
$img_NewPhoto.Location = New-Object System.Drawing.Size(358,160)
$img_NewPhoto.BorderStyle = "FixedSingle"
$img_NewPhoto.Width = 96
$img_NewPhoto.Height = 96

#### Adding all fields to form
$main_form.Controls.Add($lbl_UserSelect)
$main_form.Controls.Add($lbl_CurrentPhoto)
$main_form.Controls.Add($lbl_NewPhoto)
$main_form.Controls.Add($btn_GetPic)
$main_form.Controls.Add($btn_Submit)
$main_form.Controls.Add($btn_Clear)
$main_form.Controls.Add($cbx_UserSelection)
$main_form.Controls.Add($img_CurrentPhoto)
$main_form.Controls.Add($img_NewPhoto)

#### Monitor when the Dropdown selection is changed
$cbx_UserSelection_SelectedIndexChanged=
{
    $Global:Login = ($UArray |? {$_.Name -eq $cbx_UserSelection.Text}).Login
    Get-Thumbnail
}
$cbx_UserSelection.add_SelectedIndexChanged($cbx_UserSelection_SelectedIndexChanged)

#### Show form
$main_form.ShowDialog()
