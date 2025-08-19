Add-Type -AssemblyName System.Drawing

$sourcePath = "C:\Users\TsubotaniTA\VSCODE\RemoteWakeConnect\sample\11_16_04.png"
$destPath256 = "C:\Users\TsubotaniTA\VSCODE\RemoteWakeConnect\src\app_icon.png"
$destPathIco = "C:\Users\TsubotaniTA\VSCODE\RemoteWakeConnect\src\app.ico"

# Load the source image
$sourceImage = [System.Drawing.Image]::FromFile($sourcePath)

# Create 256x256 PNG with transparency
$resized256 = New-Object System.Drawing.Bitmap(256, 256, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$graphics256 = [System.Drawing.Graphics]::FromImage($resized256)
$graphics256.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$graphics256.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$graphics256.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
$graphics256.Clear([System.Drawing.Color]::Transparent)
$graphics256.DrawImage($sourceImage, 0, 0, 256, 256)
$resized256.Save($destPath256, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics256.Dispose()
$resized256.Dispose()

# Create a proper ICO file using Icon.FromHandle
$icon32 = New-Object System.Drawing.Bitmap(32, 32, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$graphics32 = [System.Drawing.Graphics]::FromImage($icon32)
$graphics32.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$graphics32.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$graphics32.Clear([System.Drawing.Color]::Transparent)
$graphics32.DrawImage($sourceImage, 0, 0, 32, 32)
$graphics32.Dispose()

# Convert bitmap to icon
$hIcon = $icon32.GetHicon()
$icon = [System.Drawing.Icon]::FromHandle($hIcon)

# Save the icon
$fs = [System.IO.File]::Create($destPathIco)
$icon.Save($fs)
$fs.Close()

# Clean up
[System.Runtime.InteropServices.Marshal]::DestroyIcon($hIcon)
$icon.Dispose()
$icon32.Dispose()
$sourceImage.Dispose()

Write-Host "Icons created successfully!"
Write-Host "PNG (256x256) saved to: $destPath256"
Write-Host "ICO saved to: $destPathIco"