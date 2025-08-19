Add-Type -AssemblyName System.Drawing

$sourcePath = "C:\Users\TsubotaniTA\VSCODE\RemoteWakeConnect\src\app_icon.png"
$destPath = "C:\Users\TsubotaniTA\VSCODE\RemoteWakeConnect\src\app.ico"

# Load the PNG image
$img = [System.Drawing.Image]::FromFile($sourcePath)

# Create a new bitmap with 32x32 size for ICO
$icon32 = New-Object System.Drawing.Bitmap(32, 32)
$graphics32 = [System.Drawing.Graphics]::FromImage($icon32)
$graphics32.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$graphics32.DrawImage($img, 0, 0, 32, 32)

# Save as ICO format
$icon32.Save($destPath, [System.Drawing.Imaging.ImageFormat]::Icon)

$graphics32.Dispose()
$icon32.Dispose()
$img.Dispose()

Write-Host "ICO file created at $destPath"