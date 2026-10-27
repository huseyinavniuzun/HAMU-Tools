# Runs only for a photo import. Original files are read-only sources.
Add-Type -AssemblyName System.Drawing
function Convert-HamuPhotos([string[]]$sources,[string]$destination) {
 $result=New-Object 'Collections.Generic.List[string]'
 $jpeg=[Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object MimeType -eq 'image/jpeg'
 $parameters=New-Object Drawing.Imaging.EncoderParameters(1)
 $parameters.Param[0]=New-Object Drawing.Imaging.EncoderParameter([Drawing.Imaging.Encoder]::Quality,[long]85)
 try {
  for($index=0;$index -lt $sources.Count;$index++) {
   $image=$null;$bitmap=$null;$graphics=$null;$stream=$null
   try {
    try {$image=[Drawing.Image]::FromFile($sources[$index])}
    catch {
     # WebP uses the installed Windows image decoder; keep alpha in the fallback.
     Add-Type -AssemblyName PresentationCore
     $uri=New-Object Uri($sources[$index])
     $frame=[Windows.Media.Imaging.BitmapFrame]::Create($uri,[Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache,[Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
     $encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder
     $encoder.Frames.Add($frame)
     $stream=New-Object IO.MemoryStream
     $encoder.Save($stream);$stream.Position=0
     $image=[Drawing.Image]::FromStream($stream)
    }
    $orientation=1
    if($image.PropertyIdList -contains 274) {
     $orientation=[BitConverter]::ToUInt16($image.GetPropertyItem(274).Value,0)
     $rotations=@{2=4;3=2;4=6;5=5;6=1;7=7;8=3}
     if($rotations.ContainsKey([int]$orientation)) {$image.RotateFlip([Drawing.RotateFlipType]$rotations[[int]$orientation])}
    }
    $limitW=1920;$limitH=1080
    if($image.Height -gt $image.Width) {$limitW=1080;$limitH=1920}
    $scale=[Math]::Min(1.0,[Math]::Min($limitW/[double]$image.Width,$limitH/[double]$image.Height))
    $width=[Math]::Max(1,[int][Math]::Floor($image.Width*$scale))
    $height=[Math]::Max(1,[int][Math]::Floor($image.Height*$scale))
    $alpha=($image.Flags -band 2) -ne 0 -or [Drawing.Image]::IsAlphaPixelFormat($image.PixelFormat)
    $format=[Drawing.Imaging.PixelFormat]::Format24bppRgb
    if($alpha) {$format=[Drawing.Imaging.PixelFormat]::Format32bppArgb}
    $bitmap=New-Object Drawing.Bitmap($width,$height,$format)
    $graphics=[Drawing.Graphics]::FromImage($bitmap)
    $graphics.CompositingMode=[Drawing.Drawing2D.CompositingMode]::SourceCopy
    $graphics.InterpolationMode=[Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.PixelOffsetMode=[Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $attributes=New-Object Drawing.Imaging.ImageAttributes
    try {
     $attributes.SetWrapMode([Drawing.Drawing2D.WrapMode]::TileFlipXY)
     $graphics.DrawImage($image,(New-Object Drawing.Rectangle(0,0,$width,$height)),0,0,$image.Width,$image.Height,[Drawing.GraphicsUnit]::Pixel,$attributes)
    } finally {$attributes.Dispose()}
    $extension='jpg';if($alpha){$extension='png'}
    $path=Join-Path $destination ('optimized_'+$index+'.'+$extension)
    if($alpha){$bitmap.Save($path,[Drawing.Imaging.ImageFormat]::Png)}else{$bitmap.Save($path,$jpeg,$parameters)}
    # Avoid enlarging already small JPEG/PNG files by recompressing them.
    $originalExtension=[IO.Path]::GetExtension($sources[$index]).ToLowerInvariant()
    if($scale -eq 1 -and $orientation -eq 1 -and $originalExtension -in @('.jpg','.jpeg','.png') -and (Get-Item -LiteralPath $path).Length -ge (Get-Item -LiteralPath $sources[$index]).Length) {
     [void]$result.Add($sources[$index])
    } else {[void]$result.Add($path)}
   } finally {
    if($graphics){$graphics.Dispose()};if($bitmap){$bitmap.Dispose()};if($image){$image.Dispose()};if($stream){$stream.Dispose()}
   }
  }
 } finally {$parameters.Dispose()}
 return ,$result.ToArray()
}
