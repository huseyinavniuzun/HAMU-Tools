$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression
$paths=@([IO.File]::ReadAllLines($manifest,[Text.Encoding]::Unicode))
if($paths.Count -eq 0){throw 'No images'}
if($optimize){$paths=Convert-HamuPhotos $paths ([IO.Path]::GetDirectoryName($output))}
$seedStream=New-Object IO.MemoryStream(,[Convert]::FromBase64String($seed64))
$seedZip=New-Object IO.Compression.ZipArchive($seedStream,[IO.Compression.ZipArchiveMode]::Read,$false)
$fileStream=[IO.File]::Open($output,[IO.FileMode]::CreateNew,[IO.FileAccess]::ReadWrite)
$zip=New-Object IO.Compression.ZipArchive($fileStream,[IO.Compression.ZipArchiveMode]::Create,$false)
function Add-TextEntry([string]$name,[string]$text){
 $entry=$zip.CreateEntry($name,[IO.Compression.CompressionLevel]::NoCompression)
 $writer=New-Object IO.StreamWriter($entry.Open(),[Text.UTF8Encoding]::new($false))
 try{$writer.Write($text)}finally{$writer.Dispose()}
}
try{
 $dynamic=@('[Content_Types].xml','xl/worksheets/sheet1.xml','xl/metadata.xml','xl/richData/rdrichvalue.xml','xl/richData/richValueRel.xml','xl/richData/_rels/richValueRel.xml.rels')
 foreach($entry in $seedZip.Entries){
  if($dynamic -contains $entry.FullName -or $entry.FullName.StartsWith('xl/media/')){continue}
  $dest=$zip.CreateEntry($entry.FullName,[IO.Compression.CompressionLevel]::NoCompression)
  $srcStream=$entry.Open();$destStream=$dest.Open()
  try{$srcStream.CopyTo($destStream)}finally{$srcStream.Dispose();$destStream.Dispose()}
 }
 $reader=New-Object IO.StreamReader($seedZip.GetEntry('[Content_Types].xml').Open())
 try{[xml]$types=$reader.ReadToEnd()}finally{$reader.Dispose()}
 $extensions=@{}
 foreach($item in $types.DocumentElement.ChildNodes){if($item.LocalName -eq 'Default'){$extensions[$item.Extension]=$true}}
 $mime=@{png='image/png';jpg='image/jpeg';jpeg='image/jpeg';bmp='image/bmp';tif='image/tiff';tiff='image/tiff';gif='image/gif';webp='image/webp'}
 $rows=New-Object Text.StringBuilder
 $future=New-Object Text.StringBuilder
 $metadata=New-Object Text.StringBuilder
 $values=New-Object Text.StringBuilder
 $rels=New-Object Text.StringBuilder
 $images=New-Object Text.StringBuilder
 for($i=0;$i -lt $paths.Count;$i++){
  $number=$i+1;$extension=[IO.Path]::GetExtension($paths[$i]).TrimStart('.').ToLowerInvariant()
  if(-not $mime.ContainsKey($extension)){throw 'Unsupported image format'}
  if(-not $extensions.ContainsKey($extension)){
   $node=$types.CreateElement('Default',$types.DocumentElement.NamespaceURI)
   $node.SetAttribute('Extension',$extension);$node.SetAttribute('ContentType',$mime[$extension])
   [void]$types.DocumentElement.AppendChild($node);$extensions[$extension]=$true
  }
  $entry=$zip.CreateEntry("xl/media/image$number.$extension",[IO.Compression.CompressionLevel]::NoCompression)
  $sourceStream=[IO.File]::OpenRead($paths[$i]);$destStream=$entry.Open()
  try{$sourceStream.CopyTo($destStream)}finally{$sourceStream.Dispose();$destStream.Dispose()}
  [void]$rows.Append(('<row r="{0}"><c r="A{0}" t="e" vm="{0}"><v>#VALUE!</v></c></row>' -f $number))
  [void]$future.Append(('<bk><extLst><ext uri="{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}"><xlrd:rvb i="{0}"/></ext></extLst></bk>' -f $i))
  [void]$metadata.Append(('<bk><rc t="1" v="{0}"/></bk>' -f $i))
  [void]$values.Append(('<rv s="0"><v>{0}</v><v>5</v></rv>' -f $i))
  [void]$rels.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image{0}.{1}"/>' -f $number,$extension))
  [void]$images.Append(('<rel r:id="rId{0}"/>' -f $number))
 }
 $count=$paths.Count
 Add-TextEntry '[Content_Types].xml' $types.OuterXml
 Add-TextEntry 'xl/worksheets/sheet1.xml' ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A'+$count+'"/><sheetData>'+$rows.ToString()+'</sheetData></worksheet>')
 Add-TextEntry 'xl/metadata.xml' ('<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="'+$count+'">'+$future.ToString()+'</futureMetadata><valueMetadata count="'+$count+'">'+$metadata.ToString()+'</valueMetadata></metadata>')
 Add-TextEntry 'xl/richData/rdrichvalue.xml' ('<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="'+$count+'">'+$values.ToString()+'</rvData>')
 Add-TextEntry 'xl/richData/richValueRel.xml' ('<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+$images.ToString()+'</richValueRels>')
 Add-TextEntry 'xl/richData/_rels/richValueRel.xml.rels' ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+$rels.ToString()+'</Relationships>')
}finally{$zip.Dispose();$fileStream.Dispose();$seedZip.Dispose();$seedStream.Dispose()}
