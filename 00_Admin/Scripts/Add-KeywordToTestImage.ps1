[CmdletBinding(DefaultParameterSetName = 'Add')]
param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Image,
    [Parameter(Mandatory, ParameterSetName = 'Add')][ValidateNotNullOrEmpty()][string]$Keyword,
    [Parameter(Mandatory, ParameterSetName = 'Clear')][switch]$ClearKeywords,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
$imagePath = $Image
if ($Image -eq [IO.Path]::GetFileName($Image) -and -not (Test-Path -LiteralPath $Image)) {
    $imagePath = Join-Path (Join-Path $PSScriptRoot '..\Metadata Test') $Image
}
$file = Get-Item -LiteralPath $imagePath
if ($file.PSIsContainer -or $file.Extension -notin '.jpg', '.jpeg') {
    throw 'Choose an existing JPEG file.'
}
$path = $file.FullName
$keywordText = $Keyword.Trim()
if (-not $ClearKeywords -and -not $keywordText) { throw 'Enter a test keyword.' }

if (-not $Apply) {
    if ($ClearKeywords) {
        [pscustomobject]@{ Image = $path; Change = 'Clear all keywords' }
    } else {
        [pscustomobject]@{ Image = $path; KeywordToAdd = $keywordText }
    }
    return
}

function Clear-JpegKeywords([string]$PhotoPath) {
    Add-Type -AssemblyName PresentationCore, WindowsBase
    $temp = Join-Path ([IO.Path]::GetDirectoryName($PhotoPath)) ([IO.Path]::GetRandomFileName())
    [IO.File]::Copy($PhotoPath, $temp)
    try {
        $stream = [IO.File]::Open($temp, 'Open', 'ReadWrite', 'None')
        try {
            $decoder = [Windows.Media.Imaging.JpegBitmapDecoder]::new($stream, 'PreservePixelFormat', 'None')
            if (-not $decoder.Frames[0].Metadata.Keywords) {
                return [pscustomobject]@{ Image = $PhotoPath; Result = 'No keywords to clear' }
            }
            $writer = $decoder.Frames[0].CreateInPlaceBitmapMetadataWriter()
            # Windows Imaging Component removes the keyword fields across JPEG metadata formats.
            $writer.RemoveQuery('System.Keywords')
            if (-not $writer.TrySave()) { throw 'Could not save the JPEG metadata in place.' }
        } finally {
            $stream.Dispose()
        }
        $check = [Windows.Media.Imaging.JpegBitmapDecoder]::new([Uri]::new($temp), 'PreservePixelFormat', 'OnLoad')
        if ($check.Frames[0].Metadata.Keywords) { throw 'Keyword verification failed after clearing.' }
        [IO.File]::Replace($temp, $PhotoPath, [NullString]::Value)
        [pscustomobject]@{ Image = $PhotoPath; Result = 'Keywords cleared' }
    } finally {
        if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    }
}

$lock = [Threading.Mutex]::new($false, 'Local\SportsPhotographyKeywordTest')
try {
    if (-not $lock.WaitOne(0)) { throw 'Another keyword update is in progress. Try again shortly.' }
    try {
        if ($ClearKeywords) {
            Clear-JpegKeywords $path
            return
        }
        $jpeg = [IO.File]::ReadAllBytes($path)
        if ($jpeg.Length -lt 4 -or $jpeg[0] -ne 0xff -or $jpeg[1] -ne 0xd8) {
            throw 'The test file is not a JPEG.'
        }

        # Find the standard XMP APP1 segment before the compressed image data.
        $xmpHeader = "http://ns.adobe.com/xap/1.0/`0"
        $encoding = [Text.UTF8Encoding]::new($false, $true)
        $xmlOffset = 0
        for ($offset = 2; $offset + 4 -le $jpeg.Length -and $jpeg[$offset] -eq 0xff; $offset = $end) {
            $marker = $jpeg[$offset + 1]
            if ($marker -eq 0xda -or $marker -eq 0xd9) { break }
            $length = 256 * [int]$jpeg[$offset + 2] + $jpeg[$offset + 3]
            $end = $offset + 2 + $length
            if ($length -lt 2 -or $end -gt $jpeg.Length) { throw 'Invalid JPEG segment length.' }
            if ($marker -eq 0xe1 -and $length -ge 2 + $xmpHeader.Length -and
                [Text.Encoding]::ASCII.GetString($jpeg, $offset + 4, $xmpHeader.Length) -ceq $xmpHeader) {
                $xmlOffset = $offset + 4 + $xmpHeader.Length
                break
            }
        }
        $insertPacket = -not $xmlOffset
        $xmp = [xml]::new()
        $xmp.PreserveWhitespace = $true
        if ($insertPacket) {
            if ($offset + 1 -ge $jpeg.Length -or $jpeg[$offset] -ne 0xff -or $jpeg[$offset + 1] -ne 0xda) {
                throw 'The JPEG has no readable start-of-scan marker.'
            }
            $xmp.LoadXml('<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about=""/></rdf:RDF></x:xmpmeta>')
        } else {
            $xmp.LoadXml($encoding.GetString($jpeg, $xmlOffset, $end - $xmlOffset))
        }
        $ns = [Xml.XmlNamespaceManager]::new($xmp.NameTable)
        $ns.AddNamespace('dc', 'http://purl.org/dc/elements/1.1/')
        $ns.AddNamespace('rdf', 'http://www.w3.org/1999/02/22-rdf-syntax-ns#')
        $subject = $xmp.SelectSingleNode('//rdf:Description/dc:subject', $ns)
        if ($subject) {
            $bag = $subject.SelectSingleNode('rdf:Bag', $ns)
            if (-not $bag) { throw 'The JPEG uses an unsupported XMP keyword list format.' }
        } else {
            $description = $xmp.SelectSingleNode('//rdf:RDF/rdf:Description', $ns)
            if (-not $description) { throw 'The JPEG needs an XMP description.' }
            $subject = $xmp.CreateElement('dc', 'subject', $ns.LookupNamespace('dc'))
            $bag = $xmp.CreateElement('rdf', 'Bag', $ns.LookupNamespace('rdf'))
            [void]$subject.AppendChild($bag)
            [void]$description.AppendChild($subject)
        }
        if ($bag.SelectNodes('rdf:li', $ns).InnerText -ccontains $keywordText) {
            [pscustomobject]@{ Image = $path; Keyword = $keywordText; Result = 'Already present' }
            return
        }
        $keywordNode = $xmp.CreateElement('rdf', 'li', $ns.LookupNamespace('rdf'))
        $keywordNode.InnerText = $keywordText
        [void]$bag.AppendChild($keywordNode)

        $packet = $encoding.GetBytes($xmp.OuterXml)
        $newLength = 2 + $xmpHeader.Length + $packet.Length
        if ($newLength -gt [uint16]::MaxValue) { throw 'The updated keyword list is too large for this JPEG segment.' }
        $lengthBytes = [byte[]]@(($newLength -shr 8),($newLength -band 255))
        $temp = Join-Path ([IO.Path]::GetDirectoryName($path)) ([IO.Path]::GetRandomFileName())
        try {
            $stream = [IO.File]::Open($temp, 'CreateNew', 'Write')
            try {
                if ($insertPacket) {
                    $header = [Text.Encoding]::ASCII.GetBytes($xmpHeader)
                    $stream.Write($jpeg, 0, $offset)
                    $stream.WriteByte(0xff)
                    $stream.WriteByte(0xe1)
                    $stream.Write($lengthBytes, 0, 2)
                    $stream.Write($header, 0, $header.Length)
                    $end = $offset
                } else {
                    $jpeg[$offset + 2] = $lengthBytes[0]
                    $jpeg[$offset + 3] = $lengthBytes[1]
                    $stream.Write($jpeg, 0, $xmlOffset)
                }
                $stream.Write($packet, 0, $packet.Length)
                $stream.Write($jpeg, $end, $jpeg.Length - $end)
            } finally {
                $stream.Dispose()
            }
            [IO.File]::Replace($temp, $path, [NullString]::Value)
        } finally {
            if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
        }
        [pscustomobject]@{ Image = $path; Keyword = $keywordText; Result = 'Added' }
    } finally {
        $lock.ReleaseMutex()
    }
} finally {
    $lock.Dispose()
}
