function Open-Decoder($Stream) {
    [Windows.Media.Imaging.BitmapDecoder]::Create($Stream, 'PreservePixelFormat', 'OnDemand')
}

function Get-ExifCaptureTime($File) {
    $stream = $null
    try {
        $stream = [IO.File]::Open($File.FullName, 'Open', 'Read', 'ReadWrite')
        $metadata = (Open-Decoder $stream).Frames[0].Metadata
        foreach ($query in '/app1/ifd/exif/{ushort=36867}','/ifd/exif/{ushort=36867}') {
            $stamp = $metadata.GetQuery($query)
            $parsed = [datetime]::MinValue
            if ($stamp -and [datetime]::TryParseExact($stamp.Trim(), 'yyyy:MM:dd HH:mm:ss',
                [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                return $parsed
            }
        }
    } catch {
        return
    } finally {
        if ($stream) { $stream.Dispose() }
    }
}
function Get-CaptureTime($Files, $ShellFolder) {
    foreach ($file in $Files) {
        if ($file.Extension -in ($jpeg + @('.tif','.tiff'))) {
            $exif = Get-ExifCaptureTime $file
            if ($exif) { return $exif }
        }
        if (-not $ShellFolder) { continue }
        try {
            $item = $ShellFolder.ParseName($file.Name)
            if ($item) {
                $date = $item.ExtendedProperty('System.Photo.DateTaken')
                if ($date -is [datetime] -and $date.Year -gt 1900) { return $date }
            }
        } catch { continue }
    }
    if ($Files[0].BaseName -match '^(\d{4}\.\d{2}\.\d{2}_\d{2}\.\d{2}\.\d{2})_') {
        $date = [datetime]::MinValue
        if ([datetime]::TryParseExact($Matches[1], 'yyyy.MM.dd_HH.mm.ss',
            [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) {
            return $date
        }
    }
}
