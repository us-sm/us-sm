<# :
@echo off
rem Word Merge: drag .doc / .docx files onto this file (or double-click it).
copy /y "%~f0" "%TEMP%\wordmerge_run.ps1" >nul
start "" powershell -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TEMP%\wordmerge_run.ps1" %*
exit /b
#>
Add-Type -AssemblyName System.Windows.Forms
trap { [void][System.Windows.Forms.MessageBox]::Show("Unexpected error:`n" + $_.Exception.Message, 'Word Merge', 'OK', 'Error'); exit 1 }
# ===== .docx merge (port of the web page's merge-core.js) =====
# Each source keeps its own section settings (margins, page size, orientation,
# headers/footers). PowerShell 5.1 compatible.

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:NS_W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$script:NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:NS_PR  = 'http://schemas.openxmlformats.org/package/2006/relationships'
$script:NS_CT  = 'http://schemas.openxmlformats.org/package/2006/content-types'
$script:NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
$script:NS_MC  = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
$script:NS_X   = 'http://www.w3.org/2000/xmlns/'
$script:RT     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
$script:SKIP_TYPES = @('styles','numbering','settings','webSettings','fontTable','theme','footnotes',
  'endnotes','comments','customXml','glossaryDocument','people','commentsExtended','commentsIds',
  'commentsExtensible','stylesWithEffects')

function Expand-Docx([string]$path, [string]$dir) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        foreach ($e in $zip.Entries) {
            if ($e.FullName.EndsWith('/')) { continue }
            $dest = Join-Path $dir ($e.FullName -replace '\\','/')
            New-Item -ItemType Directory -Path (Split-Path $dest -Parent) -Force | Out-Null
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($e, $dest, $true)
        }
    } finally { $zip.Dispose() }
}

function Compress-Docx([string]$dir, [string]$outPath) {
    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    $fs = [System.IO.File]::Open($outPath, 'CreateNew')
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $root = (Resolve-Path -LiteralPath $dir).Path.TrimEnd('\','/')
        $files = Get-ChildItem -LiteralPath $root -Recurse -File -Force
        # [Content_Types].xml first, like Word does
        $files = @($files | Sort-Object { if ($_.Name -eq '[Content_Types].xml') { 0 } else { 1 } })
        foreach ($f in $files) {
            $rel = $f.FullName.Substring($root.Length + 1) -replace '\\','/'
            $entry = $zip.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
            $s = $entry.Open()
            try { $b = [System.IO.File]::ReadAllBytes($f.FullName); $s.Write($b, 0, $b.Length) } finally { $s.Dispose() }
        }
    } finally { $zip.Dispose(); $fs.Dispose() }
}

function P([string]$dir, [string]$part) { return (Join-Path $dir ($part -replace '/', [IO.Path]::DirectorySeparatorChar)) }

function Read-Xml([string]$dir, [string]$part) {
    $p = P $dir $part
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    $d = New-Object System.Xml.XmlDocument
    $d.PreserveWhitespace = $true
    $d.Load($p)
    return ,$d
}

function Save-Xml([System.Xml.XmlDocument]$doc, [string]$dir, [string]$part) {
    $p = P $dir $part
    New-Item -ItemType Directory -Path (Split-Path $p -Parent) -Force | Out-Null
    $set = New-Object System.Xml.XmlWriterSettings
    $set.Encoding = New-Object System.Text.UTF8Encoding($false)
    $w = [System.Xml.XmlWriter]::Create($p, $set)
    try { $doc.Save($w) } finally { $w.Dispose() }
}

function Kids($el, [string]$ns, [string]$name) {
    $r = @()
    foreach ($c in $el.ChildNodes) { if ($c.NodeType -eq 'Element' -and $c.NamespaceURI -eq $ns -and $c.LocalName -eq $name) { $r += $c } }
    return ,$r
}
function All($node, [string]$ns, [string]$name) { return ,@($node.GetElementsByTagName($name, $ns)) }

function Get-WAttr($el, [string]$name) { return $el.GetAttribute($name, $script:NS_W) }
function Set-NsAttr($el, [string]$prefix, [string]$ns, [string]$name, [string]$value) {
    $a = $el.GetAttributeNode($name, $ns)
    if ($a) { $a.Value = $value; return }
    $a = $el.OwnerDocument.CreateAttribute($prefix, $name, $ns)
    $a.Value = $value
    [void]$el.Attributes.Append($a)
}
function Set-WAttr($el, [string]$name, [string]$value) { Set-NsAttr $el 'w' $script:NS_W $name $value }
function New-W($doc, [string]$name) { return $doc.CreateElement('w', $name, $script:NS_W) }

function DirOf([string]$p) { $i = $p.LastIndexOf('/'); if ($i -ge 0) { return $p.Substring(0, $i + 1) } else { return '' } }
function RelsPathOf([string]$p) { return (DirOf $p) + '_rels/' + $p.Substring($p.LastIndexOf('/') + 1) + '.rels' }
function Resolve-Part([string]$base, [string]$target) {
    if ($target.StartsWith('/')) { return $target.Substring(1) }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($x in ((DirOf $base) + $target).Split('/')) {
        if ($x -eq '..') { if ($out.Count) { $out.RemoveAt($out.Count - 1) } }
        elseif ($x -ne '.' -and $x -ne '') { $out.Add($x) }
    }
    return ($out -join '/')
}
function Relative-Part([string]$fromPart, [string]$toPath) {
    $from = @((DirOf $fromPart).Split('/') | Where-Object { $_ })
    $to = $toPath.Split('/')
    $i = 0
    while ($i -lt $from.Count -and $i -lt $to.Count - 1 -and $from[$i] -eq $to[$i]) { $i++ }
    return ('../' * ($from.Count - $i)) + (($to[$i..($to.Count - 1)]) -join '/')
}
function RelType([string]$t) { return $t.Substring($t.LastIndexOf('/') + 1) }

function Get-MainPart([string]$dir) {
    $rels = Read-Xml $dir '_rels/.rels'
    foreach ($r in (All $rels $script:NS_PR 'Relationship')) {
        if ($r.GetAttribute('Type').EndsWith('/officeDocument')) { return $r.GetAttribute('Target').TrimStart('/') }
    }
    throw 'Not a Word document'
}

$script:PAPER = @{ letter = @(12240, 15840, 1); legal = @(12240, 20160, 5); a4 = @(11906, 16838, 9) }
# Force every section to one paper size, keeping each section's orientation.
# Page size of a .docx's first section as @(short side, long side, code).
function Get-PageSizeOf([string]$docxPath) {
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('docxsize_' + [Guid]::NewGuid().ToString('N'))
    try {
        Expand-Docx $docxPath $tmp
        $doc = Read-Xml $tmp (Get-MainPart $tmp)
        foreach ($sp in (All $doc $script:NS_W 'sectPr')) {
            if ($sp.ParentNode.LocalName -eq 'sectPrChange') { continue }
            $sz = (Kids $sp $script:NS_W 'pgSz')[0]
            if (-not $sz) { return $null }
            $w = 0; $h = 0
            [void][int]::TryParse((Get-WAttr $sz 'w'), [ref]$w); [void][int]::TryParse((Get-WAttr $sz 'h'), [ref]$h)
            if (-not $w -or -not $h) { return $null }
            return ,@([Math]::Min($w, $h), [Math]::Max($w, $h), (Get-WAttr $sz 'code'))
        }
        return $null
    } finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

# $paper: 'letter' | 'legal' | 'a4' | '' | @(short, long, code)
function Set-PaperSize([string]$dir, $paper) {
    if (-not $paper) { return }
    if ($paper -is [array]) { $spec = $paper }
    elseif ($script:PAPER.ContainsKey([string]$paper)) { $spec = $script:PAPER[[string]$paper] }
    else { return }
    $main = Get-MainPart $dir
    $doc = Read-Xml $dir $main
    foreach ($sp in (All $doc $script:NS_W 'sectPr')) {
        if ($sp.ParentNode.LocalName -eq 'sectPrChange') { continue }
        $sz = (Kids $sp $script:NS_W 'pgSz')[0]
        if (-not $sz) {
            $sz = New-W $doc 'pgSz'
            $before = $null
            foreach ($c in $sp.ChildNodes) { if ($c.NodeType -eq 'Element' -and @('headerReference','footerReference','footnotePr','endnotePr','type') -notcontains $c.LocalName) { $before = $c; break } }
            [void]$sp.InsertBefore($sz, $before)
        }
        $w = 0; $h = 0
        [void][int]::TryParse((Get-WAttr $sz 'w'), [ref]$w); [void][int]::TryParse((Get-WAttr $sz 'h'), [ref]$h)
        $landscape = ((Get-WAttr $sz 'orient') -eq 'landscape') -or ($w -gt $h)
        if ($landscape) { Set-WAttr $sz 'w' ([string]$spec[1]); Set-WAttr $sz 'h' ([string]$spec[0]) }
        else { Set-WAttr $sz 'w' ([string]$spec[0]); Set-WAttr $sz 'h' ([string]$spec[1]) }
        if ([string]$spec[2] -ne '') { Set-WAttr $sz 'code' ([string]$spec[2]) } else { $sz.RemoveAttribute('code', $script:NS_W) }
    }
    Save-Xml $doc $dir $main
}

function Merge-DocxFiles([string[]]$paths, [string]$outPath, $paper = '') {
    if ($paths.Count -lt 2) { throw 'Need at least 2 files' }
    $work = Join-Path ([IO.Path]::GetTempPath()) ('docxmerge_' + [Guid]::NewGuid().ToString('N'))
    try {
        $outDir = Join-Path $work 'out'
        Expand-Docx $paths[0] $outDir
        # remember each input's compatibility settings before anything changes
        $compats = @(Get-CompatInfo $outDir)
        for ($i = 1; $i -lt $paths.Count; $i++) {
            $srcDir = Join-Path $work ('src' + $i)
            Expand-Docx $paths[$i] $srcDir
            $compats += Get-CompatInfo $srcDir
            Add-DocxTo $outDir $srcDir $i
        }
        Set-OldestCompat $outDir $compats
        Set-PaperSize $outDir $paper
        Compress-Docx $outDir $outPath
    } finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Compatibility mode (layout rules) lives in settings.xml <w:compat>.
# Files converted from .doc are in mode 11; the merged file uses the OLDEST mode
# of all inputs so .doc-derived pages keep their original layout.
function Get-SettingsPart([string]$dir) {
    $main = Get-MainPart $dir
    $rels = Read-Xml $dir (RelsPathOf $main)
    foreach ($r in (All $rels $script:NS_PR 'Relationship')) {
        if ((RelType $r.GetAttribute('Type')) -eq 'settings') { return (Resolve-Part $main $r.GetAttribute('Target')) }
    }
    return $null
}
function Get-CompatInfo([string]$dir) {
    $part = Get-SettingsPart $dir
    $mode = 12; $compat = $null
    if ($part) {
        $sd = Read-Xml $dir $part
        if ($sd) {
            $compat = (Kids $sd.DocumentElement $script:NS_W 'compat')[0]
            if ($compat) {
                foreach ($cs in (Kids $compat $script:NS_W 'compatSetting')) {
                    if ($cs.GetAttribute('name', $script:NS_W) -eq 'compatibilityMode') { $mode = [int]$cs.GetAttribute('val', $script:NS_W) }
                }
            }
        }
    }
    return ,@{ mode = $mode; compat = $compat }
}
function Set-OldestCompat([string]$outDir, $compats) {
    $oldest = $compats[0]
    foreach ($c in $compats) { if ($c.mode -lt $oldest.mode) { $oldest = $c } }
    if ($oldest -eq $compats[0] -or -not $oldest.compat) { return }
    $part = Get-SettingsPart $outDir
    if (-not $part) { return }
    $sd = Read-Xml $outDir $part
    $root = $sd.DocumentElement
    $new = $sd.ImportNode($oldest.compat, $true)
    $cur = (Kids $root $script:NS_W 'compat')[0]
    if ($cur) { [void]$root.ReplaceChild($new, $cur) }
    else {
        # schema order: compat comes right before docVars/rsids/mathPr/... ; place it before the first of those
        $after = @('docVars','rsids','mathPr','attachedSchema','themeFontLang','clrSchemeMapping','doNotIncludeSubdocsInStats','doNotAutoCompressPictures','forceUpgrade','captions','readModeInkLockDown','smartTagType','schemaLibrary','shapeDefaults','doNotEmbedSmartTags','decimalSymbol','listSeparator')
        $before = $null
        foreach ($ch in $root.ChildNodes) { if ($ch.NodeType -eq 'Element' -and $after -contains $ch.LocalName) { $before = $ch; break } }
        [void]$root.InsertBefore($new, $before)
    }
    Save-Xml $sd $outDir $part
}

function Add-DocxTo([string]$out, [string]$src, [int]$idx) {
    $xW = $script:NS_W; $xR = $script:NS_R; $xPR = $script:NS_PR; $xCT = $script:NS_CT
    $outMain = Get-MainPart $out
    $srcMain = Get-MainPart $src
    $outDoc = Read-Xml $out $outMain
    $srcDoc = Read-Xml $src $srcMain
    $outRels = Read-Xml $out (RelsPathOf $outMain)
    $srcRels = Read-Xml $src (RelsPathOf $srcMain)
    $outCT = Read-Xml $out '[Content_Types].xml'
    $srcCT = Read-Xml $src '[Content_Types].xml'
    $prefix = 'm' + $idx + '_'

    # ---- content types
    $srcOverride = @{}
    foreach ($o in (All $srcCT $xCT 'Override')) { $srcOverride[$o.GetAttribute('PartName')] = $o.GetAttribute('ContentType') }
    $ctRoot = $outCT.DocumentElement
    $ensureDefault = {
        param($ext)
        foreach ($d in (All $outCT $xCT 'Default')) { if ($d.GetAttribute('Extension') -ieq $ext) { return } }
        foreach ($s in (All $srcCT $xCT 'Default')) {
            if ($s.GetAttribute('Extension') -ieq $ext) {
                $d = $outCT.CreateElement('Default', $xCT)
                $d.SetAttribute('Extension', $s.GetAttribute('Extension'))
                $d.SetAttribute('ContentType', $s.GetAttribute('ContentType'))
                [void]$ctRoot.InsertBefore($d, $ctRoot.FirstChild)
                return
            }
        }
    }
    $addOverride = {
        param($srcPath, $newPath)
        $ct = $srcOverride['/' + $srcPath]
        if ($ct) {
            $o = $outCT.CreateElement('Override', $xCT)
            $o.SetAttribute('PartName', '/' + $newPath)
            $o.SetAttribute('ContentType', $ct)
            [void]$ctRoot.AppendChild($o)
        } else {
            & $ensureDefault ($newPath.Substring($newPath.LastIndexOf('.') + 1))
        }
    }

    # ---- copy parts (recursively with their rels)
    $copied = @{}
    $uniqueName = {
        param($path)
        $d = DirOf $path; $b = $path.Substring($d.Length)
        $n = $d + $prefix + $b; $k = 1
        while (Test-Path -LiteralPath (P $out $n)) { $n = $d + $prefix + $k + '_' + $b; $k++ }
        return $n
    }
    $copyPart = $null
    $copyPart = {
        param($srcPath)
        if ($copied.ContainsKey($srcPath)) { return $copied[$srcPath] }
        $sp = P $src $srcPath
        if (-not (Test-Path -LiteralPath $sp)) { return $null }
        $newPath = & $uniqueName $srcPath
        $copied[$srcPath] = $newPath
        $np = P $out $newPath
        New-Item -ItemType Directory -Path (Split-Path $np -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $sp -Destination $np
        & $addOverride $srcPath $newPath
        $rd = Read-Xml $src (RelsPathOf $srcPath)
        if ($rd) {
            foreach ($rel in (All $rd $xPR 'Relationship')) {
                if ($rel.GetAttribute('TargetMode') -eq 'External') { continue }
                $t = Resolve-Part $srcPath $rel.GetAttribute('Target')
                $n2 = & $copyPart $t
                if ($n2) { $rel.SetAttribute('Target', (Relative-Part $newPath $n2)) }
            }
            Save-Xml $rd $out (RelsPathOf $newPath)
        }
        return $newPath
    }

    # ---- main document relationships
    $relRoot = $outRels.DocumentElement
    $used = @{}
    foreach ($r in (All $outRels $xPR 'Relationship')) { $used[$r.GetAttribute('Id')] = $true }
    $script:nextRelId = 1
    $newRelId = {
        while ($used.ContainsKey('rId' + $script:nextRelId)) { $script:nextRelId++ }
        $id = 'rId' + $script:nextRelId; $used[$id] = $true; return $id
    }
    $addRel = {
        param($typeName, $target)
        $nr = $outRels.CreateElement('Relationship', $xPR)
        $id = & $newRelId
        $nr.SetAttribute('Id', $id); $nr.SetAttribute('Type', $script:RT + $typeName); $nr.SetAttribute('Target', $target)
        [void]$relRoot.AppendChild($nr)
        return $id
    }
    $findRel = {
        param($rels, $typeName)
        foreach ($r in (All $rels $xPR 'Relationship')) { if ((RelType $r.GetAttribute('Type')) -eq $typeName) { return $r } }
        return $null
    }

    $ridMap = @{}
    $styledParts = New-Object System.Collections.Generic.List[string]
    foreach ($rel in (All $srcRels $xPR 'Relationship')) {
        $type = $rel.GetAttribute('Type')
        if ($script:SKIP_TYPES -contains (RelType $type)) { continue }
        $nr = $outRels.CreateElement('Relationship', $xPR)
        if ($rel.GetAttribute('TargetMode') -eq 'External') {
            $nr.SetAttribute('Target', $rel.GetAttribute('Target'))
            $nr.SetAttribute('TargetMode', 'External')
        } else {
            $np = & $copyPart (Resolve-Part $srcMain $rel.GetAttribute('Target'))
            if (-not $np) { continue }
            $nr.SetAttribute('Target', (Relative-Part $outMain $np))
            if (@('header','footer') -contains (RelType $type)) { $styledParts.Add($np) }
        }
        $id = & $newRelId
        $nr.SetAttribute('Id', $id); $nr.SetAttribute('Type', $type)
        # attribute order: Id first is nicer
        [void]$relRoot.AppendChild($nr)
        $ridMap[$rel.GetAttribute('Id')] = $id
    }

    $srcBody = (Kids $srcDoc.DocumentElement $xW 'body')[0]
    $outBody = (Kids $outDoc.DocumentElement $xW 'body')[0]

    # ---- drop comment anchors
    foreach ($n in @('commentRangeStart','commentRangeEnd','commentReference')) {
        foreach ($e in (All $srcDoc $xW $n)) {
            $p = $e.ParentNode
            [void]$p.RemoveChild($e)
            if ($p.NamespaceURI -eq $xW -and $p.LocalName -eq 'r') {
                $other = $false
                foreach ($c in $p.ChildNodes) { if ($c.NodeType -eq 'Element' -and $c.LocalName -ne 'rPr') { $other = $true } }
                if (-not $other -and $p.ParentNode) { [void]$p.ParentNode.RemoveChild($p) }
            }
        }
    }
    # ---- remap r:* attributes, drop w14 ids
    foreach ($el in @($srcBody.GetElementsByTagName('*'))) {
        foreach ($a in @($el.Attributes)) {
            if ($a.NamespaceURI -eq $xR -and $ridMap.ContainsKey($a.Value)) { $a.Value = $ridMap[$a.Value] }
        }
        $el.RemoveAttribute('paraId', $script:NS_W14)
        $el.RemoveAttribute('textId', $script:NS_W14)
    }
    # ---- unique drawing ids & bookmark ids
    $maxDocPr = 0
    foreach ($e in @($outDoc.GetElementsByTagName('*'))) { if ($e.LocalName -eq 'docPr') { $v = 0; [void][int]::TryParse($e.GetAttribute('id'), [ref]$v); if ($v -gt $maxDocPr) { $maxDocPr = $v } } }
    $off = $maxDocPr + 1
    foreach ($e in @($srcBody.GetElementsByTagName('*'))) {
        if ($e.LocalName -eq 'docPr' -or ($e.LocalName -eq 'cNvPr' -and $e.ParentNode.LocalName -eq 'nvPicPr')) {
            $v = 0; [void][int]::TryParse($e.GetAttribute('id'), [ref]$v); $e.SetAttribute('id', [string]($v + $off))
        }
    }
    $maxBm = 0
    foreach ($e in (All $outDoc $xW 'bookmarkStart')) { $v = 0; [void][int]::TryParse((Get-WAttr $e 'id'), [ref]$v); if ($v -gt $maxBm) { $maxBm = $v } }
    foreach ($e in ((All $srcDoc $xW 'bookmarkStart') + (All $srcDoc $xW 'bookmarkEnd'))) {
        $v = 0; [void][int]::TryParse((Get-WAttr $e 'id'), [ref]$v); Set-WAttr $e 'id' ([string]($v + $maxBm + 1))
    }

    # ---- numbering
    $numMap = $null
    $srcNumRel = & $findRel $srcRels 'numbering'
    if ($srcNumRel) {
        $srcNumPath = Resolve-Part $srcMain $srcNumRel.GetAttribute('Target')
        $srcNum = Read-Xml $src $srcNumPath
        $outNumRel = & $findRel $outRels 'numbering'
        if ($srcNum -and -not $outNumRel) {
            $np = & $copyPart $srcNumPath
            [void](& $addRel 'numbering' (Relative-Part $outMain $np))
        } elseif ($srcNum) {
            $outNumPath = Resolve-Part $outMain $outNumRel.GetAttribute('Target')
            $outNum = Read-Xml $out $outNumPath
            $nroot = $outNum.DocumentElement
            $absMax = -1; foreach ($e in (All $outNum $xW 'abstractNum')) { $v = [int](Get-WAttr $e 'abstractNumId'); if ($v -gt $absMax) { $absMax = $v } }
            $numMax = 0;  foreach ($e in (All $outNum $xW 'num')) { $v = [int](Get-WAttr $e 'numId'); if ($v -gt $numMax) { $numMax = $v } }
            $absMap = @{}; $numMap = @{}
            $firstNum = (Kids $nroot $xW 'num')[0]
            $rnd = New-Object System.Random
            foreach ($a in (Kids $srcNum.DocumentElement $xW 'abstractNum')) {
                $c = $outNum.ImportNode($a, $true)
                $absMax++
                $absMap[(Get-WAttr $a 'abstractNumId')] = $absMax
                Set-WAttr $c 'abstractNumId' ([string]$absMax)
                $nsid = (Kids $c $xW 'nsid')[0]
                if ($nsid) { Set-WAttr $nsid 'val' ('{0:X8}' -f $rnd.Next()) }
                [void]$nroot.InsertBefore($c, $firstNum)
            }
            $cleanup = (Kids $nroot $xW 'numIdMacAtCleanup')[0]
            foreach ($n in (Kids $srcNum.DocumentElement $xW 'num')) {
                $c = $outNum.ImportNode($n, $true)
                $numMax++
                $numMap[(Get-WAttr $n 'numId')] = $numMax
                Set-WAttr $c 'numId' ([string]$numMax)
                $an = (Kids $c $xW 'abstractNumId')[0]
                if ($an -and $absMap.ContainsKey((Get-WAttr $an 'val'))) { Set-WAttr $an 'val' ([string]$absMap[(Get-WAttr $an 'val')]) }
                [void]$nroot.InsertBefore($c, $cleanup)
            }
            # paragraphs numbered through a style: pin the source's own list
            $sStyRel = & $findRel $srcRels 'styles'
            $sSty = $null
            if ($sStyRel) { $sSty = Read-Xml $src (Resolve-Part $srcMain $sStyRel.GetAttribute('Target')) }
            if ($sSty) {
                $styNum = @{}
                foreach ($st in (All $sSty $xW 'style')) {
                    $np = @($st.GetElementsByTagName('numPr', $xW))[0]
                    if (-not $np) { continue }
                    $ni = (Kids $np $xW 'numId')[0]
                    if (-not $ni) { continue }
                    $il = (Kids $np $xW 'ilvl')[0]
                    $lvl = '0'; if ($il) { $lvl = Get-WAttr $il 'val' }
                    $styNum[(Get-WAttr $st 'styleId')] = @{ num = (Get-WAttr $ni 'val'); lvl = $lvl }
                }
                $ORDER = @('pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl')
                foreach ($pPr in (All $srcDoc $xW 'pPr')) {
                    $ps = (Kids $pPr $xW 'pStyle')[0]
                    if (-not $ps) { continue }
                    $info = $styNum[(Get-WAttr $ps 'val')]
                    if (-not $info -or (Kids $pPr $xW 'numPr').Count) { continue }
                    $npEl = New-W $srcDoc 'numPr'
                    $il = New-W $srcDoc 'ilvl'; Set-WAttr $il 'val' $info.lvl
                    $ni = New-W $srcDoc 'numId'; Set-WAttr $ni 'val' $info.num
                    [void]$npEl.AppendChild($il); [void]$npEl.AppendChild($ni)
                    $before = $null
                    foreach ($c in $pPr.ChildNodes) { if ($c.NodeType -eq 'Element' -and $ORDER -notcontains $c.LocalName) { $before = $c; break } }
                    [void]$pPr.InsertBefore($npEl, $before)
                }
            }
            foreach ($e in (All $srcDoc $xW 'numId')) {
                $v = Get-WAttr $e 'val'
                if ($numMap.ContainsKey($v)) { Set-WAttr $e 'val' ([string]$numMap[$v]) }
            }
            Save-Xml $outNum $out $outNumPath
        }
    }

    # ---- styles ("keep source formatting")
    # A source style whose id also exists in the base doc but is defined
    # differently (or inherits from such a style, or the documents' defaults
    # differ) is copied under a new id, and the source content points at it.
    $styleMap = @{}; $srcDefault = @{}
    $srcStyRel = & $findRel $srcRels 'styles'; $outStyRel = & $findRel $outRels 'styles'
    if ($srcStyRel -and $outStyRel) {
        $outStyPath = Resolve-Part $outMain $outStyRel.GetAttribute('Target')
        $outSty = Read-Xml $out $outStyPath
        $srcSty = Read-Xml $src (Resolve-Part $srcMain $srcStyRel.GetAttribute('Target'))
        if ($srcSty -and $outSty) {
            $outById = @{}; $srcById = @{}; $srcOrder = @()
            foreach ($e in (Kids $outSty.DocumentElement $xW 'style')) { $outById[(Get-WAttr $e 'styleId')] = $e }
            foreach ($e in (Kids $srcSty.DocumentElement $xW 'style')) { $srcById[(Get-WAttr $e 'styleId')] = $e; $srcOrder += $e }
            $canon = {
                param($el)
                if (-not $el) { return '' }
                $c = $el.CloneNode($true)
                $nodes = @($c) + @($c.GetElementsByTagName('*'))
                foreach ($x in $nodes) {
                    foreach ($a in @($x.Attributes)) { if ($a.LocalName -like 'rsid*') { [void]$x.Attributes.Remove($a) } }
                }
                foreach ($n in @('rsid','uiPriority','semiHidden','unhideWhenUsed','qFormat','locked','personal','personalCompose','personalReply','autoRedefine','hidden')) {
                    foreach ($x in (Kids $c $xW $n)) { [void]$c.RemoveChild($x) }
                }
                return ($c.OuterXml -replace '\s+xmlns(:\w+)?="[^"]*"', '')
            }
            $defaultsDiffer = (& $canon (Kids $outSty.DocumentElement $xW 'docDefaults')[0]) -ne (& $canon (Kids $srcSty.DocumentElement $xW 'docDefaults')[0])
            $memo = @{}
            $needsRename = $null
            $needsRename = {
                param($id)
                if ($memo.ContainsKey($id)) { return $memo[$id] }
                $memo[$id] = $false
                $se = $srcById[$id]
                if (-not $se) { return $false }
                $r = $false
                if ($outById.ContainsKey($id)) { $r = $defaultsDiffer -or ((& $canon $se) -ne (& $canon $outById[$id])) }
                $based = (Kids $se $xW 'basedOn')[0]
                if (-not $r -and $outById.ContainsKey($id) -and $based -and (& $needsRename (Get-WAttr $based 'val'))) { $r = $true }
                $memo[$id] = $r
                return $r
            }
            $usedIds = @{}; foreach ($k in $outById.Keys) { $usedIds[$k] = $true }
            $usedNames = @{}; foreach ($n in (All $outSty $xW 'name')) { $usedNames[(Get-WAttr $n 'val').ToLower()] = $true }
            foreach ($e in $srcOrder) {
                $id = Get-WAttr $e 'styleId'
                if (& $needsRename $id) {
                    $nid = $id + '_' + ($idx + 1); $k = 2
                    while ($usedIds.ContainsKey($nid) -or $srcById.ContainsKey($nid)) { $nid = $id + '_' + ($idx + 1) + '_' + $k; $k++ }
                    $usedIds[$nid] = $true
                    $styleMap[$id] = $nid
                }
                if ((Get-WAttr $e 'default') -eq '1') { $srcDefault[(Get-WAttr $e 'type')] = $id }
            }
            $srcDD = (Kids $srcSty.DocumentElement $xW 'docDefaults')[0]
            $ddProps = {
                param($tag, $inner)
                if (-not $srcDD) { return ,@() }
                $a = @($srcDD.GetElementsByTagName($tag, $xW))[0]
                if (-not $a) { return ,@() }
                $b = (Kids $a $xW $inner)[0]
                if (-not $b) { return ,@() }
                $r = @(); foreach ($n in $b.ChildNodes) { if ($n.NodeType -eq 'Element') { $r += $n } }
                return ,$r
            }
            $P_ORDER = @('pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange')
            $R_ORDER = @('rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath')
            $STYLE_ORDER = @('name','aliases','basedOn','next','link','autoRedefine','hidden','uiPriority','semiHidden','unhideWhenUsed','qFormat','locked','personal','personalCompose','personalReply','rsid','pPr','rPr','tblPr','trPr','tcPr','tblStylePr')
            $ensureChild = {
                param($st, $name)
                $el = (Kids $st $xW $name)[0]
                if ($el) { return $el }
                $el = New-W $outSty $name
                $pos = [Array]::IndexOf($STYLE_ORDER, $name)
                $before = $null
                foreach ($c in $st.ChildNodes) { if ($c.NodeType -eq 'Element' -and [Array]::IndexOf($STYLE_ORDER, $c.LocalName) -gt $pos) { $before = $c; break } }
                [void]$st.InsertBefore($el, $before)
                return $el
            }
            foreach ($e in $srcOrder) {
                $id = Get-WAttr $e 'styleId'
                $renamed = $styleMap[$id]
                if (-not $renamed -and $outById.ContainsKey($id)) { continue }
                $c = $outSty.ImportNode($e, $true)
                if ($renamed) {
                    Set-WAttr $c 'styleId' $renamed
                    $c.RemoveAttribute('default', $xW)
                    Set-WAttr $c 'customStyle' '1'
                    $nm = (Kids $c $xW 'name')[0]
                    $base = $id; if ($nm) { $base = Get-WAttr $nm 'val' }
                    $base = $base + ' (' + ($idx + 1) + ')'; $name = $base; $k = 2
                    while ($usedNames.ContainsKey($name.ToLower())) { $name = $base + ' ' + $k; $k++ }
                    $usedNames[$name.ToLower()] = $true
                    Set-WAttr (& $ensureChild $c 'name') 'val' $name
                    foreach ($a in (Kids $c $xW 'aliases')) { [void]$c.RemoveChild($a) }
                    if ($defaultsDiffer -and -not (Kids $c $xW 'basedOn').Count) {
                        foreach ($spec in @(@('rPrDefault','rPr'), @('pPrDefault','pPr'))) {
                            if ($spec[1] -eq 'pPr' -and (Get-WAttr $c 'type') -ne 'paragraph') { continue }
                            $props = & $ddProps $spec[0] $spec[1]
                            if (-not $props.Count) { continue }
                            $holder = & $ensureChild $c $spec[1]
                            foreach ($pn in $props) {
                                $have = (Kids $holder $xW $pn.LocalName)[0]
                                if ($have) {
                                    foreach ($a in @($pn.Attributes)) {
                                        if (-not $have.HasAttribute($a.LocalName, $a.NamespaceURI)) { Set-NsAttr $have $a.Prefix $a.NamespaceURI $a.LocalName $a.Value }
                                    }
                                    continue
                                }
                                $order = $R_ORDER; if ($spec[1] -eq 'pPr') { $order = $P_ORDER }
                                $pos = [Array]::IndexOf($order, $pn.LocalName)
                                $before = $null
                                foreach ($ch in $holder.ChildNodes) { if ($ch.NodeType -eq 'Element' -and [Array]::IndexOf($order, $ch.LocalName) -gt $pos) { $before = $ch; break } }
                                [void]$holder.InsertBefore($outSty.ImportNode($pn, $true), $before)
                            }
                        }
                    }
                }
                foreach ($n in @('basedOn','next','link')) {
                    foreach ($r in (Kids $c $xW $n)) { $v = Get-WAttr $r 'val'; if ($styleMap.ContainsKey($v)) { Set-WAttr $r 'val' $styleMap[$v] } }
                }
                if ($numMap) {
                    foreach ($x in @($c.GetElementsByTagName('numId', $xW))) {
                        $v = Get-WAttr $x 'val'
                        if ($numMap.ContainsKey($v)) { Set-WAttr $x 'val' ([string]$numMap[$v]) }
                    }
                }
                [void]$outSty.DocumentElement.AppendChild($c)
            }
            Save-Xml $outSty $out $outStyPath
        }
    }
    # Point source content at the (possibly renamed) source styles. Paragraphs and
    # tables without an explicit style use the document default, which in the
    # merged file would be the base doc's default, so make it explicit.
    $remapStyles = {
        param($root)
        $doc = $root.OwnerDocument
        foreach ($tag in @('pStyle','rStyle','tblStyle')) {
            foreach ($e in @($root.GetElementsByTagName($tag, $xW))) {
                $v = Get-WAttr $e 'val'
                if ($styleMap.ContainsKey($v)) { Set-WAttr $e 'val' $styleMap[$v] }
            }
        }
        $defP = $null; if ($srcDefault['paragraph'] -and $styleMap.ContainsKey($srcDefault['paragraph'])) { $defP = $styleMap[$srcDefault['paragraph']] }
        $defT = $null; if ($srcDefault['table'] -and $styleMap.ContainsKey($srcDefault['table'])) { $defT = $styleMap[$srcDefault['table']] }
        if ($defP) {
            foreach ($p in @($root.GetElementsByTagName('p', $xW))) {
                $pPr = (Kids $p $xW 'pPr')[0]
                if ($pPr -and (Kids $pPr $xW 'pStyle').Count) { continue }
                if (-not $pPr) { $pPr = New-W $doc 'pPr'; [void]$p.InsertBefore($pPr, $p.FirstChild) }
                $ps = New-W $doc 'pStyle'; Set-WAttr $ps 'val' $defP
                [void]$pPr.InsertBefore($ps, $pPr.FirstChild)
            }
        }
        if ($defT) {
            foreach ($t in @($root.GetElementsByTagName('tbl', $xW))) {
                $tblPr = (Kids $t $xW 'tblPr')[0]
                if (-not $tblPr -or (Kids $tblPr $xW 'tblStyle').Count) { continue }
                $ts = New-W $doc 'tblStyle'; Set-WAttr $ts 'val' $defT
                [void]$tblPr.InsertBefore($ts, $tblPr.FirstChild)
            }
        }
    }
    & $remapStyles $srcBody
    foreach ($part in $styledParts) {
        $d = Read-Xml $out $part
        if ($d) { & $remapStyles $d.DocumentElement; Save-Xml $d $out $part }
    }

    # ---- footnotes / endnotes
    foreach ($kind in @('footnotes','endnotes')) {
        $one = $kind.Substring(0, $kind.Length - 1)
        $sRel = & $findRel $srcRels $kind
        if (-not $sRel) { continue }
        $refs = All $srcDoc $xW ($one + 'Reference')
        if (-not $refs.Count) { continue }
        $sPath = Resolve-Part $srcMain $sRel.GetAttribute('Target')
        $sDoc = Read-Xml $src $sPath
        & $remapStyles $sDoc.DocumentElement
        $oRel = & $findRel $outRels $kind
        if (-not $oRel) {
            $np = & $copyPart $sPath
            Save-Xml $sDoc $out $np
            [void](& $addRel $kind (Relative-Part $outMain $np))
            continue
        }
        $oPath = Resolve-Part $outMain $oRel.GetAttribute('Target')
        $oDoc = Read-Xml $out $oPath
        $max = 0; foreach ($e in (Kids $oDoc.DocumentElement $xW $one)) { $v = [int](Get-WAttr $e 'id'); if ($v -gt $max) { $max = $v } }
        $idMap = @{}
        foreach ($n in (Kids $sDoc.DocumentElement $xW $one)) {
            $t = Get-WAttr $n 'type'
            if ($t -and $t -ne 'normal') { continue }
            $c = $oDoc.ImportNode($n, $true)
            $max++
            $idMap[(Get-WAttr $n 'id')] = $max
            Set-WAttr $c 'id' ([string]$max)
            [void]$oDoc.DocumentElement.AppendChild($c)
        }
        foreach ($r in $refs) { $v = Get-WAttr $r 'id'; if ($idMap.ContainsKey($v)) { Set-WAttr $r 'id' ([string]$idMap[$v]) } }
        Save-Xml $oDoc $out $oPath
    }

    # ---- namespaces declared on the source root
    $oRoot = $outDoc.DocumentElement; $sRoot = $srcDoc.DocumentElement
    foreach ($a in @($sRoot.Attributes)) {
        if ($a.Prefix -eq 'xmlns' -and -not $oRoot.HasAttribute($a.Name)) { [void]$oRoot.SetAttribute($a.Name, $a.Value) }
    }
    $sIgn = $sRoot.GetAttribute('Ignorable', $script:NS_MC)
    if ($sIgn) {
        $cur = New-Object System.Collections.Generic.List[string]
        foreach ($x in ($oRoot.GetAttribute('Ignorable', $script:NS_MC) -split '\s+')) { if ($x) { $cur.Add($x) } }
        foreach ($x in ($sIgn -split '\s+')) { if ($x -and -not $cur.Contains($x) -and $oRoot.HasAttribute('xmlns:' + $x)) { $cur.Add($x) } }
        if (-not $oRoot.HasAttribute('xmlns:mc')) { [void]$oRoot.SetAttribute('xmlns:mc', $script:NS_MC) }
        Set-NsAttr $oRoot 'mc' $script:NS_MC 'Ignorable' ($cur -join ' ')
    }

    # ---- sections
    $outSect = (Kids $outBody $xW 'sectPr')[0]
    if ($outSect) {
        [void]$outBody.RemoveChild($outSect)
        $p = New-W $outDoc 'p'; $pPr = New-W $outDoc 'pPr'
        [void]$pPr.AppendChild($outSect); [void]$p.AppendChild($pPr)
        [void]$outBody.AppendChild($p)
    }
    $firstSrcSect = $null
    foreach ($sp in (All $srcDoc $xW 'sectPr')) {
        $pn = $sp.ParentNode.LocalName
        if ($pn -eq 'rPr' -or $pn -eq 'sectPrChange') { continue }
        foreach ($t in (Kids $sp $xW 'type')) { [void]$sp.RemoveChild($t) }
        if (-not $firstSrcSect -and ($pn -eq 'pPr' -or $sp.ParentNode -eq $srcBody)) { $firstSrcSect = $sp }
    }
    if ($firstSrcSect -and $outSect) {
        foreach ($kind in @('header','footer')) {
            $baseRefs = Kids $outSect $xW ($kind + 'Reference')
            if ($baseRefs.Count -and -not (Kids $firstSrcSect $xW ($kind + 'Reference')).Count) {
                $tag = 'ftr'; if ($kind -eq 'header') { $tag = 'hdr' }
                $path = & $uniqueName ('word/' + $kind + '.xml')
                $fp = P $out $path
                [IO.File]::WriteAllText($fp, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<w:' + $tag + ' xmlns:w="' + $xW + '"><w:p/></w:' + $tag + '>', (New-Object System.Text.UTF8Encoding($false)))
                $o = $outCT.CreateElement('Override', $xCT)
                $o.SetAttribute('PartName', '/' + $path)
                $o.SetAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + $kind + '+xml')
                [void]$ctRoot.AppendChild($o)
                $id = & $addRel $kind (Relative-Part $outMain $path)
                $types = @{}
                foreach ($r in $baseRefs) { $types[(Get-WAttr $r 'type')] = $true }
                foreach ($t in $types.Keys) {
                    $ref = New-W $srcDoc ($kind + 'Reference')
                    Set-WAttr $ref 'type' $t
                    Set-NsAttr $ref 'r' $xR 'id' $id
                    [void]$firstSrcSect.InsertBefore($ref, $firstSrcSect.FirstChild)
                }
            }
        }
    }
    foreach ($n in @($srcBody.ChildNodes)) { [void]$outBody.AppendChild($outDoc.ImportNode($n, $true)) }

    Save-Xml $outDoc $out $outMain
    Save-Xml $outRels $out (RelsPathOf $outMain)
    Save-Xml $outCT $out '[Content_Types].xml'
}

# ===== GUI: pick order, convert .doc with Word, merge =====
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:items = New-Object System.Collections.ArrayList   # full paths

function Add-Paths($paths) {
    foreach ($a in $paths) {
        if (Test-Path -LiteralPath $a -PathType Container) {
            $found = Get-ChildItem -LiteralPath $a -File | Where-Object { ($_.Extension -ieq '.doc' -or $_.Extension -ieq '.docx') -and -not $_.Name.StartsWith('~$') } | Sort-Object Name
            foreach ($f in $found) { [void]$script:items.Add($f.FullName) }
        } elseif (Test-Path -LiteralPath $a -PathType Leaf) {
            $ext = [IO.Path]::GetExtension($a)
            if (($ext -ieq '.doc' -or $ext -ieq '.docx') -and -not ([IO.Path]::GetFileName($a)).StartsWith('~$')) { [void]$script:items.Add((Resolve-Path -LiteralPath $a).Path) }
        }
    }
    Refresh-List
}

function Refresh-List([int]$select = -1) {
    $list.BeginUpdate()
    $list.Items.Clear()
    for ($i = 0; $i -lt $script:items.Count; $i++) {
        [void]$list.Items.Add(('{0}.  {1}' -f ($i + 1), [IO.Path]::GetFileName($script:items[$i])))
    }
    $list.EndUpdate()
    if ($select -ge 0 -and $select -lt $list.Items.Count) { $list.SelectedIndex = $select }
    $btnMerge.Enabled = ($script:items.Count -ge 2)
    Refresh-Paper
    $status.Text = if ($script:items.Count -lt 2) { 'Add 2 or more .doc / .docx files.' } else { 'Top file = first page. Drag to reorder.' }
}

function Move-ListItem([int]$from, [int]$to) {
    if ($from -lt 0 -or $to -lt 0 -or $from -ge $script:items.Count -or $to -ge $script:items.Count -or $from -eq $to) { return }
    $x = $script:items[$from]
    $script:items.RemoveAt($from)
    $script:items.Insert($to, $x)
    Refresh-List $to
}

$font = New-Object System.Drawing.Font('Segoe UI', 10)
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Word Merge'
$form.Size = New-Object System.Drawing.Size(560, 460)
$form.MinimumSize = New-Object System.Drawing.Size(440, 340)
$form.StartPosition = 'CenterScreen'
$form.Font = $font
$form.AllowDrop = $true

$hint = New-Object System.Windows.Forms.Label
$hint.Text = 'Drop .doc / .docx files here. Each file keeps its own margins and starts on a new page.'
$hint.Dock = 'Top'; $hint.Height = 44; $hint.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 0)

$list = New-Object System.Windows.Forms.ListBox
$list.Dock = 'Fill'
$list.AllowDrop = $true
$list.IntegralHeight = $false
$list.Font = New-Object System.Drawing.Font('Segoe UI', 11)

$side = New-Object System.Windows.Forms.FlowLayoutPanel
$side.Dock = 'Right'; $side.Width = 120; $side.FlowDirection = 'TopDown'; $side.Padding = New-Object System.Windows.Forms.Padding(8, 0, 8, 0)
function New-Btn($text) { $b = New-Object System.Windows.Forms.Button; $b.Text = $text; $b.Width = 100; $b.Height = 32; $side.Controls.Add($b); return $b }
$btnAdd = New-Btn 'Add files...'
$btnUp = New-Btn 'Move up'
$btnDown = New-Btn 'Move down'
$btnRemove = New-Btn 'Remove'
$btnClear = New-Btn 'Clear'

$paperRow = New-Object System.Windows.Forms.FlowLayoutPanel
$paperRow.Dock = 'Bottom'; $paperRow.Height = 38; $paperRow.Padding = New-Object System.Windows.Forms.Padding(8, 6, 0, 0)
$paperLbl = New-Object System.Windows.Forms.Label
$paperLbl.Text = 'Paper size:'; $paperLbl.AutoSize = $true; $paperLbl.Margin = New-Object System.Windows.Forms.Padding(3, 6, 6, 0)
$paperBox = New-Object System.Windows.Forms.ComboBox
$paperBox.DropDownStyle = 'DropDownList'; $paperBox.Width = 220
$paperBox.Width = 300
# choices: fixed sizes, then "Same as file N" for each file (tracked by path so reordering keeps it)
$script:paperChoice = 'letter'
$script:paperValues = @()
$script:paperRefreshing = $false
function Refresh-Paper {
    if ($script:paperChoice -like 'file:*' -and -not ($script:items -contains $script:paperChoice.Substring(5))) { $script:paperChoice = 'letter' }
    $script:paperRefreshing = $true
    $paperBox.BeginUpdate(); $paperBox.Items.Clear()
    $vals = @('letter', 'legal', 'a4', '')
    [void]$paperBox.Items.AddRange(@('Letter (8.5 x 11 in)', 'Legal (8.5 x 14 in)', 'A4 (210 x 297 mm)', 'Keep original sizes'))
    for ($i = 0; $i -lt $script:items.Count; $i++) {
        $vals += ('file:' + $script:items[$i])
        [void]$paperBox.Items.Add(('Same as file {0}: {1}' -f ($i + 1), [IO.Path]::GetFileName($script:items[$i])))
    }
    $script:paperValues = $vals
    $paperBox.EndUpdate()
    $paperBox.SelectedIndex = [Math]::Max(0, [Array]::IndexOf($vals, $script:paperChoice))
    $script:paperRefreshing = $false
}
$paperBox.Add_SelectedIndexChanged({
    if (-not $script:paperRefreshing -and $paperBox.SelectedIndex -ge 0) { $script:paperChoice = $script:paperValues[$paperBox.SelectedIndex] }
})
$paperRow.Controls.Add($paperLbl); $paperRow.Controls.Add($paperBox)

$bottom = New-Object System.Windows.Forms.Panel
$bottom.Dock = 'Bottom'; $bottom.Height = 56; $bottom.Padding = New-Object System.Windows.Forms.Padding(10, 10, 12, 10)
$status = New-Object System.Windows.Forms.Label
$status.Dock = 'Fill'; $status.TextAlign = 'MiddleLeft'; $status.ForeColor = [System.Drawing.Color]::DimGray
$btnMerge = New-Object System.Windows.Forms.Button
$btnMerge.Text = 'Merge'; $btnMerge.Width = 120; $btnMerge.Dock = 'Right'
$btnMerge.BackColor = [System.Drawing.Color]::FromArgb(43, 91, 215); $btnMerge.ForeColor = [System.Drawing.Color]::White; $btnMerge.FlatStyle = 'Flat'
$bottom.Controls.Add($status); $bottom.Controls.Add($btnMerge)

$mid = New-Object System.Windows.Forms.Panel
$mid.Dock = 'Fill'; $mid.Padding = New-Object System.Windows.Forms.Padding(10, 6, 0, 0)
$mid.Controls.Add($list)
$form.Controls.Add($mid); $form.Controls.Add($side); $form.Controls.Add($paperRow); $form.Controls.Add($bottom); $form.Controls.Add($hint)

# ---- drag & drop: files from Explorer, and reordering inside the list
$script:dragFrom = -1
$list.Add_MouseDown({
    param($s, $e)
    $i = $list.IndexFromPoint($e.Location)
    if ($i -ge 0 -and $e.Button -eq 'Left') {
        $script:dragFrom = $i
        $list.SelectedIndex = $i
        [void]$list.DoDragDrop([string]$i, [System.Windows.Forms.DragDropEffects]::Move)
        $script:dragFrom = -1
    }
})
$dragOver = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' }
    elseif ($script:dragFrom -ge 0) { $e.Effect = 'Move' }
    else { $e.Effect = 'None' }
}
$list.Add_DragEnter($dragOver); $list.Add_DragOver($dragOver)
$form.Add_DragEnter($dragOver); $form.Add_DragOver($dragOver)
$dropFiles = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        Add-Paths ($e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop))
    }
}
$form.Add_DragDrop($dropFiles)
$list.Add_DragDrop({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { & $dropFiles $s $e; return }
    if ($script:dragFrom -lt 0) { return }
    $pt = $list.PointToClient((New-Object System.Drawing.Point($e.X, $e.Y)))
    $to = $list.IndexFromPoint($pt)
    if ($to -lt 0) { $to = $script:items.Count - 1 }
    Move-ListItem $script:dragFrom $to
})

$btnAdd.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Word documents (*.doc;*.docx)|*.doc;*.docx'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog($form) -eq 'OK') { Add-Paths $dlg.FileNames }
})
$btnUp.Add_Click({ $i = $list.SelectedIndex; if ($i -gt 0) { Move-ListItem $i ($i - 1) } })
$btnDown.Add_Click({ $i = $list.SelectedIndex; if ($i -ge 0) { Move-ListItem $i ($i + 1) } })
$btnRemove.Add_Click({ $i = $list.SelectedIndex; if ($i -ge 0) { $script:items.RemoveAt($i); Refresh-List ([Math]::Min($i, $script:items.Count - 1)) } })
$btnClear.Add_Click({ $script:items.Clear(); Refresh-List })

# ---- merge
function Save-AsDocxKeepCompat($doc, [string]$path) {
    # .docx (FileFormat 12) with the document's ORIGINAL compatibility mode,
    # passed as a named argument (positional COM args break in PowerShell 5.1).
    $mode = [int]$doc.CompatibilityMode
    try {
        [void]$doc.GetType().InvokeMember('SaveAs2', [Reflection.BindingFlags]::InvokeMethod, $null, $doc,
            [object[]]@($path, 12, $mode), $null, $null, [string[]]@('FileName', 'FileFormat', 'CompatibilityMode'))
    } catch {
        $doc.SaveAs2($path, 12)
    }
    if ([int]$doc.CompatibilityMode -ne $mode) { try { $doc.SetCompatibilityMode($mode); $doc.Save() } catch {} }
}
function Convert-DocToDocx($word, [string]$path, [string]$outPath) {
    $doc = $word.Documents.Open($path, $false, $true, $false, '')
    try {
        Save-AsDocxKeepCompat $doc $outPath
    } finally { $doc.Close(0) | Out-Null }
}

$btnMerge.Add_Click({
    $first = $script:items[0]
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Word document (*.docx)|*.docx|Word 97-2003 document (*.doc)|*.doc'
    $dlg.FileName = 'merged'
    $dlg.InitialDirectory = [IO.Path]::GetDirectoryName($first)
    $dlg.OverwritePrompt = $true
    if ($dlg.ShowDialog($form) -ne 'OK') { return }
    $outPath = $dlg.FileName
    if ($script:items -contains $outPath) {
        [void][System.Windows.Forms.MessageBox]::Show($form, 'Choose a different name: that file is in the list.', 'Word Merge'); return
    }

    $form.UseWaitCursor = $true; $btnMerge.Enabled = $false
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('wordmerge_' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp | Out-Null
    $word = $null
    try {
        $paths = @()
        for ($i = 0; $i -lt $script:items.Count; $i++) {
            $p = $script:items[$i]
            if ([IO.Path]::GetExtension($p) -ieq '.doc') {
                $status.Text = 'Converting ' + [IO.Path]::GetFileName($p) + ' ...'; $form.Refresh()
                if (-not $word) {
                    try { $word = New-Object -ComObject Word.Application } catch { throw 'Microsoft Word (desktop) is needed to read .doc files.' }
                    $word.Visible = $false; $word.DisplayAlerts = 0
                }
                $conv = Join-Path $tmp ('{0:D3}.docx' -f $i)
                try { Convert-DocToDocx $word $p $conv } catch { throw ('Could not convert ' + [IO.Path]::GetFileName($p) + ': ' + $_.Exception.Message) }
                $paths += $conv
            } else {
                $paths += $p
            }
        }
        if ($word) { $word.Quit() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word); $word = $null }

        $status.Text = 'Merging...'; $form.Refresh()
        $paper = $script:paperChoice
        if ($paper -like 'file:*') {
            $refIdx = [Array]::IndexOf(@($script:items), $paper.Substring(5))
            $paper = Get-PageSizeOf $paths[$refIdx]
            if (-not $paper) { $paper = '' }
        }
        $asDoc = ([IO.Path]::GetExtension($outPath) -ieq '.doc')
        if (-not $asDoc) {
            Merge-DocxFiles $paths $outPath $paper
        } else {
            # merge to a temporary .docx, then let Word save it as .doc
            $tmpDocx = Join-Path $tmp 'merged.docx'
            Merge-DocxFiles $paths $tmpDocx $paper
            $status.Text = 'Saving as .doc ...'; $form.Refresh()
            if (-not $word) {
                try { $word = New-Object -ComObject Word.Application } catch { throw 'Microsoft Word (desktop) is needed to save as .doc.' }
                $word.Visible = $false; $word.DisplayAlerts = 0
            }
            $d = $word.Documents.Open($tmpDocx, $false, $true, $false, '')
            try {
                if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
                $d.SaveAs2($outPath, 0)   # 0 = Word 97-2003 .doc
            } finally { $d.Close(0) | Out-Null }
            $word.Quit() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word); $word = $null
        }
        $status.Text = 'Done: ' + [IO.Path]::GetFileName($outPath)
        $r = [System.Windows.Forms.MessageBox]::Show($form, "Saved:`n$outPath`n`nOpen it now?", 'Word Merge', 'YesNo', 'Information')
        if ($r -eq 'Yes') { Start-Process -FilePath $outPath }
    } catch {
        $status.Text = 'Failed.'
        [void][System.Windows.Forms.MessageBox]::Show($form, "Merge failed:`n" + $_.Exception.Message, 'Word Merge', 'OK', 'Error')
    } finally {
        if ($word) { try { $word.Quit() | Out-Null } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
        $form.UseWaitCursor = $false; $btnMerge.Enabled = ($script:items.Count -ge 2)
    }
})

Refresh-List
if ($args.Count) { Add-Paths $args }
[void]$form.ShowDialog()
