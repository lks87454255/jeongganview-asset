<#
  민요채보 생성 스크립트 (Windows PowerShell 5.1, UTF-8 BOM 으로 저장할 것)

  폴더 구조 (jeongganview-asset 레포):
    sheets-index.json                    ← 앱 목록 (generate_sheets_index.py 로 다시 만드는 것을 권장)
    sheets\민요\                         ← $Root (기본값: 이 스크립트의 상위 폴더)
      csv\{곡}.csv                       ← 앱 CSV v2 (util/JeongganCsv.kt)
      원본\민요채보-1~3\*.png|jpg         ← 채보 스캔 (CSV '원본' 행은 원본\ 기준 상대 경로)
      검수\{곡}.xlsx, 검수\_목록.xlsx      ← 검수 엑셀
      _작업\tsv\*.tsv                     ← 입력 (판독규격.md 의 TSV v2)
      _작업\생성리포트.txt

  규칙
   - #이어짐 으로 연결된 이미지들은 한 곡(한 문서)의 연속 페이지.
   - 이미지 한 장 = CSV 1페이지. 한 장이 10줄을 넘으면 10줄씩 다음 페이지로 나눔.
   - 줄 n(페이지 안 1..10) → 앱 내부 열: 율명 = 2n(대), 가사 = 2n-1(소).
   - 율명 칸은 비슷한 글자를 대표 글자로 바꾸고(app JeongganSymbols.ALIASES 와 동일),
     율명·팔레트 기호가 아닌 글자는 "확인 필요"로 보고한다.

  사용 : powershell -NoProfile -ExecutionPolicy Bypass -File build_all.ps1
#>
param([string]$Root = (Split-Path -Parent $PSScriptRoot))   # sheets\민요
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$utf8 = New-Object System.Text.UTF8Encoding($false)

$TsvDir     = Join-Path $Root '_작업\tsv'
$ReviewDir  = Join-Path $Root '검수'
$RepoDir    = Split-Path -Parent (Split-Path -Parent $Root)   # jeongganview-asset
$CsvDir     = Join-Path $Root 'csv'
$ReportPath = Join-Path $Root '_작업\생성리포트.txt'
$Category   = '민요'
$MaxLinesPerPage = 10
$Beats = @('4/4', '3/4', '2/4', '정악')

# ───────────────────── 기호표 (앱 Yulmyeong.kt / JeongganSymbols.kt 와 동일) ─────────────────────
$YUL = @(
  '㣴','㣕','㣖','㣣','㣨','㣡','㣸','㣩','𢓡','㣮','㣳','㣹',
  '僙','㐲','㑀','俠','㑬','㑖','𠐭','㑣','侇','㑲','㒇','㒣',
  '黃','大','太','夾','姑','仲','㽔','林','夷','南','無','應',
  '潢','汏','汰','浹','㴌','㳞','㶋','淋','洟','湳','潕','㶐',
  '㶂','𣴘','㳲','㴺','㵈','㴢','㶙','㵉','㴣','㵜','㶃','㶝')
$MARKS = @('―','△','И','ﾉ','^','⌝','ㅋ','⌞','է','⊏','⊔','·','○','⁚','‹','／')
$KNOWN = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
foreach ($s in ($YUL + $MARKS)) { [void]$KNOWN.Add($s) }
$ALIAS = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
function Add-Alias([string]$to, [string[]]$from) { foreach ($f in $from) { $ALIAS[$f] = $to } }
Add-Alias '―' @('-', '‐', '–', '—', '−', 'ㅡ')
Add-Alias '△' @('▵')
Add-Alias '·' @('.', 'ㆍ', '•', '∙', '⋅')
Add-Alias '○' @('o', 'O', '◯', '〇')
Add-Alias '⁚' @(':', '︰', '：')
Add-Alias '‹' @('<', '〈', '＜')
Add-Alias '／' @('/', '∕')
if ($YUL.Count -ne 60) { throw "율명표 60자 아님: $($YUL.Count)" }

function Get-CodePoints([string]$s) {
  $list = New-Object System.Collections.Generic.List[string]
  $i = 0
  while ($i -lt $s.Length) {
    if ([char]::IsHighSurrogate($s[$i]) -and $i + 1 -lt $s.Length) { $list.Add($s.Substring($i, 2)); $i += 2 }
    else { $list.Add($s.Substring($i, 1)); $i++ }
  }
  return , $list
}

# TSV 율명(| 구분) → 앱 셀 텍스트(\n 구분, 대표 글자)
function Convert-Pitch([string]$raw) {
  if ([string]::IsNullOrWhiteSpace($raw)) { return '' }
  $parts = foreach ($seg in $raw.Split('|')) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($cp in (Get-CodePoints $seg.Trim())) {
      if ([string]::IsNullOrWhiteSpace($cp)) { continue }
      if ($ALIAS.ContainsKey($cp)) { [void]$sb.Append($ALIAS[$cp]) } else { [void]$sb.Append($cp) }
    }
    $sb.ToString()
  }
  $parts = @($parts)
  while ($parts.Count -gt 0 -and $parts[0] -eq '') { $parts = @($parts | Select-Object -Skip 1) }
  while ($parts.Count -gt 0 -and $parts[-1] -eq '') { $parts = @($parts | Select-Object -First ($parts.Count - 1)) }
  return ($parts -join "`n")
}

function Get-Unknown([string]$text) {
  $out = New-Object System.Collections.Generic.List[string]
  foreach ($cp in (Get-CodePoints $text)) {
    if (-not [string]::IsNullOrWhiteSpace($cp) -and -not $KNOWN.Contains($cp) -and -not $out.Contains($cp)) { $out.Add($cp) }
  }
  return , $out
}

# ───────────────────── TSV 읽기 ─────────────────────
$Problems = New-Object System.Collections.Generic.List[string]   # 리포트용

function Read-Tsv([IO.FileInfo]$file) {
  $meta = @{}
  $cells = New-Object System.Collections.Generic.List[object]
  $header = $false
  $n = 0
  foreach ($ln in [IO.File]::ReadAllLines($file.FullName, $utf8)) {
    $n++
    if ($ln.Trim() -eq '') { continue }
    if ($ln.StartsWith('#')) {
      $kv = $ln.Substring(1).Split("`t", 2)
      $meta[$kv[0].Trim()] = $(if ($kv.Length -gt 1) { $kv[1].Trim() } else { '' })
      continue
    }
    $f = $ln.Split("`t")
    if (-not $header -and $f[0].Trim() -eq '줄') { $header = $true; continue }
    $jul = 0; $jg = 0
    if ($f.Length -lt 2 -or -not [int]::TryParse($f[0].Trim(), [ref]$jul) -or -not [int]::TryParse($f[1].Trim(), [ref]$jg) -or $jul -lt 1 -or $jg -lt 1) {
      $Problems.Add("[형식] $($file.Name) ${n}행 읽을 수 없음: $ln"); continue
    }
    $raw = $(if ($f.Length -gt 2) { $f[2].Trim() } else { '' })
    $lyr = $(if ($f.Length -gt 3) { $f[3].Trim() } else { '' })
    $note = $(if ($f.Length -gt 4) { ($f[4..($f.Length - 1)] -join ' ').Trim() } else { '' })
    $pitch = Convert-Pitch $raw
    if (-not $pitch -and -not $lyr) { continue }   # 빈 정간 행(검사용)은 건너뜀
    $cells.Add([pscustomobject]@{ Jul = $jul; Jg = $jg; Raw = $raw; Pitch = $pitch; Lyric = $lyr; Note = $note; Unknown = (Get-Unknown $pitch) })
  }
  $folder = $file.BaseName.Split('_', 2)[0]
  $stem = $file.BaseName.Split('_', 2)[1]
  $g = { param($k) if ($meta.ContainsKey($k)) { $meta[$k] } else { '' } }
  $rows = 0; [void][int]::TryParse((& $g '행수'), [ref]$rows)
  $lines = 0; [void][int]::TryParse((& $g '줄수'), [ref]$lines)
  $maxJg = ($cells | Measure-Object Jg -Maximum).Maximum; if (-not $maxJg) { $maxJg = 0 }
  $maxJul = ($cells | Measure-Object Jul -Maximum).Maximum; if (-not $maxJul) { $maxJul = 0 }
  # 같은 (줄,정간)이 두 번 나오면 뒤의 것을 버리고 보고
  $seen = @{}
  $dedup = New-Object System.Collections.Generic.List[object]
  foreach ($c in $cells) {
    $k = "$($c.Jul)-$($c.Jg)"
    if ($seen.ContainsKey($k)) { $Problems.Add("[중복] $($file.Name) 줄$($c.Jul) 정간$($c.Jg) 두 번 — 뒤의 것 무시"); continue }
    $seen[$k] = 1; $dedup.Add($c)
  }
  [pscustomobject]@{
    File = $file.Name; Folder = $folder; Stem = $stem
    Title = (& $g '제목'); Sub = (& $g '부제'); Key = (& $g '키'); Page = (& $g '쪽')
    Beat = (& $g '박자'); Rows = [math]::Max($rows, $maxJg); RowsMeta = $rows; LinesMeta = $lines; MaxJul = $maxJul
    Source = (& $g '원본'); Next = (& $g '이어짐'); Memo = (& $g '메모')
    Cells = $dedup
    PresentLines = @($dedup | ForEach-Object Jul | Sort-Object -Unique).Count
  }
}

$tsvs = @{}
foreach ($f in Get-ChildItem -LiteralPath $TsvDir -Filter *.tsv | Sort-Object Name) { $tsvs[$f.Name] = Read-Tsv $f }

# ───────────────────── 이어짐 → 곡(문서) 묶기 ─────────────────────
function Resolve-Next([string]$v) {
  if ([string]::IsNullOrWhiteSpace($v)) { return $null }
  $v = $v.Trim(); if (-not $v.EndsWith('.tsv')) { $v += '.tsv' }
  foreach ($k in $tsvs.Keys) { if ($k -ieq $v) { return $k } }
  return $null
}
$nextOf = @{}; $hasPrev = @{}
foreach ($t in $tsvs.Values) {
  $n = Resolve-Next $t.Next
  if ($t.Next -and -not $n) { $Problems.Add("[이어짐] $($t.File) → '$($t.Next)' 파일 없음 — 무시") }
  elseif ($n -eq $t.File) { $Problems.Add("[이어짐] $($t.File) 자기 자신 — 무시") }
  elseif ($n) {
    if ($hasPrev.ContainsKey($n)) { $Problems.Add("[이어짐] $n 앞장이 둘($($hasPrev[$n]), $($t.File)) — 뒤의 것 무시") }
    else { $nextOf[$t.File] = $n; $hasPrev[$n] = $t.File }
  }
}
$docs = New-Object System.Collections.Generic.List[object]
$used = @{}
$broken = @{}
function Add-Doc($head) {
  $chain = New-Object System.Collections.Generic.List[object]
  $cur = $head
  while ($cur -and -not $used.ContainsKey($cur)) {
    $t = $tsvs[$cur]
    if ($chain.Count -and ($t.Beat.Trim() -ne $chain[0].Beat.Trim() -or $t.Rows -ne $chain[0].Rows)) {
      $Problems.Add("[이어짐] $($chain[-1].File) → $cur : 박자/행수 다름($($chain[0].Beat)/$($chain[0].Rows) ≠ $($t.Beat)/$($t.Rows)) — 따로 곡으로 분리")
      $broken[$cur] = 1
      break
    }
    $used[$cur] = 1; $chain.Add($t); $cur = $nextOf[$cur]
  }
  $docs.Add([pscustomobject]@{ Images = $chain })
}
foreach ($k in ($tsvs.Keys | Sort-Object)) { if (-not $hasPrev.ContainsKey($k)) { Add-Doc $k } }
foreach ($k in ($tsvs.Keys | Sort-Object)) {
  if (-not $used.ContainsKey($k)) {
    if (-not $broken.ContainsKey($k)) { $Problems.Add("[이어짐] $k 순환 연결 — 따로 처리") }
    Add-Doc $k
  }
}

# ───────────────────── 문서 내용 계산 ─────────────────────
function Get-BaseTitle($t) {
  if ($t.Title) { return $t.Title.Trim() }
  return "제목없음 ($($t.Stem -replace '\.jpg$', ''))"
}
function Clean-FileName([string]$s) {
  $s = ($s -replace '[\\/:*?"<>|\r\n\t]', ' ' -replace '\s+', ' ').Trim().TrimEnd('.')
  return $s.Normalize([Text.NormalizationForm]::FormC)
}
function Clean-Tag([string]$s) { return (Clean-FileName ($s -replace '[()\[\]]', ' ')) }

foreach ($d in $docs) {
  $h = $d.Images[0]
  $d | Add-Member Title (Get-BaseTitle $h)
  $d | Add-Member Sub $h.Sub
  $d | Add-Member Key (@($d.Images | ForEach-Object Key | Where-Object { $_ }) | Select-Object -First 1)
  $d | Add-Member PageNo ((@($d.Images | ForEach-Object Page | Where-Object { $_ })) -join ', ')
  $beat = $h.Beat.Trim()
  if ($beat -notin $Beats) { $Problems.Add("[박자] $($h.File) '$beat' 미지원 → 3/4"); $beat = '3/4' }
  foreach ($im in $d.Images) { if ($im.Beat.Trim() -ne $h.Beat.Trim()) { $Problems.Add("[박자] $($im.File) '$($im.Beat)' ≠ 첫 장 '$($h.Beat)' — 첫 장 기준") } }
  $d | Add-Member Beat $beat
  $rows = ($d.Images | Measure-Object Rows -Maximum).Maximum
  if (-not $rows -or $rows -lt 1) { $rows = 12 }
  $d | Add-Member RowCount ([int]$rows)

  # 페이지: 이미지마다, 10줄 넘으면 나눔
  $pages = New-Object System.Collections.Generic.List[object]
  foreach ($im in $d.Images) {
    $nPages = [math]::Max(1, [math]::Ceiling([math]::Max($im.MaxJul, 1) / $MaxLinesPerPage))
    for ($p = 0; $p -lt $nPages; $p++) {
      $cells = @{}
      $list = New-Object System.Collections.Generic.List[object]
      foreach ($c in $im.Cells) {
        if ([math]::Floor(($c.Jul - 1) / $MaxLinesPerPage) -ne $p) { continue }
        $local = (($c.Jul - 1) % $MaxLinesPerPage) + 1
        if ($c.Pitch) { $cells["$(2 * $local),$($c.Jg)"] = $c.Pitch }
        if ($c.Lyric) { $cells["$(2 * $local - 1),$($c.Jg)"] = $c.Lyric }
        $list.Add([pscustomobject]@{ Cell = $c; Local = $local })
      }
      $linesOnPage = [math]::Min($MaxLinesPerPage, [math]::Max(1, $im.MaxJul - $p * $MaxLinesPerPage))
      $pages.Add([pscustomobject]@{ Image = $im; Part = $p + 1; Parts = $nPages; Cells = $cells; Items = $list; Lines = $linesOnPage })
    }
  }
  $d | Add-Member Pages $pages

  $all = @($d.Images | ForEach-Object { $_.Cells })
  $d | Add-Member CellCount $all.Count
  $d | Add-Member NoteCount @($all | Where-Object { $_.Note }).Count
  $d | Add-Member UnknownCount ([int](($all | ForEach-Object { $_.Unknown.Count } | Measure-Object -Sum).Sum))
  $present = ($d.Images | ForEach-Object { $_.PresentLines } | Measure-Object -Sum).Sum
  $sparse = $all.Count -lt (0.3 * $present * $d.RowCount)
  $ratio = $(if ($all.Count) { $d.NoteCount / $all.Count } else { 1 })
  $memo = ($d.Images | ForEach-Object Memo) -join ' '
  $reasons = @()
  if ($d.UnknownCount) { $reasons += "팔레트 외 글자 $($d.UnknownCount)" }
  if ($ratio -ge 0.15) { $reasons += ('비고 {0:P0}' -f $ratio) }
  if ($sparse) { $reasons += '정간 적음(누락 의심)' }
  if ($memo -match '자진모리|근사|가까운') { $reasons += '박자 근사' }
  if ($memo -match '두 곡') { $reasons += '한 장에 두 곡' }
  if (-not $h.Title) { $reasons += '제목 없음' }
  if ($all.Count -eq 0) { $reasons += '내용 없음' }
  $grade = $(if ($d.UnknownCount -eq 0 -and $ratio -lt 0.15 -and -not $sparse -and $all.Count) { '상' }
             elseif ($d.UnknownCount -le 10 -and $ratio -lt 0.5 -and -not $sparse -and $all.Count) { '중' } else { '하' })
  $d | Add-Member Grade $grade
  $d | Add-Member Reasons ($reasons -join ', ')
}

# 이름 정하기 (겹치면 부제 → 채보 폴더 → 파일명 순으로 구분)
foreach ($grp in ($docs | Group-Object { Clean-FileName $_.Title })) {
  if ($grp.Count -eq 1) { $grp.Group[0] | Add-Member Name $grp.Name; continue }
  foreach ($d in $grp.Group) {
    $sub = Clean-Tag $d.Sub
    $sameSub = @($grp.Group | Where-Object { (Clean-Tag $_.Sub) -eq $sub }).Count
    $tag = $(if ($sub -and $sameSub -eq 1 -and $sub -notmatch '^\d+$') { $sub } else { "채보$($d.Images[0].Folder)" })
    $d | Add-Member Name "$($grp.Name) ($tag)"
  }
}
foreach ($grp in ($docs | Group-Object Name | Where-Object Count -gt 1)) {
  foreach ($d in $grp.Group) { $d.Name = Clean-FileName "$($d.Name) $($d.Images[0].Stem)" }
}
$docs = @($docs | Sort-Object Name)

# ───────────────────── CSV v2 ─────────────────────
function Quote-Csv([string]$v) {
  if ($v.Length -gt 0 -and ('=+-@'.IndexOf($v[0]) -ge 0 -or ($v.Length -gt 1 -and $v[0] -eq "'" -and "=+-@'".IndexOf($v[1]) -ge 0))) { $v = "'" + $v }
  if ($v.IndexOfAny([char[]]@(',', '"', "`n", "`r")) -ge 0 -or ($v.Length -gt 0 -and ([char]::IsWhiteSpace($v[0]) -or [char]::IsWhiteSpace($v[-1])))) {
    return '"' + $v.Replace('"', '""') + '"'
  }
  return $v
}
function HeaderName([int]$c) { "$([math]::Floor(($c + 1) / 2))열($(if ($c % 2 -eq 0) { '대' } else { '소' }))" }
function BeatLabel([string]$b) { if ($b.Contains('/')) { "${b}박자" } else { $b } }

function Build-Csv($d) {
  $rows = New-Object System.Collections.Generic.List[object]
  $rows.Add(@('형식', '정간보CSV', '2'))
  $rows.Add(@('타이틀', $d.Title))
  $rows.Add(@('박자', (BeatLabel $d.Beat)))
  $rows.Add(@('행수', "$($d.RowCount)"))
  if ($d.Sub) { $rows.Add(@('부제', $d.Sub)) }
  if ($d.Key) { $rows.Add(@('키', $d.Key)) }
  if ($d.PageNo) { $rows.Add(@('쪽', $d.PageNo)) }
  $rows.Add(@('원본', (($d.Images | ForEach-Object Source) -join ' / ')))
  $rows.Add(@('상태', '판독 초안(검수 전)'))
  $header = @('정간번호') + (20..1 | ForEach-Object { HeaderName $_ })
  for ($p = 0; $p -lt $d.Pages.Count; $p++) {
    $rows.Add(@('페이지', "$($p + 1)"))
    $rows.Add($header)
    for ($r = 1; $r -le $d.RowCount; $r++) {
      $line = @("$r")
      foreach ($c in 20..1) { $v = $d.Pages[$p].Cells["$c,$r"]; $line += $(if ($v) { $v } else { '' }) }
      $rows.Add($line)
    }
  }
  $text = (($rows | ForEach-Object { ($_ | ForEach-Object { Quote-Csv $_ }) -join ',' }) -join "`r`n") + "`r`n"
  return $text
}

# 검증용 RFC 4180 파서 (앱 JeongganCsv.parseRecords 와 같은 규칙)
function Parse-Csv([string]$text) {
  $records = New-Object System.Collections.Generic.List[object]
  $row = New-Object System.Collections.Generic.List[string]
  $field = New-Object System.Text.StringBuilder
  $inQ = $false; $i = 0
  while ($i -lt $text.Length) {
    $ch = $text[$i]
    if ($inQ) {
      if ($ch -eq '"') { if ($i + 1 -lt $text.Length -and $text[$i + 1] -eq '"') { [void]$field.Append('"'); $i++ } else { $inQ = $false } }
      else { [void]$field.Append($ch) }
    } else {
      if ($ch -eq '"' -and $field.Length -eq 0) { $inQ = $true }
      elseif ($ch -eq ',') { $row.Add($field.ToString()); [void]$field.Clear() }
      elseif ($ch -eq "`r" -or $ch -eq "`n") {
        $row.Add($field.ToString()); [void]$field.Clear()
        if (@($row | Where-Object { $_ }).Count) { $records.Add($row.ToArray()) }
        $row = New-Object System.Collections.Generic.List[string]
        if ($ch -eq "`r" -and $i + 1 -lt $text.Length -and $text[$i + 1] -eq "`n") { $i++ }
      } else { [void]$field.Append($ch) }
    }
    $i++
  }
  return , $records
}
function Unguard([string]$v) { if ($v.Length -gt 1 -and $v[0] -eq "'" -and "=+-@'".IndexOf($v[1]) -ge 0) { $v.Substring(1) } else { $v } }

function Test-Csv($d, [string]$text) {
  $recs = Parse-Csv $text
  $errs = @()
  $nHead = 0; $nPage = 0
  foreach ($rec in $recs) { if ($rec[0] -eq '정간번호') { $nHead++ } elseif ($rec[0] -eq '페이지') { $nPage++ } }
  if ($nHead -ne $d.Pages.Count) { $errs += "정간번호 행 $nHead ≠ 페이지 $($d.Pages.Count)" }
  if ($nPage -ne $d.Pages.Count) { $errs += "페이지 행 $nPage ≠ 페이지 $($d.Pages.Count)" }
  foreach ($rec in $recs) { foreach ($cell in $rec) { if ($cell.Length -and '=+-@'.IndexOf($cell[0]) -ge 0) { $errs += "수식 위험 칸: $cell" } } }
  $page = 1; $back = 0
  foreach ($rec in $recs) {
    if ($rec[0] -eq '페이지') { $page = [int]$rec[1]; continue }
    $r = 0; if (-not [int]::TryParse($rec[0], [ref]$r)) { continue }
    for ($i = 1; $i -lt $rec.Count; $i++) {
      $v = Unguard $rec[$i].Trim(); if (-not $v) { continue }
      $c = 21 - $i
      if ($d.Pages[$page - 1].Cells["$c,$r"] -ne $v) { $errs += "왕복 불일치 p$page 열$c 행$r" }
      $back++
    }
  }
  $expect = ($d.Pages | ForEach-Object { $_.Cells.Count } | Measure-Object -Sum).Sum
  if ($back -ne $expect) { $errs += "칸 수 불일치 $back ≠ $expect" }
  return , $errs
}

# ───────────────────── xlsx ─────────────────────
$STYLES = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/><family val="2"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/><family val="2"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF8CBAD"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@
# 스타일 번호
$S_HEAD = 1; $S_CELL = 2; $S_NOTE = 3; $S_LEFT = 4; $S_BAD = 5; $S_TITLE = 6

function Esc([string]$s) {
  if ($null -eq $s) { return '' }
  $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
  return [System.Security.SecurityElement]::Escape($s)
}
function ColName([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][math]::Floor(($n - 1) / 26) }; $s }

function New-Sheet([string]$name) {
  [pscustomobject]@{ Name = $name; Rows = (New-Object System.Collections.Generic.List[object]); Widths = @{}; Merges = (New-Object System.Collections.Generic.List[string]); FreezeRows = 0; FreezeCols = 0; Heights = @{}; Filter = $null }
}
# 행 추가: 값 배열 + 스타일(하나 또는 배열)
function Add-Row($sheet, [object[]]$values, $style = $S_CELL) {
  $cells = for ($i = 0; $i -lt $values.Count; $i++) {
    $st = $(if ($style -is [array]) { $style[$i] } else { $style })
    , @($values[$i], $st)
  }
  $sheet.Rows.Add(@($cells))
}
function Sheet-Xml($sh) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
  if ($sh.FreezeRows -or $sh.FreezeCols) {
    $tl = "$(ColName ($sh.FreezeCols + 1))$($sh.FreezeRows + 1)"
    $attr = ''
    if ($sh.FreezeCols) { $attr += " xSplit=`"$($sh.FreezeCols)`"" }
    if ($sh.FreezeRows) { $attr += " ySplit=`"$($sh.FreezeRows)`"" }
    $pane = $(if ($sh.FreezeRows -and $sh.FreezeCols) { 'bottomRight' } elseif ($sh.FreezeRows) { 'bottomLeft' } else { 'topRight' })
    [void]$sb.Append("<sheetViews><sheetView workbookViewId=`"0`"><pane$attr topLeftCell=`"$tl`" activePane=`"$pane`" state=`"frozen`"/></sheetView></sheetViews>")
  }
  if ($sh.Widths.Count) {
    [void]$sb.Append('<cols>')
    foreach ($c in ($sh.Widths.Keys | Sort-Object { [int]$_ })) { [void]$sb.Append("<col min=`"$c`" max=`"$c`" width=`"$($sh.Widths[$c])`" customWidth=`"1`"/>") }
    [void]$sb.Append('</cols>')
  }
  [void]$sb.Append('<sheetData>')
  for ($r = 0; $r -lt $sh.Rows.Count; $r++) {
    $rn = $r + 1
    $ht = $(if ($sh.Heights.ContainsKey($rn)) { " ht=`"$($sh.Heights[$rn])`" customHeight=`"1`"" } else { '' })
    [void]$sb.Append("<row r=`"$rn`"$ht>")
    $row = $sh.Rows[$r]
    for ($c = 0; $c -lt $row.Count; $c++) {
      $cell = $row[$c]; if ($null -eq $cell) { continue }
      $v = $cell[0]; $st = $cell[1]; $ref = "$(ColName ($c + 1))$rn"
      if ($null -eq $v -or "$v" -eq '') { [void]$sb.Append("<c r=`"$ref`" s=`"$st`"/>") }
      elseif ($v -is [int] -or $v -is [long] -or $v -is [double]) { [void]$sb.Append("<c r=`"$ref`" s=`"$st`"><v>$v</v></c>") }
      else { [void]$sb.Append("<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Esc "$v")</t></is></c>") }
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData>')
  if ($sh.Filter) { [void]$sb.Append("<autoFilter ref=`"$($sh.Filter)`"/>") }
  if ($sh.Merges.Count) {
    [void]$sb.Append("<mergeCells count=`"$($sh.Merges.Count)`">")
    foreach ($m in $sh.Merges) { [void]$sb.Append("<mergeCell ref=`"$m`"/>") }
    [void]$sb.Append('</mergeCells>')
  }
  [void]$sb.Append('</worksheet>')
  return $sb.ToString()
}
function Save-Xlsx([string]$path, $sheets) {
  $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
  $wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
  $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  for ($i = 1; $i -le $sheets.Count; $i++) {
    $ct += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
    $wb += "<sheet name=`"$(Esc $sheets[$i - 1].Name)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
    $rels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
  }
  $ct += '</Types>'; $wb += '</sheets></workbook>'
  $rels += "<Relationship Id=`"rId$($sheets.Count + 1)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>"
  $root = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
  if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
  $fs = [IO.File]::Open($path, [IO.FileMode]::CreateNew)
  $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    $add = {
      param($name, $text)
      $e = $zip.CreateEntry($name)
      $w = New-Object IO.StreamWriter($e.Open(), $utf8)
      $w.Write($text); $w.Dispose()
    }
    & $add '[Content_Types].xml' $ct
    & $add '_rels/.rels' $root
    & $add 'xl/workbook.xml' $wb
    & $add 'xl/_rels/workbook.xml.rels' $rels
    & $add 'xl/styles.xml' $STYLES
    for ($i = 1; $i -le $sheets.Count; $i++) { & $add "xl/worksheets/sheet$i.xml" (Sheet-Xml $sheets[$i - 1]) }
  } finally { $zip.Dispose(); $fs.Dispose() }
}

function Cell-Style($c) { if ($c.Unknown.Count) { $S_BAD } elseif ($c.Note) { $S_NOTE } else { $S_CELL } }

function Build-ReviewBook($d) {
  $sheets = New-Object System.Collections.Generic.List[object]

  # 정보
  $info = New-Sheet '정보'; $info.Widths[1] = 16; $info.Widths[2] = 90
  Add-Row $info @($d.Title, '') $S_TITLE
  $kv = @(
    @('CSV 파일', "sheets/$Category/$($d.Name).csv"), @('부제', $d.Sub), @('키', $d.Key), @('쪽', $d.PageNo),
    @('박자 / 행수', "$($d.Beat) / $($d.RowCount)정간"), @('페이지', "$($d.Pages.Count)"),
    @('정간 수', "$($d.CellCount)"), @('비고(노란 칸)', "$($d.NoteCount)"), @('팔레트 외 글자(주황 칸)', "$($d.UnknownCount)"),
    @('판독 신뢰도', $d.Grade), @('확인 필요', $d.Reasons))
  foreach ($x in $kv) { Add-Row $info @($x[0], $x[1]) @($S_HEAD, $S_LEFT) }
  Add-Row $info @('', '') 0
  Add-Row $info @('원본 이미지', '판독 메모') $S_HEAD
  foreach ($im in $d.Images) { Add-Row $info @($im.Source, $im.Memo) @($S_LEFT, $S_LEFT) }
  Add-Row $info @('', '') 0
  Add-Row $info @('범례', '') $S_TITLE
  foreach ($x in @(
      @('줄 순서', '원본처럼 오른쪽이 1줄. 악보 시트도 같은 배치. 한 장이 10줄을 넘으면 10줄씩 다음 페이지'),
      @('율명 칸', '정간 안에서 위→아래 순서, 칸 안 줄바꿈으로 구분 (앱과 같음)'),
      @('기호', '― 연음 · △ 쉼 · 점 · ○ 동그라미 · ⁚ 두 점 · ‹ 정간 경계 표시 · ／ 사선 (· ○ ⁚ ‹ ／ 는 화면 표시 전용)'),
      @('노란 칸', '판독자가 비고를 단 칸 — 원본과 대조 필요'),
      @('주황 칸', '율명·기호표에 없는 글자 (? 포함) — 반드시 수정'),
      @('수정 방법', 'CSV를 엑셀/구글시트로 열어 고친 뒤 CSV(UTF-8)로 저장. 이 엑셀은 검수용 보기'))) {
    Add-Row $info @($x[0], $x[1]) @($S_HEAD, $S_LEFT)
  }
  $sheets.Add($info)

  # 정간목록
  $ls = New-Sheet '정간목록'; $ls.FreezeRows = 1
  $w = @(6, 30, 6, 6, 6, 5, 14, 14, 12, 34, 10); for ($i = 0; $i -lt $w.Count; $i++) { $ls.Widths[$i + 1] = $w[$i] }
  Add-Row $ls @('페이지', '원본 이미지', '줄', '앱 줄', '정간', '박', '율명(판독)', '율명(앱)', '가사', '비고', '확인 필요 글자') $S_HEAD
  $beatSize = $(switch ($d.Beat) { '4/4' { 4 } '2/4' { 2 } '3/4' { 3 } default { 0 } })
  for ($p = 0; $p -lt $d.Pages.Count; $p++) {
    $pg = $d.Pages[$p]
    foreach ($it in ($pg.Items | Sort-Object { $_.Cell.Jul }, { $_.Cell.Jg })) {
      $c = $it.Cell; $st = Cell-Style $c
      $bak = $(if ($beatSize) { [int][math]::Floor(($c.Jg - 1) / $beatSize) + 1 } else { '' })
      Add-Row $ls @(($p + 1), $pg.Image.Source, $c.Jul, $it.Local, $c.Jg, $bak, $c.Raw, $c.Pitch, $c.Lyric, $c.Note, ($c.Unknown -join ' ')) `
        @($S_CELL, $S_LEFT, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $st, $st, $st, $S_LEFT, $(if ($c.Unknown.Count) { $S_BAD } else { $S_CELL }))
    }
  }
  $ls.Filter = "A1:K$($ls.Rows.Count)"
  $sheets.Add($ls)

  # 악보 1..N (원본 배치: 오른쪽 = 1줄)
  for ($p = 0; $p -lt $d.Pages.Count; $p++) {
    $pg = $d.Pages[$p]; $L = $pg.Lines
    $sh = New-Sheet "악보 $($p + 1)"; $sh.FreezeRows = 2; $sh.FreezeCols = 1
    $sh.Widths[1] = 6
    for ($k = 1; $k -le $L; $k++) { $pc = 2 + ($L - $k) * 2; $sh.Widths[$pc] = 9; $sh.Widths[$pc + 1] = 6 }
    $title = "$($d.Title) — $($pg.Image.Source)$(if ($pg.Parts -gt 1) { " ($($pg.Part)/$($pg.Parts))" })"
    $r1 = @($title) + @(1..($L * 2) | ForEach-Object { '' })
    $r2 = @('정간') + @(1..($L * 2) | ForEach-Object { '' })
    $s2 = @($S_HEAD) + @(1..($L * 2) | ForEach-Object { $S_HEAD })
    for ($k = 1; $k -le $L; $k++) {
      $pc = 2 + ($L - $k) * 2
      $r2[$pc - 1] = "$(($pg.Part - 1) * $MaxLinesPerPage + $k)줄"; $r2[$pc] = '가사'
    }
    Add-Row $sh $r1 $S_TITLE
    Add-Row $sh $r2 $s2
    $idx = @{}; foreach ($it in $pg.Items) { $idx["$($it.Local)-$($it.Cell.Jg)"] = $it.Cell }
    for ($jg = 1; $jg -le $d.RowCount; $jg++) {
      $vals = @($jg) + @(1..($L * 2) | ForEach-Object { '' })
      $sts = @($S_HEAD) + @(1..($L * 2) | ForEach-Object { $S_CELL })
      $maxLines = 1
      for ($k = 1; $k -le $L; $k++) {
        $c = $idx["$k-$jg"]; if (-not $c) { continue }
        $pc = 2 + ($L - $k) * 2
        $vals[$pc - 1] = $c.Pitch; $vals[$pc] = $c.Lyric
        $st = Cell-Style $c; $sts[$pc - 1] = $st; $sts[$pc] = $st
        $maxLines = [math]::Max($maxLines, @($c.Pitch -split "`n").Count)
      }
      Add-Row $sh $vals $sts
      $sh.Heights[$jg + 2] = [math]::Max(30, 16 * $maxLines)
    }
    $sheets.Add($sh)
  }
  Save-Xlsx (Join-Path $ReviewDir "$($d.Name).xlsx") $sheets
}

# ───────────────────── 실행 ─────────────────────
foreach ($dir in @($ReviewDir, $CsvDir)) { if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null } }
Get-ChildItem -LiteralPath $ReviewDir -Filter *.xlsx | Remove-Item
Get-ChildItem -LiteralPath $CsvDir -Filter *.csv | Remove-Item

$csvErrors = 0
foreach ($d in $docs) {
  $text = Build-Csv $d
  $errs = Test-Csv $d $text
  foreach ($e in $errs) { $Problems.Add("[CSV] $($d.Name): $e"); $csvErrors++ }
  $bytes = New-Object System.IO.MemoryStream
  $bytes.Write([byte[]](0xEF, 0xBB, 0xBF), 0, 3)
  $body = $utf8.GetBytes($text); $bytes.Write($body, 0, $body.Length)
  [IO.File]::WriteAllBytes((Join-Path $CsvDir "$($d.Name).csv"), $bytes.ToArray())
  Build-ReviewBook $d
}

# 목록
$list = New-Sheet '곡 목록'; $list.FreezeRows = 1
$w = @(5, 30, 20, 14, 6, 8, 6, 6, 6, 8, 6, 8, 6, 40, 50); for ($i = 0; $i -lt $w.Count; $i++) { $list.Widths[$i + 1] = $w[$i] }
Add-Row $list @('번호', 'CSV 파일', '제목', '부제', '키', '쪽', '박자', '행수', '페이지', '정간 수', '비고', '팔레트 외', '신뢰도', '확인 필요', '원본 이미지') $S_HEAD
$no = 0
foreach ($d in $docs) {
  $no++
  $g = $(switch ($d.Grade) { '상' { $S_CELL } '중' { $S_NOTE } default { $S_BAD } })
  Add-Row $list @($no, "$($d.Name).csv", $d.Title, $d.Sub, $d.Key, $d.PageNo, $d.Beat, $d.RowCount, $d.Pages.Count, $d.CellCount, $d.NoteCount, $d.UnknownCount, $d.Grade, $d.Reasons, (($d.Images | ForEach-Object Source) -join "`n")) `
    @($S_CELL, $S_LEFT, $S_LEFT, $S_LEFT, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $g, $S_LEFT, $S_LEFT)
}
$list.Filter = "A1:O$($list.Rows.Count)"

$imgs = New-Sheet '이미지별'; $imgs.FreezeRows = 1
$w = @(34, 30, 20, 12, 6, 6, 6, 6, 8, 6, 8, 26, 30, 60); for ($i = 0; $i -lt $w.Count; $i++) { $imgs.Widths[$i + 1] = $w[$i] }
Add-Row $imgs @('원본 이미지', 'TSV', '제목', '부제', '박자', '행수', '줄수', '실제 줄', '정간 수', '비고', '팔레트 외', '이어짐', 'CSV 파일', '판독 메모') $S_HEAD
foreach ($d in $docs) {
  foreach ($im in $d.Images) {
    $unk = ($im.Cells | ForEach-Object { $_.Unknown.Count } | Measure-Object -Sum).Sum
    Add-Row $imgs @($im.Source, $im.File, $im.Title, $im.Sub, $im.Beat, $im.RowsMeta, $im.LinesMeta, $im.PresentLines, $im.Cells.Count, @($im.Cells | Where-Object { $_.Note }).Count, [int]$unk, $im.Next, "$($d.Name).csv", $im.Memo) `
      @($S_LEFT, $S_LEFT, $S_LEFT, $S_LEFT, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $S_CELL, $(if ($unk) { $S_BAD } else { $S_CELL }), $S_LEFT, $S_LEFT, $S_LEFT)
  }
}
$imgs.Filter = "A1:N$($imgs.Rows.Count)"

$iss = New-Sheet '확인 필요 글자'; $iss.FreezeRows = 1
$w = @(30, 30, 6, 6, 6, 8, 16, 40); for ($i = 0; $i -lt $w.Count; $i++) { $iss.Widths[$i + 1] = $w[$i] }
Add-Row $iss @('CSV 파일', '원본 이미지', '페이지', '줄', '정간', '글자', '율명 칸', '비고') $S_HEAD
$issueCount = 0
foreach ($d in $docs) {
  for ($p = 0; $p -lt $d.Pages.Count; $p++) {
    foreach ($it in ($d.Pages[$p].Items | Sort-Object { $_.Cell.Jul }, { $_.Cell.Jg })) {
      foreach ($u in $it.Cell.Unknown) {
        $issueCount++
        Add-Row $iss @("$($d.Name).csv", $d.Pages[$p].Image.Source, ($p + 1), $it.Cell.Jul, $it.Cell.Jg, $u, $it.Cell.Pitch, $it.Cell.Note) `
          @($S_LEFT, $S_LEFT, $S_CELL, $S_CELL, $S_CELL, $S_BAD, $S_CELL, $S_LEFT)
      }
    }
  }
}
$iss.Filter = "A1:H$($iss.Rows.Count)"

$unkFreq = @{}
foreach ($d in $docs) { foreach ($im in $d.Images) { foreach ($c in $im.Cells) { foreach ($u in $c.Unknown) { $unkFreq[$u] = 1 + [int]$unkFreq[$u] } } } }
$freq = New-Sheet '글자 빈도'; $freq.FreezeRows = 1; $freq.Widths[1] = 8; $freq.Widths[2] = 10; $freq.Widths[3] = 12; $freq.Widths[4] = 50
Add-Row $freq @('글자', '횟수', '코드', '제안') $S_HEAD
$suggest = @{ '古' = '姑 (고선)으로 보임'; '神' = '㳞 (氵仲, 청중려)으로 보임'; '?' = '원본 확인' }
foreach ($kv in ($unkFreq.GetEnumerator() | Sort-Object Value -Descending)) {
  $code = 'U+' + ([char]::ConvertToUtf32($kv.Key, 0)).ToString('X4')
  Add-Row $freq @($kv.Key, $kv.Value, $code, $(if ($suggest.ContainsKey($kv.Key)) { $suggest[$kv.Key] } else { '' })) @($S_BAD, $S_CELL, $S_CELL, $S_LEFT)
}

Save-Xlsx (Join-Path $ReviewDir '_목록.xlsx') @($list, $imgs, $iss, $freq)

# sheets-index.json (generate_index.py 와 같은 형식 — 저장소에서는 그 스크립트로 다시 만든다)
function JStr([string]$s) {
  if ($null -eq $s) { return 'null' }
  $sb = New-Object System.Text.StringBuilder; [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    if ($ch -eq [char]'"') { [void]$sb.Append('\"') }
    elseif ($ch -eq [char]'\') { [void]$sb.Append('\\') }
    elseif ([int]$ch -lt 0x20) { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) }
    else { [void]$sb.Append($ch) }
  }
  [void]$sb.Append('"'); return $sb.ToString()
}
$Owner = 'lks87454255'; $Repo = 'jeongganview-asset'; $Branch = 'main'
$base = "https://raw.githubusercontent.com/$Owner/$Repo/$Branch"
$sha = [System.Security.Cryptography.SHA256]::Create()
$entries = foreach ($d in $docs) {
  $fn = "$($d.Name).csv"; $full = Join-Path $CsvDir $fn
  $bytes = [IO.File]::ReadAllBytes($full)
  $hash = -join ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') })
  $url = "$base/" + ((@('sheets', $Category, 'csv', $fn) | ForEach-Object { [Uri]::EscapeDataString($_.Normalize([Text.NormalizationForm]::FormC)) }) -join '/')
  "        {`n          `"name`": $(JStr $d.Name),`n          `"title`": $(JStr $d.Title),`n          `"fileName`": $(JStr $fn),`n          `"path`": $(JStr "sheets/$Category/csv/$fn"),`n          `"url`": $(JStr $url),`n          `"sizeBytes`": $($bytes.Length),`n          `"beat`": $(JStr $d.Beat),`n          `"rows`": $($d.RowCount),`n          `"pages`": $($d.Pages.Count),`n          `"sha256`": $(JStr $hash)`n        }"
}
$json = "{`n  `"version`": 1,`n  `"generated`": $(JStr ((Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))),`n  `"owner`": $(JStr $Owner),`n  `"repo`": $(JStr $Repo),`n  `"branch`": $(JStr $Branch),`n  `"categories`": [`n    {`n      `"name`": $(JStr $Category),`n      `"files`": [`n" + ($entries -join ",`n") + "`n      ]`n    }`n  ],`n  `"_total`": { `"categories`": 1, `"files`": $($docs.Count) }`n}`n"
[IO.File]::WriteAllText((Join-Path $RepoDir 'sheets-index.json'), $json, $utf8)

# 리포트
$rep = New-Object System.Collections.Generic.List[string]
$rep.Add("민요채보 생성 리포트  $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
$rep.Add("TSV $($tsvs.Count)개 → 곡 $($docs.Count)개 (페이지 $(($docs | ForEach-Object { $_.Pages.Count } | Measure-Object -Sum).Sum)), 정간 $(($docs | Measure-Object CellCount -Sum).Sum)")
$rep.Add("신뢰도  상 $(@($docs | Where-Object Grade -eq '상').Count) / 중 $(@($docs | Where-Object Grade -eq '중').Count) / 하 $(@($docs | Where-Object Grade -eq '하').Count)")
$rep.Add("팔레트 외 글자 ${issueCount}곳 (종류 $($unkFreq.Count)), CSV 검증 오류 ${csvErrors}건")
$rep.Add('')
$rep.Add('── 곡 ──')
foreach ($d in $docs) { $rep.Add(("{0} | {1} {2}/{3} p{4} | 정간 {5} 비고 {6} 외 {7} | {8} | {9}" -f "$($d.Name).csv", $d.Grade, $d.Beat, $d.RowCount, $d.Pages.Count, $d.CellCount, $d.NoteCount, $d.UnknownCount, (($d.Images | ForEach-Object File) -join ' → '), $d.Reasons)) }
$rep.Add('')
$rep.Add('── 처리 중 문제 ──')
if ($Problems.Count) { $Problems | ForEach-Object { $rep.Add($_) } } else { $rep.Add('없음') }
[IO.File]::WriteAllLines($ReportPath, $rep, $utf8)

"DONE docs=$($docs.Count) tsv=$($tsvs.Count) issues=$issueCount csvErrors=$csvErrors problems=$($Problems.Count)"
