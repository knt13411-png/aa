# 가상 샘플 엑셀 3건 생성기 (실제 데이터 아님)
# 사용: powershell -ExecutionPolicy Bypass -File tools\make_samples.ps1
param([string]$OutDir = (Join-Path $PSScriptRoot '..\samples'))

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null

# ---------- xlsx 작성 ----------
function Esc([string]$s) { return [System.Security.SecurityElement]::Escape($s) }

function ColName([int]$i) {
  $s = ''
  $i++
  while ($i -gt 0) { $m = ($i - 1) % 26; $s = [char](65 + $m) + $s; $i = [math]::Floor(($i - 1) / 26) }
  return $s
}

# 셀 값: [double] -> 숫자, [datetime] -> 날짜(스타일 1), $null -> 빈칸, 그 외 -> 문자열
function SheetXml($headers, $rows) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
  $cols = @(); for ($c = 0; $c -lt $headers.Count; $c++) { $cols += (ColName $c) }
  [void]$sb.Append('<row r="1">')
  for ($c = 0; $c -lt $headers.Count; $c++) { [void]$sb.Append('<c r="' + $cols[$c] + '1" t="inlineStr"><is><t>' + (Esc $headers[$c]) + '</t></is></c>') }
  [void]$sb.Append('</row>')
  $r = 1
  foreach ($row in $rows) {
    $r++
    [void]$sb.Append('<row r="' + $r + '">')
    for ($c = 0; $c -lt $row.Count; $c++) {
      $v = $row[$c]
      if ($null -eq $v) { continue }
      $ref = $cols[$c] + $r
      if ($v -is [datetime]) {
        $serial = ($v - [datetime]'1899-12-30').TotalDays
        [void]$sb.Append('<c r="' + $ref + '" s="1"><v>' + $serial.ToString('R', [Globalization.CultureInfo]::InvariantCulture) + '</v></c>')
      } elseif ($v -is [double] -or $v -is [int]) {
        [void]$sb.Append('<c r="' + $ref + '"><v>' + ([double]$v).ToString('R', [Globalization.CultureInfo]::InvariantCulture) + '</v></c>')
      } else {
        [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' + (Esc ([string]$v)) + '</t></is></c>')
      }
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData></worksheet>')
  return $sb.ToString()
}

function WriteXlsx($path, $sheets) {
  if (Test-Path $path) { Remove-Item $path -Force }
  $utf8 = New-Object System.Text.UTF8Encoding $false
  $zip = [System.IO.Compression.ZipFile]::Open($path, 'Create')
  function Put($name, $text) {
    $e = $zip.CreateEntry($name)
    $w = New-Object System.IO.StreamWriter($e.Open(), $utf8)
    $w.Write($text); $w.Close()
  }
  $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
  $wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
  $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  for ($i = 0; $i -lt $sheets.Count; $i++) {
    $n = $i + 1
    $ct += '<Override PartName="/xl/worksheets/sheet' + $n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    $wb += '<sheet name="' + (Esc $sheets[$i].Name) + '" sheetId="' + $n + '" r:id="rId' + $n + '"/>'
    $rels += '<Relationship Id="rId' + $n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $n + '.xml"/>'
    Put ('xl/worksheets/sheet' + $n + '.xml') (SheetXml $sheets[$i].Headers $sheets[$i].Rows)
  }
  $sn = $sheets.Count + 1
  $rels += '<Relationship Id="rId' + $sn + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
  $ct += '</Types>'
  $wb += '</sheets></workbook>'
  Put '[Content_Types].xml' $ct
  Put '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
  Put 'xl/workbook.xml' $wb
  Put 'xl/_rels/workbook.xml.rels' $rels
  Put 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts><fonts count="1"><font><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'
  $zip.Dispose()
}

# ---------- 가상 데이터 생성 ----------
function MakeTraining($p) {
  $rnd = New-Object System.Random $p.Seed
  $Uni = { param($a, $b) $a + ($b - $a) * $rnd.NextDouble() }
  $ExpR = { param($mean) - $mean * [math]::Log(1 - $rnd.NextDouble()) }
  $PickR = { param($arr) $arr[$rnd.Next($arr.Count)] }
  $t0 = $p.Start
  $durMin = 360

  # 부대 편성: 훈련부대 1여단(3대대x3중대x3소대), 대항군 대항군여단(2대대x3중대x3소대)
  $units = New-Object System.Collections.ArrayList
  foreach ($side in @('훈련부대', '대항군')) {
    $bde = if ($side -eq '훈련부대') { '1여단' } else { '대항군여단' }
    $nBn = if ($side -eq '훈련부대') { 3 } else { 2 }
    $nEnt = if ($side -eq '훈련부대') { 18 } else { 14 }
    for ($b = 1; $b -le $nBn; $b++) { for ($c = 1; $c -le 3; $c++) { for ($l = 1; $l -le 3; $l++) {
      $code = if ($side -eq '훈련부대') { 'T' } else { 'O' }
      $ents = @(); for ($e = 1; $e -le $nEnt; $e++) { $ents += ('{0}{1}{2}{3}-{4:D2}' -f $code, $b, $c, $l, $e) }
      # 대대별 숙련도 차이(드릴다운 확인용)
      $skill = if ($side -eq '훈련부대') { @{1 = 1.0; 2 = 0.8; 3 = 1.12}[$b] } else { 1.0 }
      [void]$units.Add(@{ Side = $side; Bde = $bde; Bn = "${b}대대"; Co = "${c}중대"; Pl = "${l}소대"; Ents = $ents; Skill = $skill })
    } } }
  }
  $alive = @{}
  foreach ($u in $units) { foreach ($e in $u.Ents) { $alive[$e] = 0 } }   # 0 없음,1 경상,2 중상,3 사망
  $entsBySide = @{ '훈련부대' = @(); '대항군' = @() }
  $unitOf = @{}
  foreach ($u in $units) { $entsBySide[$u.Side] += $u.Ents; foreach ($e in $u.Ents) { $unitOf[$e] = $u } }

  $fire = New-Object System.Collections.ArrayList
  $dmg = New-Object System.Collections.ArrayList
  $weapons = @(
    @{ N = '소총'; K = '직사'; W = 58; H = 1.0; R = @(50, 400) },
    @{ N = '기관총'; K = '직사'; W = 24; H = 0.85; R = @(100, 800) },
    @{ N = '대전차'; K = '직사'; W = 7; H = 1.25; R = @(200, 1500) },
    @{ N = '박격포'; K = '간접'; W = 7; H = 0.9; R = @(500, 3000) },
    @{ N = '포병'; K = '간접'; W = 4; H = 0.8; R = @(2000, 9000) })
  $wSum = 0; foreach ($w in $weapons) { $wSum += $w.W }

  # 교전 이벤트 (시간순 처리 위해 먼저 모두 만들고 정렬)
  $shots = New-Object System.Collections.ArrayList
  foreach ($u in $units) {
    $sp = if ($u.Side -eq '훈련부대') { $p.T } else { $p.O }
    $nEng = $rnd.Next(6, 10)
    for ($k = 0; $k -lt $nEng; $k++) {
      # 초기 집중도: 일부 교전은 첫 40분 안에 시작
      $start = if ($rnd.NextDouble() -lt $sp.Early) { & $Uni 0 40 } else { & $Uni 0 ($durMin - 20) }
      $n = $rnd.Next(10, 26)
      $t = $start
      for ($s = 0; $s -lt $n; $s++) {
        $t += (& $Uni 0.05 3.5)
        if ($t -ge $durMin) { break }
        [void]$shots.Add(@{ T = $t; U = $u; SP = $sp })
      }
    }
  }
  $sorted = $shots | Sort-Object { $_.T }
  foreach ($sh in $sorted) {
    $u = $sh.U; $sp = $sh.SP
    $shooter = & $PickR $u.Ents
    if ($alive[$shooter] -eq 3) { continue }
    $x = $rnd.NextDouble() * $wSum; $w = $weapons[0]
    foreach ($ww in $weapons) { if ($x -lt $ww.W) { $w = $ww; break }; $x -= $ww.W }
    $tSide = if ($u.Side -eq '훈련부대') { '대항군' } else { '훈련부대' }
    $target = $null
    for ($try = 0; $try -lt 6; $try++) { $cand = & $PickR $entsBySide[$tSide]; if ($alive[$cand] -lt 3) { $target = $cand; break } }
    if (-not $target) { continue }
    $hit = $rnd.NextDouble() -lt ($sp.Hit * $w.H * $u.Skill)
    $res = '없음'
    if ($hit -and $rnd.NextDouble() -lt ($sp.Dmg * $u.Skill)) {
      $q = $rnd.NextDouble()
      $res = if ($q -lt 0.35) { '사망' } elseif ($q -lt 0.65) { '중상' } else { '경상' }
      $sev = @{ '경상' = 1; '중상' = 2; '사망' = 3 }[$res]
      $tu = $unitOf[$target]
      $tsp = if ($tu.Side -eq '훈련부대') { $p.T } else { $p.O }
      if ($sev -gt $alive[$target]) { $alive[$target] = $sev }
      $mv = if ($rnd.NextDouble() -lt $tsp.MoveDmg) { 'Y' } else { 'N' }
      $cv = if ($rnd.NextDouble() -lt $tsp.Uncovered) { 'N' } else { 'Y' }
      [void]$dmg.Add(@($t0.AddMinutes($sh.T), $tu.Side, $tu.Bde, $tu.Bn, $tu.Co, $tu.Pl, $target, $res, $w.K, $cv, $mv))
    }
    $dist = [math]::Round((& $Uni $w.R[0] $w.R[1]))
    [void]$fire.Add(@($t0.AddMinutes($sh.T), $u.Side, $u.Bde, $u.Bn, $u.Co, $u.Pl, $shooter, $w.N, $w.K, $target, $tSide, $(if ($hit) { '명중' } else { '빗나감' }), $res, [double]$dist))
  }
  # 지뢰/기타 피해 소량
  for ($k = 0; $k -lt 12; $k++) {
    $u = & $PickR $units; $e = & $PickR $u.Ents
    if ($alive[$e] -eq 3) { continue }
    $sp = if ($u.Side -eq '훈련부대') { $p.T } else { $p.O }
    $res = & $PickR @('사망', '중상', '경상'); $sev = @{ '경상' = 1; '중상' = 2; '사망' = 3 }[$res]
    if ($sev -gt $alive[$e]) { $alive[$e] = $sev }
    [void]$dmg.Add(@($t0.AddMinutes((& $Uni 10 350)), $u.Side, $u.Bde, $u.Bn, $u.Co, $u.Pl, $e, $res, (& $PickR @('지뢰', '기타')), 'N', $(if ($rnd.NextDouble() -lt $sp.MoveDmg) { 'Y' } else { 'N' })))
  }

  # 탐지기록
  $det = New-Object System.Collections.ArrayList
  foreach ($u in $units) {
    $sp = if ($u.Side -eq '훈련부대') { $p.T } else { $p.O }
    $n = $rnd.Next(12, 20)
    for ($k = 0; $k -lt $n; $k++) {
      $t = & $Uni 0 $durMin
      $td = $t0.AddMinutes($t)
      $idt = $td.AddSeconds([math]::Round((& $ExpR ($sp.IdSec / $u.Skill)) + 5))
      $tgtSide = if ($u.Side -eq '훈련부대') { '대항군' } else { '훈련부대' }
      $opp = if ($rnd.NextDouble() -lt 0.12) { $null } else { $td.AddSeconds([math]::Round((& $Uni (-240) 240) + $sp.PreDet)) }
      [void]$det.Add(@($td, $u.Side, $u.Bde, $u.Bn, $u.Co, $u.Pl, (& $PickR @('육안', '육안', '열상', '레이더', 'UAV')), (& $PickR $entsBySide[$tgtSide]), $td, $idt, $opp))
    }
  }

  # 위치기록 (10분 간격)
  $pos = New-Object System.Collections.ArrayList
  foreach ($u in $units) {
    $sp = if ($u.Side -eq '훈련부대') { $p.T } else { $p.O }
    foreach ($e in $u.Ents) {
      $x = & $Uni 1000 9000; $y = & $Uni 1000 9000
      for ($m = 0; $m -le $durMin; $m += 10) {
        $moving = $rnd.NextDouble() -lt 0.4
        $v = 0.0
        if ($moving) { $v = [math]::Round([math]::Max(1, (& $Uni (-4) 4) + $sp.Speed * $u.Skill), 1); $x += (& $Uni (-1) 1) * $v * 100; $y += (& $Uni (-1) 1) * $v * 100 }
        [void]$pos.Add(@($t0.AddMinutes($m), $u.Side, $u.Bde, $u.Bn, $u.Co, $u.Pl, $e, [double][math]::Round($x), [double][math]::Round($y), [double]$v, $(if ($moving) { 'Y' } else { 'N' })))
      }
    }
  }

  # 지휘통신
  $c2 = New-Object System.Collections.ArrayList
  $mid = 0
  $pairs = New-Object System.Collections.ArrayList
  foreach ($u in $units) {
    $bn = "$($u.Bde)/$($u.Bn)"; $co = "$bn/$($u.Co)"; $pl = "$co/$($u.Pl)"
    [void]$pairs.Add(@($u, $co, $pl))
    if ($u.Pl -eq '1소대') { [void]$pairs.Add(@($u, $bn, $co)) }
    if ($u.Pl -eq '1소대' -and $u.Co -eq '1중대') { [void]$pairs.Add(@($u, $u.Bde, $bn)) }
  }
  foreach ($pr in $pairs) {
    $u = $pr[0]; $sp = if ($u.Side -eq '훈련부대') { $p.T } else { $p.O }
    $orders = @()
    for ($k = 0; $k -lt $rnd.Next(4, 8); $k++) {
      $mid++; $id = 'M{0:D5}' -f $mid; $orders += $id
      $t = $t0.AddMinutes((& $Uni 0 $durMin))
      [void]$c2.Add(@($t, $u.Side, $pr[1], $pr[2], '명령', $id, $t.AddSeconds([math]::Round((& $ExpR ($sp.OrderSec / $u.Skill)) + 10))))
    }
    for ($k = 0; $k -lt $rnd.Next(5, 10); $k++) {
      $mid++
      $t = $t0.AddMinutes((& $Uni 0 $durMin))
      [void]$c2.Add(@($t, $u.Side, $pr[2], $pr[1], '보고', (& $PickR $orders), $t.AddSeconds([math]::Round((& $ExpR ($sp.ReportSec / $u.Skill)) + 10))))
    }
  }

  # 지속지원
  $sus = New-Object System.Collections.ArrayList
  foreach ($u in $units) {
    $sp = if ($u.Side -eq '훈련부대') { $p.T } else { $p.O }
    for ($k = 0; $k -lt $rnd.Next(2, 5); $k++) {
      $t = $t0.AddMinutes((& $Uni 20 ($durMin - 30)))
      $use = [math]::Round((& $Uni 300 900))
      $qty = [math]::Round($use * $sp.Supply * (& $Uni 0.8 1.2))
      [void]$sus.Add(@($t, $t.AddMinutes([math]::Round((& $ExpR $sp.ResupMin) + 5)), $u.Side, $u.Bde, $u.Bn, $u.Co, $u.Pl, '탄약보급', [double]$qty, [double]$use))
    }
    foreach ($e in $u.Ents) {
      if ($alive[$e] -eq 1 -or $alive[$e] -eq 2) {
        $t = $t0.AddMinutes((& $Uni 30 ($durMin - 10)))
        [void]$sus.Add(@($t, $t.AddMinutes([math]::Round((& $ExpR $sp.EvacMin) + 8)), $u.Side, $u.Bde, $u.Bn, $u.Co, $u.Pl, '후송', [double]1, $null))
      }
    }
    if ($rnd.NextDouble() -lt 0.3) {
      $t = $t0.AddMinutes((& $Uni 30 ($durMin - 60)))
      [void]$sus.Add(@($t, $t.AddMinutes([math]::Round((& $Uni 30 120))), $u.Side, $u.Bde, $u.Bn, $u.Co, $u.Pl, '정비', [double]1, $null))
    }
  }

  $hdrUnit = @('진영', '여단', '대대', '중대', '소대')
  return @(
    @{ Name = '교전기록'; Headers = @('시각') + $hdrUnit + @('사수ID', '화기종류', '화기구분', '표적ID', '표적진영', '결과', '피해판정', '거리(m)'); Rows = $fire },
    @{ Name = '탐지기록'; Headers = @('시각') + $hdrUnit + @('탐지수단', '표적ID', '탐지시각', '식별시각', '상대탐지시각'); Rows = $det },
    @{ Name = '위치기록'; Headers = @('시각') + $hdrUnit + @('개체ID', 'X좌표', 'Y좌표', '속도(km/h)', '이동여부'); Rows = $pos },
    @{ Name = '피해기록'; Headers = @('시각') + $hdrUnit + @('개체ID', '피해판정', '피해원인', '엄폐여부', '이동여부'); Rows = $dmg },
    @{ Name = '지휘통신'; Headers = @('시각', '진영', '발신제대', '수신제대', '유형', '관련메시지ID', '응답시각'); Rows = $c2 },
    @{ Name = '지속지원'; Headers = @('요청시각', '완료시각', '진영', '여단', '대대', '중대', '소대', '유형', '수량', '탄약소모량'); Rows = $sus }
  )
}

# 대항군은 세 훈련 모두 같은 수준
$opfor = @{ Hit = 0.33; Dmg = 0.16; Early = 0.25; MoveDmg = 0.4; Uncovered = 0.45; Speed = 11; IdSec = 70; PreDet = 0; OrderSec = 90; ReportSec = 110; ResupMin = 35; EvacMin = 35; Supply = 0.9 }

$trainings = @(
  # 1차: 기준
  @{ File = '샘플1_2026-06_백두1호훈련.xlsx'; Seed = 101; Start = [datetime]'2026-06-16 06:00:00'; O = $opfor
     T = @{ Hit = 0.27; Dmg = 0.15; Early = 0.35; MoveDmg = 0.35; Uncovered = 0.55; Speed = 12; IdSec = 95; PreDet = 20; OrderSec = 130; ReportSec = 160; ResupMin = 40; EvacMin = 35; Supply = 0.85 } },
  # 2차: 화력 향상, 기동 저하
  @{ File = '샘플2_2026-08_백두2호훈련.xlsx'; Seed = 202; Start = [datetime]'2026-08-19 06:00:00'; O = $opfor
     T = @{ Hit = 0.36; Dmg = 0.19; Early = 0.35; MoveDmg = 0.6; Uncovered = 0.55; Speed = 8; IdSec = 90; PreDet = 25; OrderSec = 125; ReportSec = 150; ResupMin = 38; EvacMin = 34; Supply = 0.88 } },
  # 3차: 지휘통제·정보·방호 향상, 작전지속지원 저하
  @{ File = '샘플3_2026-10_백두3호훈련.xlsx'; Seed = 303; Start = [datetime]'2026-10-14 06:00:00'; O = $opfor
     T = @{ Hit = 0.35; Dmg = 0.18; Early = 0.15; MoveDmg = 0.42; Uncovered = 0.35; Speed = 11; IdSec = 55; PreDet = 110; OrderSec = 70; ReportSec = 85; ResupMin = 70; EvacMin = 60; Supply = 0.6 } }
)

foreach ($tr in $trainings) {
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $sheets = MakeTraining $tr
  $path = Join-Path $OutDir $tr.File
  WriteXlsx $path $sheets
  $counts = ($sheets | ForEach-Object { "$($_.Name)=$($_.Rows.Count)" }) -join ', '
  Write-Host "$($tr.File): $counts ($([math]::Round($sw.Elapsed.TotalSeconds,1))s)"
}
