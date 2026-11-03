# src/ 와 lib/ 를 합쳐 단일 HTML 파일(dist/사후검토_분석기.html)을 만든다.
# 사용: powershell -ExecutionPolicy Bypass -File tools\build.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object System.Text.UTF8Encoding $false
function Read-Utf8($p) { [IO.File]::ReadAllText((Join-Path $root $p), [Text.Encoding]::UTF8) }
# 인라인 스크립트 안의 </script 는 HTML 파서가 끊으므로 이스케이프, 소스맵 참조(외부 요청)는 제거
function Clean($s) { ($s -replace '//[#@] sourceMappingURL=\S*', '') -replace '</script', '<\/script' }

$libs = (Clean (Read-Utf8 'lib\xlsx.full.min.js')) + "`n;`n" + (Clean (Read-Utf8 'lib\chart.umd.min.js'))
$order = 'config', 'util', 'engine', 'insight', 'charts', 'pptx', 'deck', 'sample', 'ui'
$app = ($order | ForEach-Object { "/* ===== $_.js ===== */`n" + (Clean (Read-Utf8 "src\js\$_.js")) }) -join "`n"
$app = "(function () {`n$app`n})();"

$html = Read-Utf8 'src\app.html'
$html = $html.Replace('/*__CSS__*/', (Read-Utf8 'src\app.css'))
$html = $html.Replace('//__LIBS__', $libs)
$html = $html.Replace('//__APP__', $app)

New-Item -ItemType Directory -Force (Join-Path $root 'dist') | Out-Null
$out = Join-Path $root 'dist\사후검토_분석기.html'
[IO.File]::WriteAllText($out, $html, $utf8)
"빌드 완료: $out ($([math]::Round((Get-Item $out).Length / 1KB)) KB)"
# 외부 주소 점검 (오프라인 확인용)
$ext = [regex]::Matches($app + (Read-Utf8 'src\app.html'), 'https?://[^\s"''`)]+') | ForEach-Object Value | Where-Object { $_ -notmatch '^http://(schemas\.openxmlformats\.org|purl\.org|www\.w3\.org)/' } | Sort-Object -Unique
if ($ext) { "앱 코드 내 외부 주소: $($ext -join ', ')" } else { '앱 코드 내 외부 주소 없음' }
