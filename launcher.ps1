<#
    RangeSheet — ตัวเปิดโปรแกรม (เรียกจาก START_RANGESHEET.bat)

    เป้าหมาย: ให้คนที่ไม่มี Python / ไม่มี Git / ไม่อยากยุ่งกับ Terminal
    ดับเบิลคลิกทีเดียวแล้วใช้โปรแกรมได้เลย

    ครั้งแรกที่รัน สคริปต์จะ:
      1. หา Python ในเครื่อง — ถ้าไม่มี จะโหลด Python แบบพกพา (embeddable)
         มาไว้ในโฟลเดอร์ runtime\ เท่านั้น ไม่ได้ติดตั้งลงเครื่อง
         ไม่ต้องใช้สิทธิ์ admin และลบโฟลเดอร์ทิ้ง = หายเกลี้ยง
      2. ติดตั้งไลบรารีตาม requirements.txt ลงในพื้นที่แยกของโปรแกรม
      3. เปิดโปรแกรมขึ้นเบราว์เซอร์

    ครั้งแรกใช้เวลาราว 5-15 นาที (ต้องต่อเน็ต โหลดประมาณ 300-400 MB)
    ครั้งต่อไปเปิดภายในไม่กี่วินาที
#>

# native command เขียน stderr เป็นเรื่องปกติ อย่าให้ throw — เช็ก exit code เอง
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$PY_VERSION = '3.11.9'
$PY_URL     = "https://www.python.org/ftp/python/$PY_VERSION/python-$PY_VERSION-embed-amd64.zip"
$GETPIP_URL = 'https://bootstrap.pypa.io/get-pip.py'

function Say  { param($m) Write-Host "  $m" }
function Ok   { param($m) Write-Host "  $m" -ForegroundColor Green }
function Warn { param($m) Write-Host "  $m" -ForegroundColor Yellow }
function Die {
    param($title, $detail)
    Write-Host ""
    Write-Host "  [!] $title" -ForegroundColor Red
    Write-Host ""
    foreach ($line in $detail) { Write-Host "      $line" -ForegroundColor Yellow }
    Write-Host ""
    exit 1
}

Write-Host ""
Write-Host "  ===========================================================" -ForegroundColor Cyan
Write-Host "    RangeSheet" -ForegroundColor Cyan
Write-Host "  ===========================================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path (Join-Path $Root 'APP.py'))) {
    Die "ไม่พบไฟล์ APP.py" @(
        "ต้องวางไฟล์นี้ไว้ในโฟลเดอร์เดียวกับ APP.py",
        "ถ้าเพิ่งแตกไฟล์ ZIP มา ให้เข้าไปข้างในโฟลเดอร์ที่แตกออกมาก่อน"
    )
}

# ── 1. หา Python ที่ใช้ได้ ────────────────────────────────────────────────────
Say "[1/3] กำลังเตรียม Python..."

$VenvPy    = Join-Path $Root '.venv\Scripts\python.exe'
$RuntimePy = Join-Path $Root 'runtime\python.exe'
$Py        = $null

if     (Test-Path $VenvPy)    { $Py = $VenvPy }
elseif (Test-Path $RuntimePy) { $Py = $RuntimePy }

function Find-SystemPython {
    # คืน path ของ python.exe ในเครื่อง ถ้าเป็นเวอร์ชัน 3.10 ขึ้นไป
    #
    # หมายเหตุ: โค้ด probe ต้องไม่มีเครื่องหมายคำพูดเลย เพราะ PowerShell กลืน
    # "" ตอนส่งต่อให้โปรแกรมภายนอก ทำให้ Python เจอ string ที่ปิดไม่ครบ
    # (else "")  ->  else ")  ->  SyntaxError
    $probe = 'import sys; print(sys.executable) if sys.version_info[:2] >= (3, 10) else sys.exit(1)'
    foreach ($name in @('py', 'python')) {
        try {
            if ($name -eq 'py') { $out = & py -3 -c $probe 2>$null }
            else                { $out = & python -c $probe 2>$null }
            if ($LASTEXITCODE -eq 0 -and $out) {
                $exe = ($out | Select-Object -First 1).Trim()
                if ($exe -and (Test-Path $exe)) { return $exe }
            }
        } catch { }
    }
    return $null
}

if (-not $Py) {
    $sysPy = Find-SystemPython
    if ($sysPy) {
        Say "      เจอ Python ในเครื่องแล้ว - กำลังสร้างพื้นที่แยกให้โปรแกรม..."
        & $sysPy -m venv (Join-Path $Root '.venv')
        if ($LASTEXITCODE -eq 0 -and (Test-Path $VenvPy)) {
            $Py = $VenvPy
        } else {
            Warn "      สร้างพื้นที่แยกไม่สำเร็จ - จะใช้ Python แบบพกพาแทน"
        }
    }
}

if (-not $Py) {
    Say "      ไม่พบ Python ในเครื่อง - กำลังดาวน์โหลด Python แบบพกพา"
    Say "      (ไม่ได้ติดตั้งลงเครื่อง เก็บไว้ในโฟลเดอร์ runtime\ เท่านั้น)"

    $runtimeDir = Join-Path $Root 'runtime'
    $zipPath    = Join-Path $env:TEMP 'rangesheet_py_embed.zip'
    $getPipPath = Join-Path $env:TEMP 'rangesheet_get_pip.py'

    try {
        Invoke-WebRequest -Uri $PY_URL -OutFile $zipPath -UseBasicParsing
    } catch {
        Die "ดาวน์โหลด Python ไม่สำเร็จ" @(
            "สาเหตุที่เป็นไปได้: ไม่ได้ต่อเน็ต หรือเน็ตบริษัทบล็อก python.org",
            "ทางแก้: ลองต่อเน็ตมือถือ หรือขอ IT เปิดให้เข้า python.org",
            "",
            "รายละเอียด: $($_.Exception.Message)"
        )
    }

    try {
        Expand-Archive -Path $zipPath -DestinationPath $runtimeDir -Force
    } catch {
        Die "แตกไฟล์ Python ไม่สำเร็จ" @("รายละเอียด: $($_.Exception.Message)")
    }

    # Python แบบ embeddable ปิด site-packages ไว้ ต้องเปิดก่อน pip ถึงจะทำงาน
    $pthFile = Join-Path $runtimeDir 'python311._pth'
    Set-Content -Path $pthFile -Encoding ascii -Value @(
        'python311.zip', '.', 'Lib\site-packages', '', 'import site'
    )

    try {
        Invoke-WebRequest -Uri $GETPIP_URL -OutFile $getPipPath -UseBasicParsing
    } catch {
        Die "ดาวน์โหลดตัวติดตั้งไลบรารี (pip) ไม่สำเร็จ" @(
            "รายละเอียด: $($_.Exception.Message)"
        )
    }

    & $RuntimePy $getPipPath --no-warn-script-location | Out-Null
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $RuntimePy)) {
        Die "ติดตั้ง pip ไม่สำเร็จ" @("ลองรัน START_RANGESHEET.bat ใหม่อีกครั้ง")
    }
    $Py = $RuntimePy
}

Ok "[1/3] Python พร้อมแล้ว"
Write-Host ""

# ── 2. ติดตั้งไลบรารี (ข้ามถ้าครบแล้ว) ───────────────────────────────────────
& $Py -c "import streamlit, pandas" 2>$null | Out-Null
$depsOk = ($LASTEXITCODE -eq 0)

if (-not $depsOk) {
    Say "[2/3] กำลังติดตั้งไลบรารีที่โปรแกรมต้องใช้..."
    Warn "      ครั้งแรกใช้เวลา 5-15 นาที (โหลดประมาณ 300-400 MB) รอสักครู่นะครับ"
    Write-Host ""

    & $Py -m pip install --upgrade pip --no-warn-script-location 2>$null | Out-Null
    & $Py -m pip install -r (Join-Path $Root 'requirements.txt') --no-warn-script-location

    if ($LASTEXITCODE -ne 0) {
        Die "ติดตั้งไลบรารีไม่สำเร็จ" @(
            "ลองดับเบิลคลิก START_RANGESHEET.bat ใหม่อีกครั้ง",
            "(บางทีเน็ตหลุดกลางทาง รันซ้ำได้ ไม่ต้องลบอะไร)",
            "",
            "ถ้ายังไม่ได้ ให้ถ่ายรูปหน้าจอนี้ส่งให้ทีมพัฒนา"
        )
    }
}

Ok "[2/3] ไลบรารีพร้อมแล้ว"
Write-Host ""

# ── 3. เปิดโปรแกรม ───────────────────────────────────────────────────────────
Say "[3/3] กำลังเปิดโปรแกรม..."
Write-Host ""
Write-Host "  ===========================================================" -ForegroundColor Cyan
Write-Host "    เดี๋ยวเบราว์เซอร์จะเปิดขึ้นมาเอง" -ForegroundColor Cyan
Write-Host "    ถ้าไม่ขึ้น ให้เปิด Chrome/Edge แล้วพิมพ์:  http://localhost:8501" -ForegroundColor Cyan
Write-Host ""
Write-Host "    เลิกใช้งานแล้ว: ปิดหน้าต่างนี้ทิ้งได้เลย" -ForegroundColor Cyan
Write-Host "  ===========================================================" -ForegroundColor Cyan
Write-Host ""

& $Py -m streamlit run (Join-Path $Root 'APP.py')

Write-Host ""
Say "โปรแกรมปิดแล้ว - ปิดหน้าต่างนี้ได้เลย"
Read-Host "  กด Enter เพื่อปิด" | Out-Null
