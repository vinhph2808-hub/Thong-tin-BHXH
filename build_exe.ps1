$ErrorActionPreference = 'Stop'
$python = 'python'
if (Test-Path -LiteralPath '.venv\Scripts\python.exe') {
    $python = '.venv\Scripts\python.exe'
}
& $python -m pip install -r requirements.txt
& $python -m PyInstaller --noconfirm --clean C12_BHXH.spec
Write-Host 'Hoàn tất: dist\Xuat_Thong_Bao_BHXH_C12.exe'
