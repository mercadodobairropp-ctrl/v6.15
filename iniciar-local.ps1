Set-Location -LiteralPath $PSScriptRoot

$pythonExe = "C:\Users\avail\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$url = "http://127.0.0.1:5500/"

Start-Process $url

if (Test-Path -LiteralPath $pythonExe) {
  & $pythonExe -m http.server 5500 --bind 127.0.0.1
} else {
  py -m http.server 5500 --bind 127.0.0.1
}
