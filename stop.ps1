# 停止正在运行的 VbeComplete（只杀命令行里含 VbeComplete 的 python 进程，
# 不会影响其他 Python 程序）
$names = "Name = 'pythonw.exe' OR Name = 'python.exe' OR Name = 'py.exe' OR Name = 'pyw.exe'"
$procs = Get-CimInstance Win32_Process -Filter $names -ErrorAction SilentlyContinue
$n = 0
foreach ($p in $procs) {
    if ($p.CommandLine -and $p.CommandLine -like "*VbeComplete*") {
        try {
            Stop-Process -Id $p.ProcessId -Force
            Write-Host ("killed PID " + $p.ProcessId)
            $n++
        } catch {
            Write-Host ("failed to kill PID " + $p.ProcessId + ": " + $_.Exception.Message)
        }
    }
}
if ($n -eq 0) {
    Write-Host "No VbeComplete process is running."
} else {
    Write-Host ("Stopped " + $n + " process(es).")
}
