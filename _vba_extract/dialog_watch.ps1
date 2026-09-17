param([string]$Log = 'C:\citrixlabph\globalsmile\_vba_extract\dialog_watch.log')

$src2 = @'
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;
public class WinUtil {
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc f, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);

    public static List<string> TopWindows() {
        var res = new List<string>();
        EnumWindows((h, l) => {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (IsWindowVisible(h)) {
                var t = new StringBuilder(512); GetWindowText(h, t, 512);
                var c = new StringBuilder(256); GetClassName(h, c, 256);
                res.Add(h.ToInt64() + "|" + pid + "|" + c + "|" + t);
            }
            return true;
        }, IntPtr.Zero);
        return res;
    }
    public static List<string> ChildTexts(IntPtr hwnd) {
        var res = new List<string>();
        EnumChildWindows(hwnd, (h, l) => {
            var t = new StringBuilder(1024); GetWindowText(h, t, 1024);
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            res.Add(c + "|" + t);
            return true;
        }, IntPtr.Zero);
        return res;
    }
}
'@
Add-Type -TypeDefinition $src2 -Language CSharp
Add-Type -AssemblyName System.Windows.Forms

function W($m) { "$(Get-Date -Format 'HH:mm:ss.fff')  $m" | Out-File -FilePath $Log -Append -Encoding utf8 }

$deadline = (Get-Date).AddSeconds(240)
$lastF5 = Get-Date
W "watcher started"
while ((Get-Date) -lt $deadline) {
    $excelPids = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
    $wins = [WinUtil]::TopWindows()
    foreach ($w in $wins) {
        $parts = $w -split '\|', 4
        $hwnd = [IntPtr][long]$parts[0]
        $proc = [int]$parts[1]
        $cls  = $parts[2]
        $tit  = $parts[3]
        if ($excelPids -notcontains $proc) { continue }
        if ($cls -eq '#32770') {
            $kids = [WinUtil]::ChildTexts($hwnd) | Where-Object { $_ -match '^(Static|Button)\|' -and $_ -notmatch '^\w+\|\s*$' }
            W "DIALOG hwnd=$($parts[0]) title='$tit' kids=$($kids -join ' ; ')"
            $btnEnd   = [WinUtil]::FindWindowEx($hwnd, [IntPtr]::Zero, 'Button', 'End')
            $btnOk    = [WinUtil]::FindWindowEx($hwnd, [IntPtr]::Zero, 'Button', 'OK')
            $btnCont  = [WinUtil]::FindWindowEx($hwnd, [IntPtr]::Zero, 'Button', 'Continue')
            $target = $btnEnd
            if ($target -eq [IntPtr]::Zero) { $target = $btnOk }
            if ($target -eq [IntPtr]::Zero) { $target = $btnCont }
            if ($target -ne [IntPtr]::Zero) {
                W "clicking button"
                [WinUtil]::SendMessage($target, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            } else {
                W "no known button; sending WM_CLOSE"
                [WinUtil]::PostMessage($hwnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            }
        }
        elseif ($tit -like '*[[]break[]]*' -and ((Get-Date) - $lastF5).TotalSeconds -ge 3) {
            W "break-mode VBE detected: '$tit' -> sending F5"
            [WinUtil]::SetForegroundWindow($hwnd) | Out-Null
            Start-Sleep -Milliseconds 300
            [System.Windows.Forms.SendKeys]::SendWait('{F5}')
            $lastF5 = Get-Date
        }
    }
    Start-Sleep -Milliseconds 700
}
W "watcher done"
