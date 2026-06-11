param(
    [int]$Seconds = 30,
    [switch]$Log
)

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class InputTools2
{
    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT
    {
        public UInt32 type;
        public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT
    {
        public Int32 dx;
        public Int32 dy;
        public UInt32 mouseData;
        public UInt32 dwFlags;
        public UInt32 time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public Int32 X;
        public Int32 Y;
    }

    [DllImport("user32.dll", SetLastError=true)]
    public static extern UInt32 SendInput(UInt32 nInputs, INPUT[] pInputs, Int32 cbSize);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetCursorPos(Int32 X, Int32 Y);

    public const UInt32 INPUT_MOUSE = 0;
    public const UInt32 MOUSEEVENTF_MOVE = 0x0001;
}
"@

function Get-MousePos {
    $p = New-Object InputTools2+POINT
    [InputTools2]::GetCursorPos([ref]$p) | Out-Null
    return $p
}

function Send-MouseMove([int]$dx, [int]$dy) {
    $input = New-Object InputTools2+INPUT
    $input.type = [uint32][InputTools2]::INPUT_MOUSE
    $input.mi.dx = $dx
    $input.mi.dy = $dy
    $input.mi.mouseData = [uint32]0
    $input.mi.dwFlags = [uint32][InputTools2]::MOUSEEVENTF_MOVE
    $input.mi.time = [uint32]0
    $input.mi.dwExtraInfo = [IntPtr]::Zero

    $inputs = [InputTools2+INPUT[]]@($input)
    $size = [System.Runtime.InteropServices.Marshal]::SizeOf([type][InputTools2+INPUT])

    [InputTools2]::SendInput([uint32]1, $inputs, $size)
}

$seq = 0

while ($true) {
    $seq++
    $before = Get-MousePos

    do {
        $dx = Get-Random -Minimum -5 -Maximum 6
        $dy = Get-Random -Minimum -5 -Maximum 6
    } while ($dx -eq 0 -and $dy -eq 0)

    $sent = Send-MouseMove -dx $dx -dy $dy
    [InputTools2]::SetCursorPos($before.X + $dx, $before.Y + $dy) | Out-Null
    $after = Get-MousePos

    if ($Log) {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Output "$ts | move#$seq | before=($($before.X),$($before.Y)) | delta=($dx,$dy) | after=($($after.X),$($after.Y)) | sent=$sent"
    }

    Start-Sleep -Seconds $Seconds
}