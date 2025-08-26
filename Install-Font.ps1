# Install-Font.ps1
# PowerShell 5.1

$fontFilePath = "\\SERVER PATH TO FONT"

$ErrorActionPreference = 'Stop'

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
if (-not (Test-Admin)) { exit 3 }

try {
    $fontFilePath = $fontFilePath.TrimEnd('\\','/')
    $fontFileName = [System.IO.Path]::GetFileName($fontFilePath)
    $fontNameCore = [System.IO.Path]::GetFileNameWithoutExtension($fontFilePath)
    $ext          = ([System.IO.Path]::GetExtension($fontFilePath)).ToLowerInvariant()
    $fontsDir     = Join-Path $env:WINDIR 'Fonts'
} catch { exit 2 }

if (-not (Test-Path -LiteralPath $fontFilePath)) { exit 2 }

function Set-FontRegistryValue {
    param($DisplayName,$FileName)
    $base  = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
    $key   = $base.CreateSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
    $key.SetValue($DisplayName,$FileName,[Microsoft.Win32.RegistryValueKind]::String)
    $key.Close(); $base.Close()
}

# Import SendMessageTimeout without 'using' directives
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class FontRefresher {
    public const int HWND_BROADCAST = 0xffff;
    public const int WM_FONTCHANGE  = 0x001D;
    public const int SMTO_ABORTIFHUNG = 0x0002;
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    public static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam,
        uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
}
"@

function Invoke-FontChangeBroadcast {
    $out = [UIntPtr]::Zero
    [void][FontRefresher]::SendMessageTimeout([IntPtr][FontRefresher]::HWND_BROADCAST,
        [uint32][FontRefresher]::WM_FONTCHANGE,[UIntPtr]::Zero,[IntPtr]::Zero,
        [uint32][FontRefresher]::SMTO_ABORTIFHUNG,2000,[ref]$out)
}

$suffix = switch ($ext) {
    '.otf' { '(OpenType)' }
    '.ttf' { '(TrueType)' }
    default { '(TrueType)' }
}
$displayName = "$fontNameCore $suffix"

$destPath = Join-Path $fontsDir $fontFileName

try {
    if (-not (Test-Path -LiteralPath $destPath)) {
        Copy-Item -LiteralPath $fontFilePath -Destination $destPath -Force
    }
    Set-FontRegistryValue -DisplayName $displayName -FileName $fontFileName
    Invoke-FontChangeBroadcast
    exit 0
}
catch {
    Write-Error ("Font install failed: {0}" -f $_.Exception.Message)
    exit 1
}
