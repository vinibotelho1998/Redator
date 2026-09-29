# Remove o Redator instalado para o usuário atual e desfaz os registros.
$dest = Join-Path $env:LOCALAPPDATA 'Programs\Redator'
Get-Process Redator -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item -Recurse -Force $dest -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Redator.lnk') -ErrorAction SilentlyContinue
Remove-Item (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Redator.lnk') -ErrorAction SilentlyContinue
$classes = 'HKCU:\Software\Classes'
foreach ($ext in '.docx', '.docm', '.dotx', '.odt', '.md', '.markdown', '.txt', '.rtf', '.doc', '.pdf') {
  Remove-Item -Recurse -Force "$classes\Redator$ext" -ErrorAction SilentlyContinue
  Remove-ItemProperty -Path "$classes\$ext\OpenWithProgids" -Name "Redator$ext" -ErrorAction SilentlyContinue
}
Remove-Item -Recurse -Force "$classes\Applications\Redator.exe" -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force 'HKCU:\Software\Redator' -ErrorAction SilentlyContinue
Remove-ItemProperty -Path 'HKCU:\Software\RegisteredApplications' -Name 'Redator' -ErrorAction SilentlyContinue
Add-Type -Namespace Win32 -Name Shell2 -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);'
[Win32.Shell2]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
"Redator removido."
