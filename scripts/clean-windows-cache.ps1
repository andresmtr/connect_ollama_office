$paths = @(
  "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
  "$env:LOCALAPPDATA\Microsoft\Office\Wef",
  "$env:LOCALAPPDATA\Microsoft\Office\16.0\WebView2",
  "$env:LOCALAPPDATA\Microsoft\Office\WebView2"
)

$removed = 0
foreach ($p in $paths) {
  if (Test-Path $p) {
    try {
      Remove-Item -Recurse -Force $p
      Write-Host "Removed: $p"
      $removed++
    } catch {
      Write-Host "Failed to remove $p : $($_.Exception.Message)"
    }
  }
}

if ($removed -eq 0) {
  Write-Host "No cache folders removed."
}
