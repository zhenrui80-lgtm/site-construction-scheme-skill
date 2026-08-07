param(
    [Parameter(Mandatory = $true)] [string] $Name,
    [Parameter(Mandatory = $true)] [string] $Url
)

$chrome = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
$root = 'C:\Users\Administrator\Desktop\Metallurgical_BOQ_AgentSkill_Fixed\agents\_work_descaling_system_revision'
$output = Join-Path (Join-Path $root 'assets') $Name
$profile = Join-Path $root ('chrome_' + [System.IO.Path]::GetFileNameWithoutExtension($Name))
$args = @(
    '--headless=new',
    '--disable-gpu',
    '--hide-scrollbars',
    '--log-level=3',
    '--window-size=1440,1100',
    '--virtual-time-budget=8000',
    "--user-data-dir=$profile",
    "--screenshot=$output",
    $Url
)
Start-Process -FilePath $chrome -ArgumentList $args -Wait -WindowStyle Hidden
if (-not (Test-Path -LiteralPath $output)) {
    throw "Screenshot failed: $Url"
}
Get-Item -LiteralPath $output | ForEach-Object { "{0}|{1}" -f $_.Name, $_.Length }
