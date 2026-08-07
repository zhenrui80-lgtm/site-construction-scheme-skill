$chrome = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
$root = 'C:\Users\Administrator\Desktop\Metallurgical_BOQ_AgentSkill_Fixed\agents\_work_descaling_system_revision'
$assets = Join-Path $root 'assets'

$pages = @(
    @('raw_20801.png', 'https://std.samr.gov.cn/gb/search/gbDetailed?id=42BA7D06A208E936E06397BE0A0ACDC9'),
    @('raw_tsg31.png', 'https://www.samr.gov.cn/zw/zfxxgk/fdzdgknr/tzsbs/art/2025/art_8d38f20133a94308b9dd72e160c0692b.html'),
    @('raw_tsgz6002.png', 'https://www.samr.gov.cn/zw/zfxxgk/fdzdgknr/tzsbs/art/2026/art_4153ad05135445e1b342d6a80814ad72.html'),
    @('raw_50235.png', 'https://ebook.chinabuilding.com.cn/zbooklib/book/detail/show?SiteID=1&bookID=61062&isEnterprise=Y'),
    @('raw_50205.png', 'https://ebook.chinabuilding.com.cn/zbooklib/book/detail/show?SiteID=1&bookID=126822'),
    @('raw_jgjt46.png', 'https://ebook.chinabuilding.com.cn/zbooklib/book/detail/show?SiteID=1&bookID=155877'),
    @('raw_50168.png', 'https://ebook.chinabuilding.com.cn/zbooklib/bookpdf/probation?SiteID=1&bookID=109848')
)

$index = 0
foreach ($page in $pages) {
    $index += 1
    $output = Join-Path $assets $page[0]
    $profile = Join-Path $root ("chrome_profile_{0}" -f $index)
    $args = @(
        '--headless=new',
        '--disable-gpu',
        '--hide-scrollbars',
        '--log-level=3',
        '--window-size=1440,1100',
        '--virtual-time-budget=8000',
        "--user-data-dir=$profile",
        "--screenshot=$output",
        $page[1]
    )
    Start-Process -FilePath $chrome -ArgumentList $args -Wait -WindowStyle Hidden
    if (-not (Test-Path -LiteralPath $output)) {
        throw "Screenshot failed: $($page[1])"
    }
    Get-Item -LiteralPath $output | ForEach-Object { "{0}|{1}" -f $_.Name, $_.Length }
}
