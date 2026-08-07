param(
    [string]$OutputDir = "C:\Users\Administrator\Desktop\Metallurgical_BOQ_AgentSkill_Fixed\agents\_work_qc_canteen_gas_to_electric\assets"
)

$ErrorActionPreference = "Stop"

function Get-Rgb([int]$R, [int]$G, [int]$B) {
    return $R + (256 * $G) + (65536 * $B)
}

function Set-TextStyle($Shape, [string]$Text, [double]$FontSize, [int]$Color, [bool]$Bold = $false, [int]$Alignment = 2) {
    $Shape.TextFrame.TextRange.Text = $Text
    $Shape.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"
    $Shape.TextFrame.TextRange.Font.Name = "Arial"
    $Shape.TextFrame.TextRange.Font.Size = $FontSize
    $Shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $Shape.TextFrame.TextRange.Font.Bold = [int]$Bold
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Alignment
    $Shape.TextFrame.VerticalAnchor = 3
    $Shape.TextFrame.WordWrap = -1
    $Shape.TextFrame.MarginLeft = 4
    $Shape.TextFrame.MarginRight = 4
    $Shape.TextFrame.MarginTop = 2
    $Shape.TextFrame.MarginBottom = 2
}

function Add-Text($Slide, [double]$X, [double]$Y, [double]$W, [double]$H, [string]$Text, [double]$FontSize, [int]$Color, [bool]$Bold = $false, [int]$Alignment = 2) {
    $shape = $Slide.Shapes.AddTextbox(1, $X, $Y, $W, $H)
    $shape.Line.Visible = 0
    $shape.Fill.Visible = 0
    Set-TextStyle $shape $Text $FontSize $Color $Bold $Alignment
    return $shape
}

function Add-Box($Slide, [double]$X, [double]$Y, [double]$W, [double]$H, [string]$Text, [int]$FillColor, [int]$LineColor, [double]$FontSize = 13, [int]$TextColor = 0, [bool]$Bold = $false) {
    $shape = $Slide.Shapes.AddShape(1, $X, $Y, $W, $H)
    $shape.Fill.ForeColor.RGB = $FillColor
    $shape.Line.ForeColor.RGB = $LineColor
    $shape.Line.Weight = 1.25
    Set-TextStyle $shape $Text $FontSize $TextColor $Bold 2
    return $shape
}

function Add-Line($Slide, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [int]$Color, [double]$Weight = 2, [bool]$Arrow = $false, [int]$DashStyle = 1) {
    $line = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
    $line.Line.ForeColor.RGB = $Color
    $line.Line.Weight = $Weight
    $line.Line.DashStyle = $DashStyle
    if ($Arrow) {
        $line.Line.EndArrowheadStyle = 3
    }
    return $line
}

function Add-Step($Slide, [double]$X, [double]$Y, [string]$Number, [string]$Title, [string]$Body, [int]$Accent) {
    $circle = $Slide.Shapes.AddShape(9, $X, $Y, 38, 38)
    $circle.Fill.ForeColor.RGB = $Accent
    $circle.Line.ForeColor.RGB = $Accent
    Set-TextStyle $circle $Number 15 (Get-Rgb 255 255 255) $true 2

    $box = Add-Box $Slide ($X + 48) ($Y - 2) 205 72 "$Title`n$Body" (Get-Rgb 247 249 251) (Get-Rgb 166 181 195) 12 (Get-Rgb 35 46 57) $false
    $box.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
    $box.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $Accent
    return $box
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$navy = Get-Rgb 31 78 120
$blue = Get-Rgb 47 117 181
$lightBlue = Get-Rgb 221 235 247
$orange = Get-Rgb 237 125 49
$red = Get-Rgb 192 0 0
$green = Get-Rgb 112 173 71
$yellow = Get-Rgb 255 192 0
$gray = Get-Rgb 127 127 127
$lightGray = Get-Rgb 242 244 247
$dark = Get-Rgb 35 46 57
$white = Get-Rgb 255 255 255

$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Add($true)
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540

    # Slide 1: overall gas-to-electric retrofit route
    $slide1 = $presentation.Slides.Add(1, 12)
    $background1 = $slide1.Shapes.AddShape(1, 0, 0, 960, 540)
    $background1.Fill.ForeColor.RGB = $white
    $background1.Line.Visible = 0
    $background1.ZOrder(1)

    $titleBar1 = $slide1.Shapes.AddShape(1, 0, 0, 960, 58)
    $titleBar1.Fill.ForeColor.RGB = $navy
    $titleBar1.Line.Visible = 0
    Set-TextStyle $titleBar1 "质检中心食堂气改电总体施工示意" 24 $white $true 2
    Add-Text $slide1 25 61 910 28 "原燃气系统安全退出 → 新低压配电系统形成 → 厨房设备接电试运行" 13 $navy $false 2 | Out-Null

    # Three main zones
    $gasZone = Add-Box $slide1 28 104 256 328 "" (Get-Rgb 252 245 244) (Get-Rgb 225 155 146) 12 $dark $false
    $elecZone = Add-Box $slide1 352 104 276 328 "" (Get-Rgb 242 247 252) (Get-Rgb 133 178 215) 12 $dark $false
    $equipZone = Add-Box $slide1 696 104 236 328 "" (Get-Rgb 244 249 241) (Get-Rgb 155 195 128) 12 $dark $false

    Add-Text $slide1 40 112 232 30 "一、原燃气系统安全退出" 17 $red $true 2 | Out-Null
    Add-Text $slide1 364 112 252 30 "二、新配电系统安装" 17 $blue $true 2 | Out-Null
    Add-Text $slide1 708 112 212 30 "三、厨房设备接电" 17 $green $true 2 | Out-Null

    # Gas route
    Add-Line $slide1 55 205 257 205 $red 8 $false | Out-Null
    $valve = $slide1.Shapes.AddShape(7, 88, 187, 38, 36)
    $valve.Fill.ForeColor.RGB = $white
    $valve.Line.ForeColor.RGB = $red
    $valve.Line.Weight = 2
    Add-Text $slide1 58 151 196 30 "DN120 原燃气管道" 14 $red $true 2 | Out-Null
    Add-Text $slide1 51 229 212 34 "关闭阀门并挂牌上锁" 12 $dark $false 2 | Out-Null

    # Blind flange
    Add-Line $slide1 160 183 160 227 $red 4 $false | Out-Null
    Add-Line $slide1 169 183 169 227 $red 4 $false | Out-Null
    Add-Text $slide1 126 252 80 28 "可靠隔断" 12 $red $true 2 | Out-Null

    # Purge cylinder / detector
    $cyl = Add-Box $slide1 52 303 50 82 "N₂" (Get-Rgb 224 236 255) $blue 18 $blue $true
    $cyl.AutoShapeType = 5
    Add-Line $slide1 102 326 154 326 $blue 3 $true | Out-Null
    Add-Line $slide1 154 213 154 326 $blue 3 $false | Out-Null
    $detector = Add-Box $slide1 186 295 70 58 "气体检测" (Get-Rgb 255 242 204) $orange 12 $dark $true
    Add-Line $slide1 221 221 221 295 $orange 2 $false 4 | Out-Null
    Add-Text $slide1 44 391 226 30 "置换合格后分段拆除、封口、转运" 12 $dark $false 2 | Out-Null

    # Electrical system
    $cab1 = Add-Box $slide1 376 248 76 130 "GGD`n1#" (Get-Rgb 214 226 243) $navy 16 $navy $true
    $cab2 = Add-Box $slide1 462 248 76 130 "GGD`n2#" (Get-Rgb 214 226 243) $navy 16 $navy $true
    for ($i = 0; $i -lt 4; $i++) {
        Add-Line $slide1 (390 + $i * 17) 280 (390 + $i * 17) 358 (Get-Rgb 150 165 180) 0.75 $false | Out-Null
        Add-Line $slide1 (476 + $i * 17) 280 (476 + $i * 17) 358 (Get-Rgb 150 165 180) 0.75 $false | Out-Null
    }
    Add-Text $slide1 372 205 170 32 "低压配电柜" 14 $navy $true 2 | Out-Null

    # Busway and cable tray
    $bus = Add-Box $slide1 365 154 236 34 "封闭铜母线 20m" (Get-Rgb 255 230 153) $orange 13 $dark $true
    Add-Line $slide1 414 188 414 245 $orange 4 $true | Out-Null
    Add-Line $slide1 500 188 500 245 $orange 4 $true | Out-Null
    $tray = Add-Box $slide1 556 219 48 159 "桥架`n200×100" $lightBlue $blue 13 $navy $true
    for ($i = 0; $i -lt 6; $i++) {
        Add-Line $slide1 563 (237 + $i * 22) 597 (237 + $i * 22) $blue 1 $false | Out-Null
    }
    Add-Line $slide1 538 300 556 300 $blue 5 $true | Out-Null
    Add-Text $slide1 371 386 230 28 "电缆敷设、终端制作、编号、防火封堵" 11 $dark $false 2 | Out-Null

    # Kitchen equipment
    Add-Line $slide1 604 300 714 300 $blue 5 $true | Out-Null
    $oven = Add-Box $slide1 720 214 82 156 "电烤箱`n/蒸箱" (Get-Rgb 230 239 218) $green 15 $dark $true
    $stove = Add-Box $slide1 814 246 90 124 "电磁灶`n/加热设备" (Get-Rgb 230 239 218) $green 14 $dark $true
    Add-Line $slide1 760 300 814 300 $blue 4 $true | Out-Null
    Add-Text $slide1 716 382 194 30 "逐台核相、试运行、记录电流及温升" 11 $dark $false 2 | Out-Null

    # Earthing and equipotential bonding
    Add-Line $slide1 392 408 882 408 $green 4 $false | Out-Null
    Add-Line $slide1 414 378 414 408 $green 3 $true | Out-Null
    Add-Line $slide1 500 378 500 408 $green 3 $true | Out-Null
    Add-Line $slide1 580 378 580 408 $green 3 $true | Out-Null
    Add-Line $slide1 760 370 760 408 $green 3 $true | Out-Null
    Add-Line $slide1 858 370 858 408 $green 3 $true | Out-Null
    Add-Text $slide1 559 411 300 24 "PE保护接地及等电位联结（黄绿标识）" 11 $green $true 2 | Out-Null

    # Section arrows
    $arrow1 = $slide1.Shapes.AddShape(33, 296, 235, 42, 60)
    $arrow1.Fill.ForeColor.RGB = $orange
    $arrow1.Line.Visible = 0
    $arrow2 = $slide1.Shapes.AddShape(33, 641, 235, 42, 60)
    $arrow2.Fill.ForeColor.RGB = $orange
    $arrow2.Line.Visible = 0

    $footer1 = $slide1.Shapes.AddShape(1, 0, 474, 960, 66)
    $footer1.Fill.ForeColor.RGB = $lightGray
    $footer1.Line.Visible = 0
    Set-TextStyle $footer1 "控制原则：燃气介质和隔断方式先确认；置换检测不合格不拆除；接地和保护试验不合格不送电；负荷试运行异常不移交。" 14 $dark $true 2

    # Slide 2: gas isolation and nitrogen purge sequence
    $slide2 = $presentation.Slides.Add(2, 12)
    $background2 = $slide2.Shapes.AddShape(1, 0, 0, 960, 540)
    $background2.Fill.ForeColor.RGB = $white
    $background2.Line.Visible = 0
    $background2.ZOrder(1)

    $titleBar2 = $slide2.Shapes.AddShape(1, 0, 0, 960, 58)
    $titleBar2.Fill.ForeColor.RGB = $navy
    $titleBar2.Line.Visible = 0
    Set-TextStyle $titleBar2 "原燃气管道停气、隔断与氮气置换工序示意" 24 $white $true 2
    Add-Text $slide2 30 63 900 24 "适用于拆除前安全准备；实际隔断位置、介质性质和检测标准以业主能源管理单位书面确认为准" 12 $red $true 2 | Out-Null

    # Central process pipe
    Add-Line $slide2 80 230 880 230 $red 11 $false | Out-Null
    Add-Text $slide2 310 184 340 30 "待拆除燃气管段（DN120）" 15 $red $true 2 | Out-Null

    # Upstream valve and double blind
    $upValve = $slide2.Shapes.AddShape(7, 138, 210, 42, 40)
    $upValve.Fill.ForeColor.RGB = $white
    $upValve.Line.ForeColor.RGB = $red
    $upValve.Line.Weight = 2
    Add-Text $slide2 107 255 102 28 "上游阀门挂牌上锁" 11 $dark $false 2 | Out-Null
    Add-Line $slide2 230 205 230 255 $red 5 $false | Out-Null
    Add-Line $slide2 240 205 240 255 $red 5 $false | Out-Null
    Add-Text $slide2 200 260 72 28 "盲板隔断" 11 $red $true 2 | Out-Null

    # Nitrogen bottles and manifold
    for ($i = 0; $i -lt 3; $i++) {
        $cylinder = Add-Box $slide2 (68 + $i * 39) 326 30 88 "N₂" (Get-Rgb 224 236 255) $blue 11 $blue $true
        $cylinder.AutoShapeType = 5
    }
    Add-Line $slide2 83 326 161 326 $blue 3 $false | Out-Null
    Add-Line $slide2 161 326 205 326 $blue 3 $true | Out-Null
    $regulator = Add-Box $slide2 205 306 76 40 "减压/汇流" $lightBlue $blue 11 $navy $true
    Add-Line $slide2 281 326 320 326 $blue 3 $true | Out-Null
    Add-Line $slide2 320 238 320 326 $blue 3 $false | Out-Null
    Add-Text $slide2 62 421 225 28 "氮气瓶组固定，减压后缓慢置换" 11 $dark $false 2 | Out-Null

    # Flow arrows along pipe
    Add-Line $slide2 330 230 430 230 $blue 4 $true | Out-Null
    Add-Line $slide2 470 230 570 230 $blue 4 $true | Out-Null
    Add-Line $slide2 610 230 710 230 $blue 4 $true | Out-Null
    Add-Text $slide2 394 249 210 25 "置换方向" 12 $blue $true 2 | Out-Null

    # Detector and vent point
    $detector2 = Add-Box $slide2 600 315 110 60 "可燃气体 / O₂`n连续检测" (Get-Rgb 255 242 204) $orange 12 $dark $true
    Add-Line $slide2 655 242 655 315 $orange 2 $false 4 | Out-Null
    $vent = Add-Box $slide2 775 310 110 60 "安全放散点`n警戒与监护" (Get-Rgb 252 228 214) $red 12 $dark $true
    Add-Line $slide2 830 238 830 310 $red 3 $false | Out-Null

    # Downstream valve
    $downValve = $slide2.Shapes.AddShape(7, 726, 210, 42, 40)
    $downValve.Fill.ForeColor.RGB = $white
    $downValve.Line.ForeColor.RGB = $red
    $downValve.Line.Weight = 2
    Add-Text $slide2 700 255 94 28 "下游隔断确认" 11 $dark $false 2 | Out-Null

    # Step rail at bottom
    $steps = @(
        @("1", "作业许可", "确认介质、范围及停气窗口"),
        @("2", "关闭隔断", "阀门挂牌上锁并安装盲板"),
        @("3", "氮气置换", "低点进、高点出，缓慢连续"),
        @("4", "检测确认", "按介质标准检测并签字"),
        @("5", "分段拆除", "复测合格，防火花、专人监护")
    )
    $stepX = 42
    foreach ($step in $steps) {
        $circle = $slide2.Shapes.AddShape(9, $stepX, 458, 28, 28)
        $circle.Fill.ForeColor.RGB = $navy
        $circle.Line.Visible = 0
        Set-TextStyle $circle $step[0] 11 $white $true 2
        Add-Text $slide2 ($stepX + 32) 450 136 20 $step[1] 11 $navy $true 1 | Out-Null
        Add-Text $slide2 ($stepX + 32) 470 136 40 $step[2] 9 $dark $false 1 | Out-Null
        if ($stepX -lt 750) {
            Add-Line $slide2 ($stepX + 152) 472 ($stepX + 170) 472 $orange 2 $true | Out-Null
        }
        $stepX += 178
    }

    Add-Text $slide2 320 394 310 33 "检测未合格：继续置换，不得拆除" 14 $red $true 2 | Out-Null

    $pptxPath = Join-Path $OutputDir "质检中心食堂气改电施工示意图源文件.pptx"
    $presentation.SaveAs($pptxPath, 24)
    $slide1.Export((Join-Path $OutputDir "总体施工示意图.png"), "PNG", 1920, 1080)
    $slide2.Export((Join-Path $OutputDir "燃气停用置换示意图.png"), "PNG", 1920, 1080)
}
finally {
    if ($presentation -ne $null) {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    if ($ppt -ne $null) {
        $ppt.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Get-Item -LiteralPath (Join-Path $OutputDir "总体施工示意图.png"), (Join-Path $OutputDir "燃气停用置换示意图.png") |
    Select-Object Name, Length


