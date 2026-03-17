Add-Type -AssemblyName System.Drawing, System.Windows.Forms

# --- 初期設定 ---
# レジストリから現在のDPI設定を取得（デフォルト値の計算）
$windowMetrics = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics"
$appliedDPI = if ($windowMetrics.AppliedDPI) { $windowMetrics.AppliedDPI } else { 96 }
$defaultScale = [math]::Round(($appliedDPI / 96) * 100)

$bitmap = New-Object System.Drawing.Bitmap(1, 1)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

# --- フォーム作成 ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "座標・色取得くん"
$Form.Size = "220, 160"
$Form.TopMost = $true # 常に最前面に表示
$Form.FormBorderStyle = "FixedSingle"

$label = New-Object System.Windows.Forms.Label
$label.Location = "10, 10"
$label.Size = "180, 50"
$label.Font = New-Object System.Drawing.Font("Consolas", 11)
$label.Text = "Wait..."

$pullLabel = New-Object System.Windows.Forms.Label
$pullLabel.Location = "10, 70"
$pullLabel.Size = "100, 20"
$pullLabel.Text = "表示スケール:"

$dpiScaling = New-Object System.Windows.Forms.ComboBox
$dpiScaling.Location = "110, 67"
$dpiScaling.Size = "70, 30"
$dpiScaling.DropDownStyle = "DropDownList"
@("100%", "125%", "150%", "175%", "200%") | ForEach-Object { [void]$dpiScaling.Items.Add($_) }
$dpiScaling.Text = "$($defaultScale)%"

# --- メインロジック（タイマー） ---
$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = 150 # 200msより少し速めに設定
$Timer.Add_Tick({
    # コンボボックスから倍率を取得（1.25 など）
    $scaleValue = [double]($dpiScaling.Text.Replace("%", "")) / 100
    
    # カーソル位置の取得
    $pos = [System.Windows.Forms.Cursor]::Position
    
    # 高DPI環境での座標補正（必要に応じて）
    $targetX = [int]($pos.X * $scaleValue)
    $targetY = [int]($pos.Y * $scaleValue)

    try {
        # 画面の色を取得
        $graphics.CopyFromScreen($pos.X, $pos.Y, 0, 0, $bitmap.Size)
        $pixel = $bitmap.GetPixel(0, 0)
        
        $hex = "#{0:X2}{1:X2}{2:X2}" -f $pixel.R, $pixel.G, $pixel.B
        $rgb = "R:{0} G:{1} B:{2}" -f $pixel.R, $pixel.G, $pixel.B
        
        $label.Text = "X: $($pos.X), Y: $($pos.Y)`n$hex`n$rgb"
    } catch {
        $label.Text = "Error: Out of range"
    }
})

# フォームを閉じたらタイマーを止める
$Form.Add_FormClosed({
    $Timer.Stop()
    $graphics.Dispose()
    $bitmap.Dispose()
})

$Form.Controls.AddRange(@($label, $pullLabel, $dpiScaling))
$Timer.Start()
$Form.ShowDialog()