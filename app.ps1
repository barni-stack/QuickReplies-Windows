$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$instanceLock = $null
try {
$lockDir = Join-Path $env:LOCALAPPDATA 'QuickReplies'
[void][IO.Directory]::CreateDirectory($lockDir)
try {
    $instanceLock = [IO.File]::Open((Join-Path $lockDir 'instance.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
} catch [IO.IOException] {
    [void][Windows.Forms.MessageBox]::Show('Другая копия программы ещё работает. Закройте её окно или завершите процесс Windows PowerShell, запущенный для Быстрых ответов.','Быстрые ответы')
    return
}
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$dir = Join-Path $env:LOCALAPPDATA 'QuickReplies'
[IO.Directory]::CreateDirectory($dir) | Out-Null
$file = Join-Path $dir 'templates.json'
$items = @()
function Read-Items($path) {
    $raw = [IO.File]::ReadAllText($path)
    if (!$raw.TrimStart().StartsWith('[')) { throw 'Нужен файл шаблонов в формате JSON.' }
    foreach ($row in (ConvertFrom-Json -InputObject $raw)) {
        if ($row.title -isnot [string] -or !$row.title.Trim() -or $row.body -isnot [string] -or !$row.body.Trim()) { throw 'В файле есть запись без названия или текста.' }
        [pscustomobject]@{ title=$row.title; body=$row.body; folder=[string]$row.folder; pinned=($row.pinned -eq $true) }
    }
}
function Warn($message) { [Windows.Forms.MessageBox]::Show($message,'Быстрые ответы') | Out-Null }
if (Test-Path $file) {
    try { $items = @(Read-Items $file) } catch { Warn "Не удалось прочитать $file. $_"; return }
}
function Commit($next) {
    try {
        $json = ConvertTo-Json -InputObject @($next) -Depth 4
        [IO.File]::WriteAllText("$file.tmp",$json)
        if ([IO.File]::Exists($file)) { [IO.File]::Replace("$file.tmp",$file,"$file.bak") }
        else { [IO.File]::Move("$file.tmp",$file) }
        $script:items = @($next)
        Refresh-List
        return $true
    } catch { Warn "Сохранение не удалось. $_"; return $false }
}
function Control($type,$parent,$caption,$x,$y,$w,$h) {
    $c = New-Object "Windows.Forms.$type"
    $c.Text=$caption
    $c.SetBounds($x,$y,$w,$h)
    $parent.Controls.Add($c)
    return $c
}
function Window($caption,$width,$height) {
    $f = New-Object Windows.Forms.Form
    $f.Text=$caption
    $f.Font=New-Object Drawing.Font('Segoe UI',10)
    $f.ClientSize=New-Object Drawing.Size($width,$height)
    $f.StartPosition='CenterScreen'
    return $f
}
$f=Window 'Быстрые ответы' 794 530
$f.MinimumSize=$f.Size
$f.KeyPreview=$true
$label=Control Label $f 'Поиск по названию, папке или тексту' 16 14 480 22
$q=Control TextBox $f '' 16 39 490 30
$new=Control Button $f '+ Новый шаблон' 524 34 152 36
$top=Control CheckBox $f 'Поверх окон' 524 76 162 28
$list=Control ListBox $f '' 16 86 490 242
$list.Anchor='Top,Bottom,Left,Right'
$list.DisplayMember='Label'
$preview=Control TextBox $f '' 16 338 490 142
$preview.Multiline=$true
$preview.ReadOnly=$true
$preview.ScrollBars='Vertical'
$preview.Anchor='Bottom,Left,Right'
function Side($caption,$y) {
    $b=Control Button $f $caption 524 $y 252 36
    $b.Anchor='Top,Right'
    return $b
}
$new.Anchor='Top,Right'
$top.Anchor='Top,Right'
$copy=Side 'Копировать  •  Enter' 119
$edit=Side 'Редактировать' 161
$pin=Side 'Закрепить' 203
$del=Side 'Удалить' 245
$export=Side 'Сохранить копию…' 338
$import=Side 'Восстановить из копии…' 380
$status=Control Label $f '' 16 494 751 25
$status.Anchor='Bottom,Left,Right'
function Selection {
    $s=$list.SelectedItem
    $preview.Text=if ($s) { $s.Row.body } else { '' }
    foreach ($b in @($copy,$edit,$pin,$del)) { $b.Enabled=($null -ne $s) }
    $pin.Text=if ($s -and $s.Row.pinned) { 'Открепить' } else { 'Закрепить' }
}
function Refresh-List {
    $list.BeginUpdate()
    $list.Items.Clear()
    foreach ($row in ($script:items | Sort-Object @{Expression='pinned';Descending=$true},title)) {
        $haystack="$($row.title) $($row.folder) $($row.body)"
        if ($haystack.IndexOf($q.Text.Trim(),[StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        $prefix=if ($row.pinned) { '★ ' } else { '' }
        $suffix=if ($row.folder) { "   ·   $($row.folder)" } else { '' }
        [void]$list.Items.Add([pscustomobject]@{Label="$prefix$($row.title)$suffix";Row=$row})
    }
    if ($list.Items.Count) { $list.SelectedIndex=0 }
    $list.EndUpdate()
    Selection
    $status.Text=if (!$script:items.Count) { 'Добавьте первый рабочий ответ кнопкой «Новый шаблон».' } else { "Найдено: $($list.Items.Count)    •    Ctrl+F — поиск" }
}
function Editor($old) {
    $dialog=Window 'Шаблон ответа' 540 434
    $dialog.FormBorderStyle='FixedDialog'
    $dialog.MaximizeBox=$false
    $null=Control Label $dialog 'Название' 16 12 450 22
    $name=Control TextBox $dialog ([string]$old.title) 16 36 506 30
    $null=Control Label $dialog 'Папка (по желанию)' 16 76 450 22
    $folder=Control TextBox $dialog ([string]$old.folder) 16 100 506 30
    $null=Control Label $dialog 'Текст для вставки' 16 140 450 22
    $body=Control TextBox $dialog '' 16 164 506 205
    $body.Multiline=$true
    $body.MaxLength=0
    $body.ScrollBars='Vertical'
    $body.Text=[string]$old.body
    $ok=Control Button $dialog 'Сохранить' 268 385 126 36
    $cancel=Control Button $dialog 'Отмена' 400 385 122 36
    $cancel.DialogResult='Cancel'
    $dialog.CancelButton=$cancel
    $ok.Add_Click({
        if (!$name.Text.Trim() -or !$body.Text.Trim()) { Warn 'Введите название и текст ответа.'; return }
        $row=[pscustomobject]@{title=$name.Text.Trim(); folder=$folder.Text.Trim(); body=$body.Text; pinned=[bool]$old.pinned}
        $next=@($script:items | Where-Object { $_ -ne $old }) + $row
        if (Commit $next) { $dialog.Close() }
    })
    [void]$dialog.ShowDialog($f)
    $dialog.Dispose()
}
function Copy-Current {
    if ($list.SelectedItem) {
        try { [Windows.Forms.Clipboard]::SetText($list.SelectedItem.Row.body); $status.Text="Скопировано: $($list.SelectedItem.Row.title)   •   Вставьте ответ через Ctrl+V" }
        catch { Warn 'Буфер обмена занят. Попробуйте ещё раз.' }
    }
}
$new.Add_Click({ Editor $null })
$edit.Add_Click({ Editor $list.SelectedItem.Row })
$copy.Add_Click({ Copy-Current })
$list.Add_DoubleClick({ Copy-Current })
$pin.Add_Click({
    $row=$list.SelectedItem.Row
    $next=@(foreach ($r in $script:items) {
        if ($r -eq $row) { [pscustomobject]@{title=$r.title;folder=$r.folder;body=$r.body;pinned=!$r.pinned} } else { $r }
    })
    [void](Commit $next)
})
$del.Add_Click({
    $row=$list.SelectedItem.Row
    if ([Windows.Forms.MessageBox]::Show("Удалить «$($row.title)»?",'Удаление','YesNo') -eq 'Yes') {
        [void](Commit @($script:items | Where-Object { $_ -ne $row }))
    }
})
$export.Add_Click({
    $d=New-Object Windows.Forms.SaveFileDialog
    $d.Filter='Копия шаблонов (*.json)|*.json'
    $d.FileName='Ответы.json'
    if ($d.ShowDialog($f) -eq 'OK') {
        try { [IO.File]::WriteAllText($d.FileName,(ConvertTo-Json -InputObject @($script:items) -Depth 4)); $status.Text='Копия сохранена.' } catch { Warn $_.Exception.Message }
    }
    $d.Dispose()
})
$import.Add_Click({
    $d=New-Object Windows.Forms.OpenFileDialog
    $d.Filter='Копия шаблонов (*.json)|*.json'
    if ($d.ShowDialog($f) -eq 'OK') {
        try {
            $loaded=@(Read-Items $d.FileName)
            if ([Windows.Forms.MessageBox]::Show("Заменить текущие шаблоны ($($script:items.Count)) записями из копии ($($loaded.Count))?",'Восстановление','YesNo') -eq 'Yes') { [void](Commit $loaded) }
        } catch { Warn $_.Exception.Message }
    }
    $d.Dispose()
})
$q.Add_TextChanged({ Refresh-List })
$list.Add_SelectedIndexChanged({ Selection })
$top.Add_CheckedChanged({ $f.TopMost=$top.Checked })
$f.Add_KeyDown({
    param($sender,$e)
    if ($e.Control -and $e.KeyCode -eq 'F') { $q.Focus(); $q.SelectAll(); $e.SuppressKeyPress=$true }
    elseif ($e.KeyCode -eq 'Enter' -and ($q.Focused -or $list.Focused)) { Copy-Current; $e.SuppressKeyPress=$true }
    elseif ($q.Focused -and $e.KeyCode -eq 'Down' -and $list.Items.Count) { $list.Focus(); $e.SuppressKeyPress=$true }
})
Refresh-List
$f.Add_Shown({ $q.Select() })
[void]$f.ShowDialog()

} catch {
    [void][Windows.Forms.MessageBox]::Show("Ошибка запуска: $($_.Exception.Message)`r`nСтрока: $($_.InvocationInfo.ScriptLineNumber)",'Быстрые ответы')
} finally {
    if ($instanceLock) { $instanceLock.Dispose() }
}
