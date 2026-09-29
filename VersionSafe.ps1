<#
.SYNOPSIS
    Версионное резервное копирование папки с дедупликацией через жёсткие ссылки.

.DESCRIPTION
    Каждый запуск создаёт в папке назначения новый снимок — папку с именем
    вида yyyy-MM-ddTHH-mm-ss. Файлы, у которых размер и время изменения
    совпадают с предыдущим снимком, не копируются, а связываются с ним
    жёсткими ссылками, поэтому новый снимок занимает место только под
    изменённые файлы (принцип Apple Time Machine / rsync --link-dest).

    Жизненный цикл запуска:
      1. Проверки: источник и назначение не пересекаются, скрипт не запущен
         повторно (файл VersionSafe.lock), назначение поддерживает ссылки.
      2. Удаляются папки *.inProgress, оставшиеся от прерванных запусков.
      3. (Опционально) проверяются и восстанавливаются разорванные ссылки.
      4. Снимок собирается в папке <дата>.inProgress и при успехе
         переименовывается в <дата>. При сбое папка удаляется.
      5. Только после успешного снимка выполняется GFS-ротация и очистка логов.

    Символические ссылки и точки соединения (junction) внутри источника
    пропускаются: иначе в снимок попадут чужие данные или возникнет цикл.

    Коды завершения:
      0 — снимок создан без ошибок;
      1 — снимок не создан (скрипт завершился исключением);
      2 — снимок создан, но часть файлов или папок прочитать не удалось.

    Подробная документация — в README.md.

.PARAMETER SourcePath
    Папка-источник. Можно передать через конвейер (только одну папку).

.PARAMETER DestinationPath
    Папка для снимков, желательно на другом физическом диске. Создаётся при
    отсутствии. Не может совпадать с источником или находиться внутри него.

.PARAMETER FailIfOldBackupsAreMissing
    Прервать работу, если в папке назначения нет ни одного готового снимка.
    Защищает от бэкапа «в пустоту», когда внешний диск не подключён.

.PARAMETER NoHardLinks
    Полное копирование без жёстких ссылок (FAT32, exFAT, сетевые папки).
    Включается автоматически, если пробная ссылка в назначении не создаётся.

.PARAMETER NoLogging
    Не записывать лог (транскрипт) запуска.

.PARAMETER NoStatistics
    Не выводить статистику снимка.

.PARAMETER DailyCount
    Сколько хранить самых свежих снимков (по умолчанию 7).

.PARAMETER WeeklyCount
    Сколько хранить недель, по последнему снимку из каждой (по умолчанию 4).

.PARAMETER MonthlyCount
    Сколько хранить месяцев, по последнему снимку из каждого (по умолчанию 12).

.PARAMETER YearlyCount
    Сколько хранить лет, по последнему снимку из каждого (по умолчанию 5).
    Если все четыре счётчика равны 0 — ротация отключается.

.PARAMETER LogRetentionDays
    Сколько дней хранить логи неудачных запусков (у которых нет снимка).
    Лог успешного снимка удаляется вместе со снимком при ротации.
    0 — хранить все (по умолчанию).

.PARAMETER RotationDisabled
    Полностью отключить ротацию (хранить все снимки).

.PARAMETER VerifyLinks
    Перед созданием снимка найти разорванные ссылки: файлы, которые
    совпадают с файлом из предыдущего снимка, но хранятся отдельной копией.

.PARAMETER VerifyLinksOlderThanDays
    Проверять только снимки старше N дней. 0 — все (по умолчанию).

.PARAMETER RepairBrokenLinks
    Восстанавливать найденные разорванные ссылки: если содержимое файлов
    совпадает (SHA-256), копия заменяется жёсткой ссылкой. Включает -VerifyLinks.

.EXAMPLE
    .\VersionSafe.ps1 -SourcePath C:\Data -DestinationPath D:\Backups

.EXAMPLE
    .\VersionSafe.ps1 C:\Data D:\Backups -DailyCount 14 -WeeklyCount 0 -MonthlyCount 6 -YearlyCount 0

.EXAMPLE
    .\VersionSafe.ps1 C:\Data D:\Backups -RepairBrokenLinks -WhatIf

.NOTES
    Лицензия: GPL v3 или новее.
#>

#Requires -Version 5.0

[CmdletBinding(SupportsShouldProcess = $true)]
Param(
    [Parameter(Mandatory, Position = 0, ValueFromPipeline = $true)]
    [Alias("Source", "Path")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [String]$SourcePath,

    [Parameter(Mandatory, Position = 1)]
    [Alias("Destination", "Target")]
    [ValidateNotNullOrEmpty()]
    [String]$DestinationPath,

    [Switch]$FailIfOldBackupsAreMissing = $false,

    [Alias("FullCopy")]
    [Switch]$NoHardLinks = $false,

    [Switch]$NoLogging = $false,

    [Switch]$NoStatistics = $false,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$DailyCount = 7,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$WeeklyCount = 4,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MonthlyCount = 12,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$YearlyCount = 5,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$LogRetentionDays = 0,

    [Switch]$RotationDisabled = $false,

    [Switch]$VerifyLinks = $false,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$VerifyLinksOlderThanDays = 0,

    [Switch]$RepairBrokenLinks = $false
)

begin {
    $script:SourceCount = 0

    #region Вспомогательные функции

    # Жёсткие ссылки и идентификаторы файлов берём напрямую из WinAPI:
    # New-Item -ItemType HardLink в PowerShell 5.1 трактует -Target как шаблон
    # и ломается на путях с [ ], а свойство Target у жёстких ссылок ненадёжно.
    if (-not ('VersionSafe.Native' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace VersionSafe
{
    public static class Native
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct BY_HANDLE_FILE_INFORMATION
        {
            public uint FileAttributes;
            // FILETIME — две половины по 4 байта (выравнивание 4, не 8).
            public uint CreationTimeLow, CreationTimeHigh;
            public uint LastAccessTimeLow, LastAccessTimeHigh;
            public uint LastWriteTimeLow, LastWriteTimeHigh;
            public uint VolumeSerialNumber;
            public uint FileSizeHigh;
            public uint FileSizeLow;
            public uint NumberOfLinks;
            public uint FileIndexHigh;
            public uint FileIndexLow;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CreateHardLink(string lpFileName, string lpExistingFileName, IntPtr lpSecurityAttributes);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string lpFileName, uint dwDesiredAccess, uint dwShareMode,
            IntPtr lpSecurityAttributes, uint dwCreationDisposition, uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle hFile, out BY_HANDLE_FILE_INFORMATION lpFileInformation);

        public static void NewHardLink(string linkPath, string existingPath)
        {
            if (!CreateHardLink(linkPath, existingPath, IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        // Идентификатор файла на томе: одинаков у всех жёстких ссылок одного файла.
        public static string GetFileId(string path)
        {
            // Доступ 0 и общий режим чтения/записи/удаления: хватает для метаданных,
            // не мешает другим процессам. 3 = OPEN_EXISTING, 0x02000000 = BACKUP_SEMANTICS.
            using (SafeFileHandle h = CreateFile(path, 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero))
            {
                if (h.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
                BY_HANDLE_FILE_INFORMATION info;
                if (!GetFileInformationByHandle(h, out info)) throw new Win32Exception(Marshal.GetLastWin32Error());
                return string.Format("{0:X8}-{1:X8}{2:X8}", info.VolumeSerialNumber, info.FileIndexHigh, info.FileIndexLow);
            }
        }
    }
}
'@
    }

    $script:SnapshotNamePattern = '^\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}$'

    function Get-NormalizedPath {
        <#
        .SYNOPSIS
            Возвращает полный путь без завершающего разделителя (кроме корня диска).
        #>
        param([Parameter(Mandatory)][string]$Path)

        $full = [System.IO.Path]::GetFullPath($Path)
        $root = [System.IO.Path]::GetPathRoot($full)
        if ($full.Length -gt $root.Length) { $full = $full.TrimEnd('\', '/') }
        return $full
    }

    function Test-IsSameOrUnder {
        <#
        .SYNOPSIS
            Проверяет, совпадает ли путь $Path с $Parent или лежит внутри него.
            Сравнение идёт по целым компонентам: D:\Backups2 не лежит внутри D:\Backups.
        #>
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$Parent
        )

        if ($Path.Equals($Parent, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        $prefix = $Parent.TrimEnd('\') + '\'
        return $Path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)
    }

    function Get-ErrorText {
        param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
        return $ErrorRecord.Exception.GetBaseException().Message
    }

    function Test-HardLinkSupport {
        <#
        .SYNOPSIS
            Проверяет поддержку жёстких ссылок, создавая пробную ссылку прямо в
            папке назначения. Так корректно определяются и локальные ФС, и
            сетевые папки, и тома без буквы диска.
        #>
        param([Parameter(Mandatory)][string]$Path)

        $probe = [System.IO.Path]::Combine($Path, ".VersionSafe-probe-$([guid]::NewGuid().ToString('N'))")
        $link  = "$probe.link"
        try {
            [System.IO.File]::WriteAllText($probe, 'probe')
            [VersionSafe.Native]::NewHardLink($link, $probe)
            return $true
        } catch {
            Write-Verbose "Пробная жёсткая ссылка не создана: $(Get-ErrorText $_)"
            return $false
        } finally {
            foreach ($p in $link, $probe) {
                try { [System.IO.File]::Delete($p) } catch { }
            }
        }
    }

    function Get-SnapshotList {
        <#
        .SYNOPSIS
            Возвращает готовые снимки (без .inProgress), от новых к старым.
        #>
        param([Parameter(Mandatory)][string]$Root)

        if (-not (Test-Path -LiteralPath $Root -PathType Container)) { return }

        Get-ChildItem -LiteralPath $Root -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -CMatch $script:SnapshotNamePattern } |
            Sort-Object -Property Name -Descending
    }

    function ConvertFrom-SnapshotName {
        param([Parameter(Mandatory)][string]$Name)
        return [datetime]::ParseExact($Name, 'yyyy-MM-ddTHH-mm-ss',
            [System.Globalization.CultureInfo]::InvariantCulture)
    }

    function Invoke-SnapshotCopy {
        <#
        .SYNOPSIS
            Переносит источник в рабочую папку снимка: неизменённые файлы
            связываются жёсткими ссылками с предыдущим снимком, остальные
            копируются. В режиме -DryRun только считает, что было бы сделано.
        #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$SourceRoot,
            [Parameter(Mandatory)][string]$WorkPath,
            [string]$PreviousPath,
            [Switch]$DryRun
        )

        $stats = [ordered]@{
            FilesCopied = 0
            FilesLinked = 0
            BytesCopied = [long]0
            BytesTotal  = [long]0
            Skipped     = 0
            Errors      = 0
        }

        # Обход без рекурсии, чтобы не упереться в глубину стека на больших деревьях.
        # В стеке — пути относительно корня источника.
        $pending = New-Object System.Collections.Generic.Stack[string]
        $pending.Push('')

        while ($pending.Count -gt 0) {
            $relDir = $pending.Pop()
            $srcDir = if ($relDir) { [System.IO.Path]::Combine($SourceRoot, $relDir) } else { $SourceRoot }

            try {
                $children = @(Get-ChildItem -LiteralPath $srcDir -Force -ErrorAction Stop)
            } catch {
                $stats.Errors++
                Write-Warning "Не удалось прочитать папку ${srcDir}: $(Get-ErrorText $_)"
                continue
            }

            foreach ($child in $children) {
                $rel  = if ($relDir) { [System.IO.Path]::Combine($relDir, $child.Name) } else { $child.Name }
                $dest = [System.IO.Path]::Combine($WorkPath, $rel)

                # Символические ссылки и точки соединения не копируем и не обходим.
                # Облачные заглушки (OneDrive) тоже точки повторной обработки, но LinkType у них пустой.
                if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
                    ($child.LinkType -eq 'Junction' -or $child.LinkType -eq 'SymbolicLink')) {
                    $stats.Skipped++
                    Write-Verbose "Пропущена ссылка ($($child.LinkType)): $rel"
                    continue
                }

                try {
                    if ($child.PSIsContainer) {
                        if (-not $DryRun) { [void][System.IO.Directory]::CreateDirectory($dest) }
                        $pending.Push($rel)
                        continue
                    }

                    $stats.BytesTotal += $child.Length
                    $linked = $false

                    if ($PreviousPath) {
                        $prev = New-Object System.IO.FileInfo ([System.IO.Path]::Combine($PreviousPath, $rel))
                        if ($prev.Exists -and
                            $prev.Length -eq $child.Length -and
                            $prev.LastWriteTimeUtc -eq $child.LastWriteTimeUtc) {
                            if ($DryRun) {
                                $linked = $true
                            } else {
                                try {
                                    [VersionSafe.Native]::NewHardLink($dest, $prev.FullName)
                                    $linked = $true
                                } catch {
                                    # Например, достигнут предел NTFS в 1023 ссылки на файл.
                                    Write-Warning "Не удалось создать ссылку для $rel, файл будет скопирован: $(Get-ErrorText $_)"
                                }
                            }
                        }
                    }

                    if ($linked) {
                        $stats.FilesLinked++
                        Write-Verbose "Ссылка: $rel"
                    } else {
                        # File.Copy сохраняет время изменения и атрибуты — на этом строится сравнение.
                        if (-not $DryRun) { [System.IO.File]::Copy($child.FullName, $dest, $false) }
                        $stats.FilesCopied++
                        $stats.BytesCopied += $child.Length
                        Write-Verbose "Копирование: $rel"
                    }
                } catch {
                    # Нехватка места — не ошибка отдельного файла: снимок всё равно не получится.
                    $code = $_.Exception.GetBaseException().HResult -band 0xFFFF
                    if ($code -eq 39 -or $code -eq 112) {
                        throw "На диске назначения закончилось место."
                    }
                    $stats.Errors++
                    Write-Warning "Не удалось обработать $($child.FullName): $(Get-ErrorText $_)"
                }
            }
        }

        return [pscustomobject]$stats
    }

    function Write-SnapshotStatistics {
        param([Parameter(Mandatory)]$Stats)

        $savings = if ($Stats.BytesTotal -gt 0) { 1 - ($Stats.BytesCopied / $Stats.BytesTotal) } else { 0 }
        Write-Information ("Файлов: скопировано {0} ({1:N2} МБ), связано ссылками {2}. Объём источника: {3:N2} МБ. Экономия места: {4:P0}." -f `
            $Stats.FilesCopied, ($Stats.BytesCopied / 1MB), $Stats.FilesLinked, ($Stats.BytesTotal / 1MB), $savings)
        if ($Stats.Skipped -gt 0) {
            Write-Information "Пропущено символических ссылок и точек соединения: $($Stats.Skipped)."
        }
        if ($Stats.Errors -gt 0) {
            Write-Information "Ошибок: $($Stats.Errors)."
        }
    }

    function Invoke-SnapshotCleanup {
        <#
        .SYNOPSIS
            Удаляет старые снимки по схеме GFS: хранится объединение якорей
            всех уровней. Вместе со снимком удаляется его лог.
        #>
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory)][string]$Root,

            [int]$Daily   = 7,
            [int]$Weekly  = 4,
            [int]$Monthly = 12,
            [int]$Yearly  = 5,

            # Снимок, который был бы создан в режиме -WhatIf: участвует в выборе
            # якорей, но на диске его нет, и удалять его не нужно.
            [string]$PlannedSnapshot
        )

        if (($Daily + $Weekly + $Monthly + $Yearly) -eq 0) {
            Write-Verbose "Все счётчики GFS равны нулю — ротация отключена."
            return
        }

        $names = @(Get-SnapshotList -Root $Root | ForEach-Object { $_.Name })
        if ($PlannedSnapshot) { $names += $PlannedSnapshot }

        $parsed = @(foreach ($name in $names) {
            try {
                $date = ConvertFrom-SnapshotName $name
            } catch {
                Write-Warning "Не удалось разобрать дату из имени снимка $name — он не участвует в ротации."
                continue
            }
            [pscustomobject]@{ Name = $name; Date = $date }
        })
        $parsed = @($parsed | Sort-Object -Property Date -Descending)
        if ($parsed.Count -eq 0) {
            Write-Verbose "Нет снимков для ротации."
            return
        }

        # Последний снимок каждой группы, группы от новых к старым, первые $Count.
        $selectAnchors = {
            param([scriptblock]$Key, [int]$Count)
            @($parsed |
                Group-Object -Property $Key |
                ForEach-Object { $_.Group | Sort-Object -Property Date -Descending | Select-Object -First 1 } |
                Sort-Object -Property Date -Descending |
                Select-Object -First $Count)
        }

        $keepers = New-Object 'System.Collections.Generic.HashSet[string]'

        if ($Daily -gt 0) {
            $parsed | Select-Object -First $Daily | ForEach-Object { [void]$keepers.Add($_.Name) }
        }

        if ($Weekly -gt 0) {
            # Неделя по ISO 8601: год и номер недели берутся по четвергу этой недели,
            # поэтому неделя на стыке годов не распадается на две группы.
            $anchors = & $selectAnchors {
                $d = $_.Date.Date
                $thursday = $d.AddDays(3 - (([int]$d.DayOfWeek + 6) % 7))
                '{0:D4}-W{1:D2}' -f $thursday.Year, ([int][math]::Floor(($thursday.DayOfYear - 1) / 7) + 1)
            } $Weekly
            $anchors | ForEach-Object { [void]$keepers.Add($_.Name) }
            Write-Verbose "Недельные якоря: $($anchors.Count) (запрошено $Weekly)."
        }

        if ($Monthly -gt 0) {
            $anchors = & $selectAnchors { '{0:D4}-{1:D2}' -f $_.Date.Year, $_.Date.Month } $Monthly
            $anchors | ForEach-Object { [void]$keepers.Add($_.Name) }
            Write-Verbose "Месячные якоря: $($anchors.Count) (запрошено $Monthly)."
        }

        if ($Yearly -gt 0) {
            $anchors = & $selectAnchors { $_.Date.Year } $Yearly
            $anchors | ForEach-Object { [void]$keepers.Add($_.Name) }
            Write-Verbose "Годовые якоря: $($anchors.Count) (запрошено $Yearly)."
        }

        $victims = @($parsed | Where-Object { -not $keepers.Contains($_.Name) -and $_.Name -ne $PlannedSnapshot })
        if ($victims.Count -eq 0) {
            Write-Verbose "Ротация: удалять нечего."
            return
        }

        Write-Information "Ротация GFS: к удалению $($victims.Count) снимков, остаётся $($keepers.Count)."

        foreach ($v in $victims) {
            $path = Join-Path -Path $Root -ChildPath $v.Name
            if ($PSCmdlet.ShouldProcess($path, "Удалить старый снимок (GFS)")) {
                try {
                    Remove-Item -LiteralPath $path -Recurse -Force -Confirm:$false -ErrorAction Stop
                    Remove-Item -LiteralPath (Join-Path -Path $Root -ChildPath "VersionSafe-$($v.Name).log") `
                        -Force -Confirm:$false -ErrorAction SilentlyContinue
                    Write-Information "Удалён снимок: $($v.Name)"
                } catch {
                    Write-Warning "Не удалось удалить снимок $($v.Name): $(Get-ErrorText $_)"
                }
            }
        }
    }

    function Invoke-LogCleanup {
        <#
        .SYNOPSIS
            Удаляет старые логи неудачных запусков — те, у которых нет снимка.
            Логи существующих снимков удаляются только вместе со снимком.
        #>
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory)][string]$Root,
            [Parameter(Mandatory)][int]$MaxAgeDays,
            [string]$CurrentLog
        )

        if ($MaxAgeDays -le 0) { return }
        if (-not (Test-Path -LiteralPath $Root -PathType Container)) { return }

        $cutoff = (Get-Date).AddDays(-$MaxAgeDays)
        $logs = @(Get-ChildItem -LiteralPath $Root -Filter "VersionSafe-*.log" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff -and $_.FullName -ne $CurrentLog })

        foreach ($log in $logs) {
            $snapName = $log.BaseName.Substring('VersionSafe-'.Length)
            if (Test-Path -LiteralPath (Join-Path -Path $Root -ChildPath $snapName) -PathType Container) { continue }

            if ($PSCmdlet.ShouldProcess($log.FullName, "Удалить старый лог")) {
                try { Remove-Item -LiteralPath $log.FullName -Force -Confirm:$false -ErrorAction Stop }
                catch { Write-Warning "Не удалось удалить лог $($log.Name): $(Get-ErrorText $_)" }
            }
        }
    }

    function Test-SnapshotLinks {
        <#
        .SYNOPSIS
            Ищет разорванные ссылки: файл в снимке совпадает по размеру и времени
            изменения с файлом из предыдущего снимка, но хранится отдельной
            копией, а не жёсткой ссылкой на тот же файл.

        .DESCRIPTION
            Такое бывает, если снимки переносили инструментом без поддержки
            ссылок, или ссылку не удалось создать (например, предел NTFS в 1023
            ссылки на файл). Данные при этом целы — теряется только место.

            При -Repair содержимое обоих файлов сравнивается по SHA-256, и при
            совпадении копия заменяется жёсткой ссылкой. Если содержимое
            различается, это разные версии, и файл не трогается.
        #>
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory)][string]$Root,
            [int]$OlderThanDays = 0,
            [Switch]$Repair
        )

        # От старых к новым: предыдущий снимок для каждого — предыдущий в списке.
        $snapshots = @(Get-SnapshotList -Root $Root | Sort-Object -Property Name)
        if ($snapshots.Count -lt 2) {
            Write-Information "Проверка ссылок: меньше двух снимков — проверять нечего."
            return
        }

        $cutoff = if ($OlderThanDays -gt 0) { (Get-Date).AddDays(-$OlderThanDays) } else { [datetime]::MaxValue }

        $found     = 0
        $fixed     = 0
        $different = 0
        $failed    = 0

        for ($i = 1; $i -lt $snapshots.Count; $i++) {
            $current  = $snapshots[$i]
            $previous = $snapshots[$i - 1]

            try { $date = ConvertFrom-SnapshotName $current.Name } catch { continue }
            if ($date -gt $cutoff) {
                Write-Verbose "Пропускаю свежий снимок $($current.Name)."
                continue
            }

            Write-Verbose "Проверяю снимок $($current.Name) (предыдущий: $($previous.Name))"

            foreach ($file in Get-ChildItem -LiteralPath $current.FullName -Recurse -File -Force -ErrorAction SilentlyContinue) {
                $relPath  = $file.FullName.Substring($current.FullName.Length).TrimStart('\')
                $prevFile = New-Object System.IO.FileInfo ([System.IO.Path]::Combine($previous.FullName, $relPath))

                if (-not $prevFile.Exists -or
                    $prevFile.Length -ne $file.Length -or
                    $prevFile.LastWriteTimeUtc -ne $file.LastWriteTimeUtc) {
                    continue
                }

                try {
                    if ([VersionSafe.Native]::GetFileId($file.FullName) -eq [VersionSafe.Native]::GetFileId($prevFile.FullName)) {
                        continue
                    }
                } catch {
                    Write-Warning "Не удалось прочитать $($file.FullName): $(Get-ErrorText $_)"
                    continue
                }

                $found++
                Write-Information "Разорванная ссылка: [$($current.Name)] $relPath"

                if (-not $Repair) { continue }
                if (-not $PSCmdlet.ShouldProcess($file.FullName, "Заменить копию жёсткой ссылкой на файл из $($previous.Name)")) {
                    continue
                }

                $tmp = "$($file.FullName).vs-relink"
                try {
                    $hashCurrent  = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
                    $hashPrevious = (Get-FileHash -LiteralPath $prevFile.FullName -Algorithm SHA256).Hash
                    if ($hashCurrent -ne $hashPrevious) {
                        $different++
                        Write-Warning "[$($current.Name)] $relPath отличается по содержимому от $($previous.Name) — это разные версии, файл оставлен как есть."
                        continue
                    }

                    # Сначала создаём ссылку рядом и только потом заменяем копию,
                    # чтобы в любой момент в снимке оставался файл с данными.
                    [VersionSafe.Native]::NewHardLink($tmp, $prevFile.FullName)
                    $file.Attributes = [System.IO.FileAttributes]::Normal
                    [System.IO.File]::Delete($file.FullName)
                    [System.IO.File]::Move($tmp, $file.FullName)
                    $fixed++
                } catch {
                    $failed++
                    Write-Warning "Не удалось восстановить ссылку для [$($current.Name)] ${relPath}: $(Get-ErrorText $_)"
                    if ([System.IO.File]::Exists($tmp)) {
                        try {
                            if ([System.IO.File]::Exists($file.FullName)) { [System.IO.File]::Delete($tmp) }
                            else { [System.IO.File]::Move($tmp, $file.FullName) }
                        } catch { }
                    }
                }
            }
        }

        if ($found -eq 0) {
            Write-Information "Проверка ссылок: разорванных ссылок не найдено."
        } elseif ($Repair) {
            Write-Information "Разорванных ссылок: $found. Восстановлено: $fixed, разные версии (не тронуты): $different, ошибок: $failed."
        } else {
            Write-Warning "Найдено разорванных ссылок: $found. Для восстановления запустите с -RepairBrokenLinks."
        }
    }

    #endregion Вспомогательные функции
}

process {
    # Скрипт сохраняет одну папку за запуск; лишние объекты конвейера — ошибка.
    $script:SourceCount++
}

end {
    #region Основная логика

    $ErrorActionPreference = "Stop"

    # Без этого Write-Information ничего не выводит и не попадает в лог.
    if (-not $PSBoundParameters.ContainsKey('InformationAction')) {
        $InformationPreference = 'Continue'
    }

    if ($script:SourceCount -gt 1) {
        throw "За один запуск сохраняется только одна папка (передано: $($script:SourceCount)). Запускайте скрипт отдельно для каждого источника, с отдельной папкой назначения."
    }

    $isDryRun = [bool]$WhatIfPreference
    if ($RepairBrokenLinks) { $VerifyLinks = $true }

    # Нормализуем пути и проверяем, что они не пересекаются
    $resolvedSource  = Get-NormalizedPath (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    $DestinationPath = Get-NormalizedPath ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath))

    if (Test-IsSameOrUnder -Path $DestinationPath -Parent $resolvedSource) {
        throw "Папка назначения совпадает с источником или находится внутри него — это приведёт к рекурсии: $DestinationPath"
    }
    if (Test-IsSameOrUnder -Path $resolvedSource -Parent $DestinationPath) {
        throw "Источник находится внутри папки назначения: $resolvedSource"
    }

    # Проверяем предыдущие снимки ДО создания папки назначения, чтобы при
    # неподключённом диске не оставлять пустую папку на чужом томе.
    $destExists    = Test-Path -LiteralPath $DestinationPath -PathType Container
    $existingSnaps = @(if ($destExists) { Get-SnapshotList -Root $DestinationPath })

    if ($FailIfOldBackupsAreMissing -and $existingSnaps.Count -eq 0) {
        throw "В $DestinationPath нет ни одного готового снимка. Работа прервана (-FailIfOldBackupsAreMissing): возможно, диск назначения не подключён."
    }

    if (-not $destExists) {
        if ($isDryRun) {
            Write-Information "[WhatIf] Будет создана папка назначения: $DestinationPath"
        } else {
            Write-Verbose "Создаю папку назначения: $DestinationPath"
            [void][System.IO.Directory]::CreateDirectory($DestinationPath)
            $destExists = $true
        }
    }

    # Защита от одновременного запуска в одну папку назначения
    $lock     = $null
    $lockPath = Join-Path -Path $DestinationPath -ChildPath 'VersionSafe.lock'
    if (-not $isDryRun) {
        try {
            $lock = [System.IO.File]::Open($lockPath, 'OpenOrCreate', 'ReadWrite', 'None')
        } catch {
            throw "VersionSafe уже работает с $DestinationPath (файл $lockPath занят другим процессом)."
        }
    }

    $transcriptOn = $false
    $exitCode     = 0

    try {
        $now      = Get-Date
        $snapName = "{0:yyyy-MM-ddTHH-mm-ss}" -f $now
        $snapPath = Join-Path -Path $DestinationPath -ChildPath $snapName
        $workPath = Join-Path -Path $DestinationPath -ChildPath "$snapName.inProgress"
        $logPath  = Join-Path -Path $DestinationPath -ChildPath "VersionSafe-$snapName.log"

        # Лог пишется в корень назначения, а не в снимок: в снимке он мог бы
        # затереть одноимённый файл из источника.
        if (-not $NoLogging -and -not $isDryRun) {
            try {
                Start-Transcript -LiteralPath $logPath -WhatIf:$false -Confirm:$false -ErrorAction Stop | Out-Null
                $transcriptOn = $true
            } catch {
                Write-Warning "Не удалось запустить лог: $(Get-ErrorText $_)"
            }
        }

        # Проверяем поддержку жёстких ссылок пробной ссылкой в самом назначении
        if (-not $NoHardLinks) {
            if ($destExists) {
                if (-not (Test-HardLinkSupport -Path $DestinationPath)) {
                    Write-Warning "Папка назначения не поддерживает жёсткие ссылки. Переключаюсь в режим полного копирования."
                    $NoHardLinks = $true
                }
            } else {
                Write-Verbose "Папки назначения ещё нет — поддержку жёстких ссылок проверить нельзя."
            }
        }

        Write-Information "=========================================="
        Write-Information "VersionSafe: старт в $now$(if ($isDryRun) { ' (тестовый прогон -WhatIf)' })"
        Write-Information "  Источник        = $resolvedSource"
        Write-Information "  Снимок          = $snapPath"
        Write-Information "  Без ссылок      = $NoHardLinks"
        if ($RotationDisabled) {
            Write-Information "  GFS             = отключена"
        } else {
            Write-Information "  GFS             = дневных:$DailyCount недельных:$WeeklyCount месячных:$MonthlyCount годовых:$YearlyCount"
        }
        Write-Information "  Логи (дней)     = $LogRetentionDays"
        Write-Information "  Проверка ссылок = $VerifyLinks (старше $VerifyLinksOlderThanDays дней, восстановление: $RepairBrokenLinks)"
        Write-Information "=========================================="

        # Убираем незавершённые снимки прерванных запусков. Параллельный запуск
        # исключён блокировкой, поэтому чужую работу мы не удалим.
        if ($destExists) {
            $stale = @(Get-ChildItem -LiteralPath $DestinationPath -Directory -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -CMatch '^\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}\.inProgress$' })
            foreach ($s in $stale) {
                if ($PSCmdlet.ShouldProcess($s.FullName, "Удалить незавершённый снимок")) {
                    try {
                        Remove-Item -LiteralPath $s.FullName -Recurse -Force -Confirm:$false -ErrorAction Stop
                        Write-Information "Удалён незавершённый снимок: $($s.Name)"
                    } catch {
                        Write-Warning "Не удалось удалить незавершённый снимок $($s.Name): $(Get-ErrorText $_)"
                    }
                }
            }
        }

        # Проверка разорванных ссылок (опционально)
        if ($VerifyLinks) {
            if ($NoHardLinks) {
                Write-Warning "Проверка ссылок пропущена: в режиме без жёстких ссылок проверять нечего."
            } else {
                Write-Information "Проверяю разорванные ссылки..."
                Test-SnapshotLinks -Root $DestinationPath -OlderThanDays $VerifyLinksOlderThanDays -Repair:$RepairBrokenLinks
            }
        }

        # Режим снимка
        $previousPath = $null
        if ($NoHardLinks) {
            Write-Verbose "Полное копирование без ссылок."
        } elseif ($existingSnaps.Count -gt 0) {
            $previousPath = $existingSnaps[0].FullName
            Write-Verbose "Предыдущий снимок: $($existingSnaps[0].Name). Инкрементальный режим."
        } else {
            Write-Verbose "Предыдущих снимков нет. Полное копирование."
        }

        # Создание снимка
        if (-not $isDryRun) {
            if (Test-Path -LiteralPath $snapPath) {
                throw "Снимок $snapName уже существует (повторный запуск в ту же секунду?)."
            }
            [void][System.IO.Directory]::CreateDirectory($workPath)
        }

        try {
            $stats = Invoke-SnapshotCopy -SourceRoot $resolvedSource -WorkPath $workPath `
                -PreviousPath $previousPath -DryRun:$isDryRun

            if (-not $isDryRun) {
                [System.IO.Directory]::Move($workPath, $snapPath)
            }
        } catch {
            Write-Warning "Сбой при создании снимка: $(Get-ErrorText $_)"
            if (-not $isDryRun -and (Test-Path -LiteralPath $workPath)) {
                Write-Information "Удаляю незавершённый снимок $workPath"
                Remove-Item -LiteralPath $workPath -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
            }
            throw
        }

        if ($isDryRun) {
            Write-Information "[WhatIf] Снимок $snapName не создавался. План:"
        } else {
            Write-Information "Снимок $snapName завершён в $(Get-Date)."
        }

        if (-not $NoStatistics) { Write-SnapshotStatistics -Stats $stats }

        if ($stats.Errors -gt 0) {
            $exitCode = 2
            Write-Warning "Снимок неполный: не удалось обработать объектов: $($stats.Errors). Подробности — в предупреждениях выше."
        }

        # Ротация — только после того, как новый снимок готов
        if ($RotationDisabled) {
            Write-Verbose "Ротация отключена параметром -RotationDisabled."
        } else {
            $cleanupArgs = @{
                Root    = $DestinationPath
                Daily   = $DailyCount
                Weekly  = $WeeklyCount
                Monthly = $MonthlyCount
                Yearly  = $YearlyCount
            }
            if ($isDryRun) { $cleanupArgs.PlannedSnapshot = $snapName }
            Invoke-SnapshotCleanup @cleanupArgs
        }

        Invoke-LogCleanup -Root $DestinationPath -MaxAgeDays $LogRetentionDays -CurrentLog $logPath
    }
    finally {
        if ($transcriptOn) {
            try { Stop-Transcript | Out-Null } catch { }
        }
        if ($lock) {
            $lock.Dispose()
            try { [System.IO.File]::Delete($lockPath) } catch { }
        }
    }

    if ($exitCode -ne 0) { exit $exitCode }

    #endregion Основная логика
}
