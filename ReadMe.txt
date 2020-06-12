■ これは何?
wbadmin.exe を使ってフルバックアップをスケジューリングします
即時バックアップも可能です
Windows 8 / Windows Server 2012 以降をサポートしています。

■ 使い方
.\SetWbadminFullBackup.ps1 -BackupTerget バックアップ先 [-BackupTime バックアップ開始時刻] [-Now]

■ バックアップ先
・ドライブ	E:
・共有		\\FileServer\ShareName
    専用ディスクの場合は wbadmin get disks の ディスク番号
    あるいは Get-Disk の Number (同じ値が得られる)

■ 実行例
12:00 に E ドライブにフルバックアップをする(Server / Client OS両方に対応)
.\SetWbadminFullBackup.ps1 -BackupTerget E: -BackupTime 12:00

12:00 に \\FileServer\Backupにフルバックアップをする(Server / Client OS両方に対応)
.\SetWbadminFullBackup.ps1 -BackupTerget \\FileServer\Backup -BackupTime 12:00

12:00 にバックアップ専用ディスク(ディスク番号 7)にフルバックアップをする(Server OS専用)
.\SetWbadminFullBackup.ps1 -BackupTerget 7 -BackupTime 12:00

    ディスク番号は、wbadmin get disks で得られる「ディスク番号」を指定

バックアップ停止(時刻に 99:99 を指定する)
.\SetWbadminFullBackup.ps1 -BackupTerget 7 -BackupTime 99:99

E ドライブに即時フルバックアップをする(Server / Client OS両方に対応)
.\SetWbadminFullBackup.ps1 -BackupTerget E: -Now


■ Web Site
http://www.vwnet.jp/Windows/PowerShell/wbadmin_FullBackup.htm

