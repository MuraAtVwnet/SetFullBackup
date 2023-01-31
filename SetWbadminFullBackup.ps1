<#
.SYNOPSIS
wbadmin.exe を使ってフルバックアップをスケジュールします
即時バックアップも可能です

.DESCRIPTION
バックアップ先と開始時刻を指定してフルバックアップをを設定します
Windows Server / Windows Client OS の両方に対応しています
ローカルディスク/専用ディスク(Windows Server 専用)/リモート共有が指定できます

.EXAMPLE
12:00 に E ドライブにフルバックアップをする(Server / Client OS両方に対応)
SetWbadminFullBackup.ps1 -BackupTerget E: -BackupTime 12:00

.EXAMPLE
12:00 に \\FileServer\Backupにフルバックアップをする(Server / Client OS両方に対応)
SetWbadminFullBackup.ps1 -BackupTerget \\FileServer\Backup -BackupTime 12:00

.EXAMPLE
12:00 にバックアップ専用ディスク(ディスク番号 7)にフルバックアップをする(Server OS専用)
SetWbadminFullBackup.ps1 -BackupTerget 7 -BackupTime 12:00

ディスク番号は、wbadmin get disks で得られる「ディスク番号」を指定

.EXAMPLE
バックアップ停止(時刻に 99:99 を指定する)
SetWbadminFullBackup.ps1 -BackupTerget 7 -BackupTime 99:99

.EXAMPLE
E ドライブに即時フルバックアップをする(Server / Client OS両方に対応)
SetWbadminFullBackup.ps1 -BackupTerget E: -Now

.PARAMETER BackupTerget
ドライブ指定 : E:
共有指定 : \\FileServer\ShareName
ディスク番号 : 7
専用ディスクの場合は wbadmin get disks の ディスク番号

.PARAMETER BackupTime
HH:MM
24時間制で記述
99:99 を指定するとバックアップ停止

.PARAMETER Now
即時バックアップ
時刻指定は無視されます

<CommonParameters> はサポートしていません

.LINK
http://www.vwnet.jp/Windows/PowerShell/wbadmin_FullBackup.htm
#>

#################################################################################
#
# フルバックアップ設定
# バックアップ先
#	ドライブ  E:
#	共有  \\FileServer\ShareName
#	専用ディスクの場合は wbadmin get disks の ディスク番号
#	あるいは Get-Disk の Number (同じ値が得られる)
#
# ネットワークバックアップからのリストアは F10 で DOS 窓開いて IP アドレス割り当てる必要あり
#	netsh interface ipv4 show address
#	netsh interface ipv4 set address [IF NAME] static [IP Address] [Subnet Mask] [Default Gateway]
#
#################################################################################
param (
		[string]$BackupTerget,	# バックアップ先
		[string]$BackupTime,	# バックアップ開始時刻
		[switch]$Now,			# 即時バックアップ
		[switch]$WhatIf			# 動作テスト
		)

$G_ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent

$G_PSDriveName = "Backup"
$G_PSDrive = $G_PSDriveName + ":"

##########################################################################
# OS バージョン確認
##########################################################################
function CheckOSVertion(){
	$OSData = Get-WmiObject Win32_OperatingSystem
	$BuildNumber = $OSData.BuildNumber
	$strVersion = $OSData.Version
	$strVersion = $strVersion.Replace( ".$BuildNumber", "" )
	$WinVer = [decimal]$strVersion
	if( $WinVer -lt 6.2 ){
		echo "[FAIL] Windows 8 / Windows Server 2012 以降のサポートです"
		exit
	}
}

##########################################################################
# Windows Server Backup インストール と Server OS か否かの確認
##########################################################################
function InstallWindowsServerBackup(){
	$OS = (Get-WmiObject Win32_OperatingSystem).Caption
	if( $OS -Match "Server" ){
		if( (Get-WindowsFeature -Name Windows-Server-Backup).Installed -eq $false ){
			Add-WindowsFeature -Name Windows-Server-Backup	-IncludeAllSubFeature -IncludeManagementTools
		}
		return $true
	}
	else{
		return $false
	}
}


##########################################################################
# 開始時刻の整形
##########################################################################
function SetStartTime( $BackupTime ){
	# Disable Backup
	if( $BackupTime -eq "99:99" ){
		WBADMIN DISABLE BACKUP -quiet
		echo "[INFO] バックアップスケジュールを停止しました"
		exit
	}

	# 99:99 に桁数そろえる
	$hh = ($BackupTime.Split(":"))[0]
	$mm = ($BackupTime.Split(":"))[1]
	if( $hh.Length -eq 1 ){
		$hh = "0" + $hh
	}
	if( $mm.Length -eq 1 ){
		$mm = "0" + $mm
	}

	$BackupTime = $hh + ":" + $mm

	# 中身確認
	if( $BackupTime -notmatch "^[0-9]{2,2}:[0-9]{2,2}$" ){
		echo "[FAIL] バックアップ時刻が $BackupTerget が正しくない : $BackupTime"
		exit
	}
	else{
		if( ([int]$hh -gt 23) -or ([int]$mm -gt 59) ){
			echo "[FAIL] バックアップ時刻が $BackupTerget が正しくない : $BackupTime"
			exit
		}
		else{
			return $BackupTime
		}
	}
}

##########################################################################
# 専用ディスクのチェック
##########################################################################
function CheckPrivateDisk($BackupTerget){
	# バックアップ先のディスク
	$TergetDisk = Get-Disk -Number $BackupTerget -ErrorAction SilentlyContinue
	if( $TergetDisk -eq $null ){
		echo "[FAIL] Disk Number $BackupTerget が存在しません"
		exit
	}
	# オフラインならオンラインにする
	if( $TergetDisk.IsOffline ){
		Set-Disk -Number $BackupTerget -IsOffline:$false
	}

	# 初期化されていなかったら GPT にする
	if( $TergetDisk.PartitionStyle -eq "RAW" ){
		Initialize-Disk -Number $BackupTerget -PartitionStyle GPT
	}

	# wbadmin get disks から必要データーをオブジェクトにする
	$WSB_Datas = New-Object System.Collections.ArrayList

	$Lines = wbadmin get disks
	$DataOUT = $true
	foreach($Line in $Lines){
		if( $DataOUT ){
			$WSB_Data = New-Object PSObject | Select-Object DiskName, Number, GUID, ExcludeDriveLetter
			$DataOUT = $false
		}
		if( ($Line -match "^ディスク名:") -or ($Line -match "^Disk name:") ){
			$LineData = ($Line.Split(":"))[1]
			$WSB_Data.DiskName = $LineData.Trim()
		}

		if( ($Line -match "^ディスク番号:") -or ($Line -match "^Disk number:") ){
			$LineData = ($Line.Split(":"))[1]
			$WSB_Data.Number = $LineData.Trim()
		}

		if( ($Line -match "^ディスク ID:") -or ($Line -match "^Disk identifier:") ){
			$LineData = ($Line.Split(":"))[1]
			$WSB_Data.GUID = $LineData.Trim()
			[void]$WSB_Datas.Add($WSB_Data)
			$DataOUT = $true
		}
	}

	$TergetDiskData = $WSB_Datas | ? {$_.Number -eq $BackupTerget}
	if( $TergetDiskData -eq $null ){
		echo "[FAIL] wbadmin get disks が想定外の値を返しました"
		exit
	}

	$DiskNumber = $TergetDiskData.Number
	$Partitions = @(Get-Partition -DiskNumber $DiskNumber)
	# パーティションが存在する場合
	if($Partitions.Length -ge 2){
		$AccessableDrives = @($Partitions | ?{ $_.DriveLetter -match "[a-zA-Z]"})
		if( $AccessableDrives.Length -ne 0 ){
			$DriveLetters = @($AccessableDrives.DriveLetter)

			$Hostname = hostname
			# 7文字で切る
			if( $Hostname.Length -gt 7 ){
				$Hostname = $Hostname.Substring(0,7)
			}

			foreach( $DriveLetter in $DriveLetters ){
				$VolumeLabel = (Get-Volume -DriveLetter $DriveLetter).FileSystemLabel
				# ボリュームラベルが hostname YYYY_MM_DD HH:MM Disk_99 なら OK
				if( $VolumeLabel -match "^$Hostname [0-9]{4,4}_[0-9]{2,2}_[0-9]{2,2} [0-9]{2,2}:[0-9]{2,2} Disk_[0-9]+" ){
					$TergetDiskData.ExcludeDriveLetter = $DriveLetter + ":"
				}
				else{
					echo "[FAIL] データボリュームが指定した専用ディスクに含まれまれています"
					echo "[FAIL] 専用ディスクにする場合はパーティションを削除してください"
					exit
				}
			}
		}
		else{
			echo "[FAIL] 専用ディスクバックアップが設定されているか、用途不明のパーティションが存在します"
			echo "[FAIL] バックアップを解除するかパーティション状態を確認してください"
			exit
		}
	}
	return $TergetDiskData
}

##########################################################################
# バックアップ先のドライブレター取得
##########################################################################
function GetTergetDriveLeter($TergetDrive){
	if( -not ( Test-Path ($TergetDrive + "\" ))){
		echo "[FAIL] ドライブ $TergetDrive が存在しません"
		exit
	}
	$ExcludeDrive = $TergetDrive
	return $ExcludeDrive
}

##########################################################################
# パスワード取り出し
##########################################################################
function GetPassword($SecureString){
	$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
	$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
	return $Password
}

################################################
# ドメインユーザーが存在するか
################################################
function IsADUserAccunt( $DomainName, $DomainUser ){
	$hostname = hostname
	$ADUser = [ADSI]("WinNT://$DomainName/$DomainUser")
	if( $ADUser.ADsPath -ne $null ){
		return $true
	}
	else{
		return $false
	}
}

################################################
# ドメインユーザーがメンバーになっているか
################################################
function IsMemberDomainAccunt( $DomainName, $DomainUser, $LocalGroup ){
	$HostName = hostname
	$ADUser = [ADSI]("WinNT://$DomainName/$DomainUser")
	$LocalGroup = [ADSI]("WinNT://$HostName/$LocalGroup")
	return $LocalGroup.IsMember($ADUser.ADsPath)
}

###########################################################
# ドメインユーザー/ドメイングループのローカルグループ参加
###########################################################
function JoinADUser2Group( $DomainName, $DomainUser, $LocalGroup ){
	$HostName = hostname
	$ADUser = [ADSI]("WinNT://$DomainName/$DomainUser")
	$LocalGroup = [ADSI]("WinNT://$HostName/$LocalGroup")
	$LocalGroup.Add($ADUser.ADsPath)
}

################################################
# ローカルユーザーが存在するか
################################################
function IsLocalUserAccunt( $UserID ){
	$hostname = hostname
	[ADSI]$Computer = "WinNT://$hostname,computer"
	$Users = $Computer.psbase.children | ? {$_.psBase.schemaClassName -eq "User"} | Select-Object -expand Name
	return ($Users -contains $UserID)
}

################################################
# ローカルユーザーがメンバーになっているか
################################################
function IsMemberLocalAccunt( $UserID, $GroupName ){
	$hostname = hostname
	[ADSI]$Computer = "WinNT://$hostname,computer"
	$Group = $Computer.GetObject("group", $GroupName)
	$User = $Computer.GetObject("user", $UserID)
	return $Group.IsMember($User.ADsPath)
}

################################################
# グループへ参加
################################################
function JoinGroup( $UserID, $JoinGroup ){
	$hostname = hostname
	[ADSI]$Computer = "WinNT://$hostname,computer"
	$Group = $Computer.GetObject("group", $JoinGroup)
	$Group.Add("WinNT://$hostname/$UserID")
}

##########################################################################
# RW 権限があるか確認
##########################################################################
function CheckACL( $BackupTerget ){
	$TestFile = Join-Path $G_PSDrive "Test.txt"
	if( Test-Path $TestFile ){
		del $TestFile
	}
	if( Test-Path $TestFile ){
		echo "[FAIL] $BackupTerget に RW 権限がありません"
		exit
	}
	Write-Output "RW Test" | Out-File -FilePath $TestFile
	if( -not (Test-Path $TestFile )){
		echo "[FAIL] $BackupTerget に RW 権限がありません"
		exit
	}
	else{
		del $TestFile
	}
}

##########################################################################
# バックアップボリューム
##########################################################################
function GetBackupVolumes($ExcludeDrive){
	$BackupVolumes = ""
	$Volumes = Get-WmiObject win32_volume -Filter "DriveType = 3"
	foreach( $Volume in $Volumes ){
		$DriveLetter = $Volume.DriveLetter
		# ドライブレターがある
		if( $DriveLetter -match "[a-zA-Z]" ){
			# ドライブレターがある
			if( $ExcludeDrive -ne $null ){
				if( $ExcludeDrive -ne $DriveLetter ){
					$BackupVolumes += $DriveLetter + ","
				}
			}
			# バックアップ先がローカルフォルダーでなければ全てバックアップ対象にする
			else{
				$BackupVolumes += $DriveLetter + ","
			}
		}
	}
	# 末尾のカンマ消す
	$BackupVolumes = $BackupVolumes -replace ",$",""
	return $BackupVolumes
}

##########################################################################
# VM 取ってくる
##########################################################################
function GetVMs(){
	$VMNames = ""
	$VMs = @(Get-VM)
	foreach( $VM in $VMs ){
		$VMNames += "`"" + $VM.Name + "`","
	}

	# Host Component を追加
	if( $VMNames -ne "" ){
		$VMNames += "`"Host Component`""
	}
	else{
		$VMNames = $null
	}
	return $VMNames
}

##########################################################################
# MAIN
##########################################################################

if( ($BackupTerget -eq $null) -or ($BackupTerget -eq "") ){
	$ScriptFullName = $MyInvocation.MyCommand.Path
	Get-Help $ScriptFullName
	exit
}

echo "[INFO] バックアップ登録開始"

if (-not(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))) {
	echo "[FAIL] 実行には管理権限が必要"
	exit
}

# OS バージョン確認
CheckOSVertion

# Windows Server Backup インストール と Server OS か否かの確認
$IsServerOS = InstallWindowsServerBackup

### -schedule:

# 即時バックアップの時はスキップ
if( $Now -eq $false ){
	$BackupTime = SetStartTime $BackupTime
}

### バックアップ先に専用ディスクを指定
if( $BackupTerget -match "^[0-9]+$" ){
	if( $IsServerOS ){
		$TergetDiskData = CheckPrivateDisk $BackupTerget
		# バックアップ停止して一時的にドライブレターが割り当たっていたらバックアップ対象から外す
		$ExcludeDrive = $TergetDiskData.ExcludeDriveLetter
		$GUID = $TergetDiskData.GUID
		$BackupTerget = $GUID
	}
	else{
		echo "[FAIL] 専用ディスクは Server OS 機能です"
		exit
	}
}
else{
	$TergetDrive = Split-Path $BackupTerget -Qualifier -ErrorAction SilentlyContinue
	### バックアップ先にローカルドライブが指定されている場合
	if( $TergetDrive -ne $null ){
		# バックアップ先のドライブをバックアップ対象外にする
		$ExcludeDrive = GetTergetDriveLeter $TergetDrive
		# バックアップ先
		$BackupTerget = $ExcludeDrive
	}
	### バックアップ先に SMB 共有指定 が指定されている場合
	else{

		# バックアップアカウント
		$Credential = Get-Credential -Message "バック先へのアクセスアカウント(domain\user or user)"

		# アクセスアカウントを登録
		$ID = $Credential.UserName
		$GroupName = "Backup Operators"

		# ドメインアカウント
		if( $ID -match "\\" ){
			$DomainName = ($ID.Split("\"))[0]
			$BackupUserID = ($ID.Split("\"))[1]

			$Statsus = IsADUserAccunt $DomainName $BackupUserID
			if( $Statsus -eq $false ){
				echo "[FAIL] Domain:$DomainName ID:$BackupUserID が存在しません"
				exit
			}

			# アカウントを Backup Operators に入れる
			$Statsus = IsMemberDomainAccunt $DomainName $BackupUserID $GroupName
			if( $Statsus -eq $false ){
				JoinADUser2Group $DomainName $BackupUserID $GroupName
			}
		}
		# ローカルアカウント
		else{
			$Statsus = IsLocalUserAccunt $ID
			if( $Statsus -eq $false ){
				echo "[FAIL] ID:$ID が存在しません"
				exit
			}

			# アカウントを Backup Operators に入れる
			$Statsus = IsMemberLocalAccunt $ID $GroupName
			if( $Statsus -eq $false ){
				JoinGroup $ID $GroupName
			}
		}
		### -addtarget:
		# RW 権限があるか確認
		if( Test-Path $BackupTerget ){
			New-PSDrive -Name $G_PSDriveName -PSProvider FileSystem -Root $BackupTerget -Credential $Credential
			CheckACL $BackupTerget
			Remove-PSDrive $G_PSDriveName
		}
		else{
			echo "[FAIL] $BackupTerget にアクセスできません"
			exit
		}
	}
}

### -include:
# バックアップボリューム
$BackupVolumes = GetBackupVolumes $ExcludeDrive

### -hyperv:
# Hyper-V VM
if( $IsServerOS ){
	if((Get-WindowsFeature -Name Hyper-V).installed -eq $true){
		$VMNames = GetVMs
	}
}

# バックアップオプションを作る
if( $Now ){
	# 即時バックアップ
	if( $IsServerOS ){
		$Options = "-addtarget:$BackupTerget -include:$BackupVolumes -allCritical -systemState -vssFull -quiet"
	}
	else{
		$Options = "-addtarget:$BackupTerget -include:$BackupVolumes -allCritical -vssFull -quiet"
	}
}
else{
	# バックアップスケジュール
	if( $IsServerOS ){
		$Options = "-backupTarget:$BackupTerget -include:$BackupVolumes -allCritical -systemState -vssFull -schedule:$BackupTime -quiet"
	}
	else{
		$Options = "-backupTarget:$BackupTerget -include:$BackupVolumes -allCritical -vssFull -schedule:$BackupTime -quiet"
	}
}

# VM が入っていたら VM バックアップ
if( $VMNames -ne $null ){
	$Options += " -hyperv:$VMNames"
}

# SMB 共有がバックアップ先の場合は ID パスワードをセット
if( $Credential -ne $null ){
	$ID = $Credential.UserName
	$Password = GetPassword ($Credential.Password)
	$Options += " -user:$ID -password:$Password"
}

echo "[INFO] Backup options : $Options"

if( $Now ){
	$CommandLine = "WBADMIN START BACKUP " + $Options
}
else{
	$CommandLine = "WBADMIN ENABLE BACKUP " + $Options
}

echo "Cooand line : $CommandLine"

if( -not $WhatIf ){
	# バックアップ設定
	cmd /c $CommandLine
}

echo "[INFO] バックアップ登録完了"
