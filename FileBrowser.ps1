Param( $StartPath )
#region CONSTANTS & ASSEMBLIES ##########################################################################################
	Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms
	[Windows.Forms.Application]::EnableVisualStyles()

	$APPNAME = "File Browser"
	$VERSION = '1.0.0.0'
	<#
		Updates:
		1.0.0.0 - FileBrowser separated from AdminPad.
	#>

	If ($PSCommandPath) {
		$COMPILED = $False
		$SCRIPTFILEPATH = $PSCommandPath
	} Else {
		$COMPILED = $True
		$SCRIPTFILEPATH = (Get-Process -id $PID).Path
	}
	$SCRIPTDIR = Split-Path $SCRIPTFILEPATH
#endregion CONSTANTS & ASSEMBLIES ##########################################################################################

#region COMPILE PROMPT ##########################################################################################
	<#
		Prompt to compile only if running as script.
	#>
	If ($PSCommandPath -and 1) {
		Write-Host "Script path detected:`n > $PSCommandPath"
		Write-Host "`nEnter ""C"" to compile.`nAny other input to test run."
		$CompileAsk = Read-Host -Prompt "Enter"
		Clear-Host
		If ($CompileAsk -eq 'C') {
			$CompilerModule = 'PS2EXE'

			If (!(Get-Module $CompilerModule)) { # module is not imported
				Write-Host "$CompilerModule module not imported."
				If (!(Get-Module -ListAvailable $CompilerModule)) { # module is not available
					Write-Host "$CompilerModule module not available. Searching online gallery..."
					If (Find-Module -Name $CompilerModule -ErrorAction:SilentlyContinue) {
						Write-Host "$CompilerModule found. Downloading..."
						Install-Module -Name $CompilerModule -Force -Verbose -Scope CurrentUser
					} Else {
						Write-Host "Cannot find module: $CompilerModule."
						Write-Host "Script will exit now."
						Pause
						Exit
					}
				}
				Write-Host "Importing $CompilerModule module."
				Import-Module $CompilerModule -Verbose
			}
			Write-Host "Compiling..."

			$ScriptFolder = Split-Path $PSCommandPath

			Invoke-ps2exe -inputFile $PSCommandPath -iconFile "$ScriptFolder\FileBrowser.ico" `
				-noConfigFile -noConsole -noError -noOutput `
				-product $APPNAME -version $VERSION -title $APPNAME `
				-company 'therkSoft' -copyright 'Rob Saunders'

			Read-Host "Press enter to launch application"
			Start-Process "$ScriptFolder\FileBrowser.exe"
			Exit
		} Else {
			$StartPath = $CompileAsk
		}
	}
#endregion COMPILE PROMPT ##########################################################################################

#region REGISTRY SETUP ##########################################################################################
	$REGKEY = "Registry::HKCU\Software\$APPNAME"
	If (!(Test-Path -Path $REGKEY)) { New-Item -Path $REGKEY -Force | Out-Null }
	$REGKEY = Get-Item $REGKEY
#endregion REGISTRY SETUP ##########################################################################################

#region FUNCTIONS ##########################################################################################
	Function Icon-Extract {
		Param($Path, $Index, [Switch]$LargeIcon)
		$IconExtract = Add-Type -Name IconExtract -MemberDefinition '
			[DllImport("Shell32.dll", SetLastError=true)]
			public static extern int ExtractIconEx(string lpszFile, int nIconIndex, out IntPtr phiconLarge, out IntPtr phiconSmall, int nIcons);
		' -PassThru

		#Initialize variables for reference conversion
		$IconLarge, $IconSmall = 0, 0

		#Call Win32 API Function for handles
		If ($IconExtract::ExtractIconEx($Path, $Index, [ref]$IconLarge, [ref]$IconSmall, 1)) {
			[System.Drawing.Icon]::FromHandle( $( If ($LargeIcon) { $IconLarge } Else { $IconSmall } ) )
		}
	}

	Function Icon-FromFilePath {
		Param($FilePath, [Switch]$LargeIcon)

		Add-Type -TypeDefinition '
			using System;
			using System.Drawing;
			using System.Runtime.InteropServices;

			public class Icon_FromFilePath
			{
				[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
				public struct SHFILEINFO
				{
					public IntPtr hIcon;
					public int iIcon;
					public uint dwAttributes;
					[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
					public string szDisplayName;
					[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
					public string szTypeName;
				}

				[DllImport("shell32.dll", CharSet = CharSet.Unicode)]
				public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);
				
				[DllImport("user32.dll", SetLastError=true)]
				public static extern bool DestroyIcon(IntPtr hIcon);
			}
		'

		$Flags = 0x100
		If (!$LargeIcon) { $Flags += 0x1 }  # Add SHGFI_SMALLICON flag if not requesting a large icon

		$FileInfoStruct = New-Object Icon_FromFilePath+SHFILEINFO
		$StructSize = [System.Runtime.InteropServices.Marshal]::SizeOf($FileInfoStruct)

		If (Test-Path variable:Icon_FromFilePath_Cleanup) { [void][Icon_FromFilePath]::DestroyIcon($Icon_FromFilePath_Cleanup) }
		[void][Icon_FromFilePath]::SHGetFileInfo($FilePath, 0, [ref]$FileInfoStruct, $StructSize, $Flags)
		$global:Icon_FromFilePath_Cleanup = $FileInfoStruct.hIcon

		Return [System.Drawing.Icon]::FromHandle($FileInfoStruct.hIcon)
	}

	# MsgBox shortcut
	Function MsgBox($Message, $Buttons = 0, $Icon = 0, $Default = 0) {
		# $Buttons: OK, OKCancel, AbortRetryIgnore, YesNoCancel, YesNo, RetryCancel
		# $Icon: None, Info, Warn, Error
		# $DefaultIndex: 0, 1, 2
		Return [Windows.MessageBox]::Show($Message, $APPNAME, $Buttons, $Icon, $Default).ToString()
	}

	# Properties Dialog
	Function PropertiesDialog($FilePath) {
		If ($global:ShellApplication -eq $null) { $global:ShellApplication = New-Object -com Shell.Application }

		If (Test-Path $FilePath) {
			If ((Get-Item $FilePath).GetType().Name -eq 'DirectoryInfo') {
				$Folder = $ShellApplication.NameSpace($FilePath)
				$Folder.Self.InvokeVerb("Properties")
			} Else {
				$Folder = Split-Path $FilePath
				$File = Split-Path $FilePath -Leaf
				$Folder = $ShellApplication.NameSpace($Folder)
				$File = $Folder.ParseName($File)
				$File.InvokeVerb("Properties")
			}
		}
	}

	# ToolTip object
	$ToolTip = @{}
	$ToolTip | Add-Member -Type ScriptMethod -Name Popup -Value {
		param ($Text, $Control, $X, $Y, $Time)
		$tt_Temp = New-Object Windows.Forms.ToolTip -Property @{ IsBalloon = $True }
		$tt_Temp.SetToolTip($Control, ' ')
		$tt_Temp.Show($Text, $Control, $X, $Y, $Time)
		$tt_Temp.add_popup({ $this.dispose() })
	}

	$ToolTip | Add-Member -Type ScriptMethod -Name SetTip -Value {
		param ($Control, $Text)
		$tt_Temp = New-Object Windows.Forms.ToolTip -Property @{ IsBalloon = $True }
		$tt_Temp.SetToolTip($Control, $Text)
	}

	Function UpdateFileList($Path) {
		$fm_Browser.Text = $APPNAME + ' - Loading, please wait...'
		$sbl_Status.Text = ""
		$sbl_Status.ResetForeColor()

		Function GetFileTypeName($Extension) {
			If ($global:ExtensionNameCache -eq $null) { $global:ExtensionNameCache = @{'' = 'System File'; '.' = 'System File'} }

			If ($global:ExtensionNameCache[$Extension]) {
				Return $global:ExtensionNameCache[$Extension]
			} Else {
				$global:ExtensionNameCache[$Extension] = $Extension.Substring(1).ToUpper() + " File"

				If ($Assoc = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$Extension" -EA SilentlyContinue).'(default)') {
					If ($TypeName = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$Assoc" -EA SilentlyContinue).'(default)') {
						$global:ExtensionNameCache[$Extension] = $TypeName
					}
				}
				Return $global:ExtensionNameCache[$Extension]
			}
		}

		$Path = $Path -replace '"',''
		If (!$Path) { $Path = $lv_FileBrowser.Tag.CurrentPath.FullName }
		If (!$Path.EndsWith('\')) { $Path = [string]$Path + '\' }

		$GCI_Path = Get-ChildItem $Path -Force -EA SilentlyContinue -ErrorVariable GCI_Error
		If ($GCI_Error) {
			$sbl_Status.Text = "$GCI_Error (GCI_Error)"
			$sbl_Status.ForeColor = 'red'
			$fm_Browser.Text = $APPNAME
			Return
		}

		$GI_Path = Get-Item $Path -Force -EA SilentlyContinue -ErrorVariable GI_Error
		If ($GI_Error) {
			$sbl_Status.Text = "$GI_Error (GI_Error)"
			$sbl_Status.ForeColor = 'red'
			$fm_Browser.Text = $APPNAME
			Return
		}
		If ($GI_Path.GetType().Name -ne 'DirectoryInfo') {
			$sbl_Status.Text = "Path is not a valid folder"
			$sbl_Status.ForeColor = 'red'
			$fm_Browser.Text = $APPNAME
			Return
		}

		Set-Location $Path

		$lv_FileBrowser.Items.Clear()
		$global:lviArray_FileBrowser = @()
		$il_FileBrowser = New-Object Windows.Forms.ImageList -Property @{
			ImageSize = '16,16'
			ColorDepth = 'Depth32Bit'
		}
		$lv_FileBrowser.SmallImageList = $il_FileBrowser
		$il_FileBrowser.Images.Add('*FOLDER', (Icon-Extract shell32.dll 3))

		# If there is a parent folder
		If ($GI_Path.Parent) {
			$Item = New-Object Windows.Forms.ListViewItem('<Parent Folder>', '*FOLDER')
			$Item.Tag = $GI_Path.Parent
			[void]$Item.SubItems.Add('Folder')
			[void]$Item.SubItems.Add('')
			[void]$Item.SubItems.Add('')
			$global:lviArray_FileBrowser += ,$Item
		}

		$lv_FileBrowser.Tag.CurrentPath = $GI_Path
		$tb_Address.Text = $GI_Path.FullName
		$FolderBrowserDialog.SelectedPath = $GI_Path.FullName

		Foreach ($File in (Get-ChildItem $GI_Path -Force)) {
			If ($File.PSIsContainer) {
				$Item = New-Object Windows.Forms.ListViewItem($File.Name, '*FOLDER')
				$Item.Tag = $File
				[void]$Item.SubItems.Add('Folder')
				[void]$Item.SubItems.Add($File.LastWriteTime.toString('g'))
				[void]$Item.SubItems.Add('')
			} Else {
				# $il_FileBrowser.Images.Add($File.Name, [Drawing.Icon]::ExtractAssociatedIcon($File.FullName))
				$il_FileBrowser.Images.Add("ico*$($File.Name)", (Icon-FromFilePath $File.FullName))

				$Item = New-Object Windows.Forms.ListViewItem($File.Name, "ico*$($File.Name)")
				$Item.Tag = $File
				[void]$Item.SubItems.Add( (GetFileTypeName $File.Extension) )
				[void]$Item.SubItems.Add($File.LastWriteTime.toString('g'))
				[void]$Item.SubItems.Add((Format-FileSize $File.Length))
			}
			$global:lviArray_FileBrowser += ,$Item
		}

		$lv_FileBrowser.Items.AddRange($global:lviArray_FileBrowser)
		$fm_Browser.Text = $APPNAME
	}

	Function Format-FileSize($Bytes) {
		$SizeOut = [Math]::Max(1, $Bytes / 1024)
		$Units = "KB", "MB", "GB", "TB", "PB", "EB"
		$idx = 0

		While ($SizeOut -gt 1024 -and $idx -le $Units.length) {
			$SizeOut /= 1024
			$idx++
		}
		Return [String]([Math]::Round($SizeOut)) + ' ' + $Units[$idx]
	}

#endregion FUNCTIONS ##########################################################################################

#region GLOBAL OBJECTS ##########################################################################################
	$FolderBrowserDialog = New-Object Windows.Forms.FolderBrowserDialog -Property @{
		Description = 'Select folder'
	}

#endregion GLOBAL OBJECTS ##########################################################################################

#region SETUP BROWSER GUI ##########################################################################################
	$fm_Browser = New-Object Windows.Forms.Form -Property @{
		Text = "$APPNAME"
		StartPosition = 'Manual'
		ClientSize = '500,315'
		Icon = (Icon-Extract imageres.dll 205)
		SizeGripStyle = 'Show'

		# Setup key binds
		KeyPreview = $True
		add_Closing = {
			$AppContext.ExitThread()

			Stop-Process $PID # Kill the process
		}

		add_Load = {
			If (!$StartPath) { $StartPath = (Get-Location).Path }

			UpdateFileList -Path $StartPath

			$this.MinimumSize = $this.Size

			# Check registry for stored window size, restore if valid
			If (($WinSize = $REGKEY.GetValue('BrowserSize')) -and $WinSize -match '^\d+,\d+$') {
				$this.Size = $WinSize
			}
			If (($WinPos = $REGKEY.GetValue('BrowserPos')) -and $WinSize -match '^\d+,\d+$') {
				$this.Location = $WinPos
			}

			# Set the resize handler inside the load event otherwise it activates on window creation, resetting the saved window size
			$this.add_Resize({
				If ($this.WindowState -eq 'Normal') {
					Set-ItemProperty -Path $REGKEY.PSPath -Name BrowserSize -Value "$($this.Width),$($this.Height)"
					Set-ItemProperty -Path $REGKEY.PSPath -Name BrowserPos -Value "$($this.Left),$($this.Top)"
				}
			})
		}
	}

	#region CONTROLS FOR FILE BROWSER ##########################################################################################
		$fm_Browser.Controls.Add((New-Object Windows.Forms.Label -Property @{
			Bounds = "10,5,50,20"
			Text = 'A&ddress:'
			TextAlign = 'MiddleLeft'
		}))
		$fm_Browser.Controls.Add(($tb_Address = New-Object Windows.Forms.ComboBox -Property @{
			Bounds = "60,5,355,20"
			Anchor = 'Top,Left,Right'
			add_KeyDown = {
				If ($_.KeyCode -eq 'Return') {
					$_.SuppressKeyPress = !$this.DroppedDown # If Enter is suppressed when dropdown is open, it will not close
					UpdateFileList -Path $this.Text
				}
			}
			add_SelectionChangeCommitted = { UpdateFileList -Path $this.SelectedItem }
		}))
		$tb_Address.Items.AddRange((Get-PSDrive).Root -match ':\\')

		$fm_Browser.Controls.Add(($bt_GoFolder = New-Object Windows.Forms.Button -Property @{
			Bounds = "415,3,25,24"
			Anchor = 'Top,Right'
			BackgroundImage = (Icon-Extract shell32.dll 299)
			BackgroundImageLayout = 'Center'
			ImageAlign = 'MiddleCenter'
			add_Click = {
				UpdateFileList -Path $tb_Address.Text
			}
		}))
		$ToolTip.SetTip($bt_GoFolder, "Go / Refresh")
		$fm_Browser.Controls.Add(($bt_UpFolder = New-Object Windows.Forms.Button -Property @{
			Bounds = "440,3,25,24"
			Anchor = 'Top,Right'
			BackgroundImage = (Icon-Extract shell32.dll 307) #307
			BackgroundImageLayout = 'Center'
			ImageAlign = 'MiddleCenter'
			add_Click = ($UpFolder_Func = {
				If ($lv_FileBrowser.Tag.CurrentPath.Parent) {
					UpdateFileList -Path $lv_FileBrowser.Tag.CurrentPath.Parent.FullName
				}
			})
		}))
		$ToolTip.SetTip($bt_UpFolder, "Go up to parent folder")
		$fm_Browser.Controls.Add(($bt_BrowseFolder = New-Object Windows.Forms.Button -Property @{
			Bounds = "465,3,25,24"
			Anchor = 'Top,Right'
			BackgroundImage = (Icon-Extract imageres.dll 205)
			BackgroundImageLayout = 'Center'
			ImageAlign = 'MiddleCenter'
			add_Click = {
				If ($FolderBrowserDialog.ShowDialog() -eq 'OK') {
					UpdateFileList -Path $FolderBrowserDialog.SelectedPath
				}
			}
		}))
		$ToolTip.SetTip($bt_BrowseFolder, "Select folder")

	#endregion CONTROLS FOR FILE BROWSER ##########################################################################################

	#region MENU ITEMS FOR FILE BROWSER ##########################################################################################
		$cm_FileBrowser = New-Object Windows.Forms.ContextMenu -Property @{
			add_Popup = {
				If ($lv_FileBrowser.SelectedIndices[0] -ne $null) {
					$cm_FileBrowser.MenuItems | Foreach-Object { $_.Enabled = $True }

					If (Test-Path $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag -PathType Leaf) {
						$mi_OpenNewWin.Enabled = $False
					}
				} Else {
					$cm_FileBrowser.MenuItems | Foreach-Object { $_.Enabled = $False }
				}
			}
		}
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{
			DefaultItem = $True
			Text = '&Open'
			add_Click = { Invoke-Item $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName }
		}))
		[void]$cm_FileBrowser.MenuItems.Add(($mi_OpenNewWin = New-Object Windows.Forms.MenuItem -Property @{
			Text = 'Open in &new browser'
			add_Click = {
				Start-Process (Get-Process -ID $PID).Path -ArgumentList ('"'+$lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName+'"')
			}
		}))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{
			Text = 'Open wit&h...'
			add_Click = {
				OpenWith.exe $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName
			}
		}))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{ Text = '-' } ))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{
			Text = '&Copy to...'
			add_Click = {
				$SelItem = $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName

				$CopyToPath = New-Object Windows.Forms.SaveFileDialog -Property @{
					Title = 'Copy item to...'
					InitialDirectory = (Split-Path $SelItem)
					FileName = (Split-Path $SelItem -Leaf)
					Filter = 'All files (*.*)|*.*'
				}

				If ($CopyToPath.ShowDialog() -eq 'OK') {
					$Success = Copy-Item $SelItem $CopyToPath.FileName -PassThru
					If (!$Success) {
						MsgBox "Unable to copy item:`n`n$($error[0].toString())" 0 'Error'
					}
				}
			}
		}))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{
			Text = 'Move &to...'
			add_Click = {
				$SelItem = $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName

				$MoveToPath = New-Object Windows.Forms.SaveFileDialog -Property @{
					Title = 'Move item to...'
					InitialDirectory = (Split-Path $SelItem)
					FileName = (Split-Path $SelItem -Leaf)
					Filter = 'All files (*.*)|*.*'
				}

				If ($MoveToPath.ShowDialog() -eq 'OK') {
					$Success = Move-Item $SelItem $MoveToPath.FileName -PassThru
					If (!$Success) {
						MsgBox "Unable to move item:`n`n$($error[0].toString())" 0 'Error'
					}
					UpdateFileList
				}
			}
		}))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{
			Text = 'Copy p&ath'
			add_Click = { $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName | Set-Clipboard }
		}))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{ Text = '-' } ))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{
			Text = '&Delete'
			add_Click = {
				$SelItem = $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName
				If ((MsgBox "Are you sure you want to delete this file?`n$($SelItem.Tag.Name)" "YesNo" "Warning" "No") -eq 'Yes') {
					Remove-Item $SelItem
					UpdateFileList
				}
			}
		}))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{
			Text = 'Rena&me'
			add_Click = {
				$fm_Rename.Tag.FileName = $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName
				$fm_Rename.ShowDialog()
				UpdateFileList
			}
		}))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{ Text = '-' } ))
		[void]$cm_FileBrowser.MenuItems.Add((New-Object Windows.Forms.MenuItem -Property @{
			Text = 'P&roperties'
			add_Click = { PropertiesDialog $lv_FileBrowser.Items[$lv_FileBrowser.SelectedIndices[0]].Tag.FullName }
		}))
	#endregion MENU ITEMS FOR FILE BROWSER ##########################################################################################

	$fm_Browser.Controls.Add(($lv_FileBrowser = New-Object Windows.Forms.ListView -Property @{
		Anchor = 'Top,Left,Bottom,Right'
		Bounds = "5,30,490,260"
		View = 'Details'
		FullRowSelect = $True
		Multiselect = $False
		HideSelection = $False
		Margin = '0,0,0,0'
		ContextMenu = $cm_FileBrowser
		Tag = @{CurrentPath=''}
		add_KeyDown = {
			If ($_.KeyCode -eq 'F5') {
				$_.SuppressKeyPress = $True
				UpdateFileList
			} ElseIf ($_.KeyCode -eq 'Back' -or ($_.Alt -and $_.KeyCode -eq 'Up')) {
				$UpFolder_Func.invoke()
			}
		}
		add_ItemActivate = {
			If ($this.SelectedIndices.length) {
				$SelItem = $this.Items[$this.SelectedIndices[0]]
				If ($SelItem.Tag.GetFiles) {
					UpdateFileList -Path $SelItem.Tag.FullName
				} Else {
					Invoke-Item $SelItem.Tag.FullName
				}
			}
		}
		add_ColumnClick = {
			If ($_.Column -eq $global:Column) {
				$global:ReverseSort = -not $global:ReverseSort
			} Else {
				$global:ReverseSort = $True
			}
			$global:Column = $_.Column
			$global:lviArray_FileBrowser = $lviArray_FileBrowser | Sort-Object -Property @{
				Expression = {
					If ($_.SubItems[0].Text -eq '<Parent Folder>') { Return ' ' } # Always sort the parent folder entry to top

					$SortKey = ''
					If ($_.Tag.GetFiles) { # Detect that item is a folder
						$SortKey += ' ' # Prefix with space to always sort to top
					}

					If ($lv_FileBrowser.Columns[$Column].Text -eq 'Date') {
						$SortKey += $_.Tag.LastWriteTime.toString('s') # Sorting by date, return file date in sortable format
					} ElseIf ($lv_FileBrowser.Columns[$Column].Text -eq 'Size') {
						$SortKey += $_.Tag.Length
					} Else {
						$SortKey += $_.SubItems[$Column].Text
					}

					If ($lv_FileBrowser.Columns[$Column].Text -ne 'Name') {
						$SortKey += $_.SubItems[0].Text # If not already sorting by name, suffix the filename. This alphabetizes items with the same file type
					}
					Return $SortKey
				}
				Ascending = $global:ReverseSort
			}

			$lv_FileBrowser.BeginUpdate()
			$lv_FileBrowser.Items.Clear()
			$lv_FileBrowser.Items.AddRange($lviArray_FileBrowser)
			$lv_FileBrowser.EndUpdate()
		}
	}))

	[void]$lv_FileBrowser.Columns.Add('Name', 250)
	[void]$lv_FileBrowser.Columns.Add('Type', 120)
	[void]$lv_FileBrowser.Columns.Add('Date', 120)
	[void]$lv_FileBrowser.Columns.Add('Size', 50, 'Right')

	$fm_Browser.Controls.Add(($sb_Browser = New-Object Windows.Forms.StatusStrip))
	$sb_Browser.Items.Add(($sbl_Status = New-Object Windows.Forms.ToolStripStatusLabel))

#endregion SETUP BROWSER GUI ##########################################################################################

#region SETUP RENAME GUI ##########################################################################################
	# Rename dialog
	$fm_Rename = New-Object Windows.Forms.Form -Property @{
		Owner = $fm_Browser
		Text = 'Rename file...'
		StartPosition = 'Manual'
		ClientSize = '250,65'
		# Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Process -id $PID).Path)
		Icon = Icon-FromFilePath (Get-Process -id $PID).Path
		FormBorderStyle = 'FixedToolWindow'
		ShowInTaskbar = $False
		Tag = @{}

		# Setup key binds
		KeyPreview = $True
		add_Load = {
			$Mouse = [System.Windows.Forms.Cursor]::Position
			$this.SetDesktopLocation(
				($Mouse.X - $this.Width / 2),
				($Mouse.Y - $this.Height / 2)
			)
			$this.TopMost = $fm_Browser.TopMost
			$tb_RenameFile.Text = Split-Path $this.Tag.FileName -Leaf
			$tb_RenameFile.Select()
		}
	}

	$fm_Rename.Controls.Add(($tb_RenameFile = New-Object Windows.Forms.TextBox -Property @{
		Bounds = '5,5, 240,20'
		add_TextChanged = {
			$InvalidChars = '[\\/:*?"<>|]'
			If ($this.Text -match $InvalidChars) {
				MsgBox "A file name cannot contain the following characters:`n \ / : * ? `" < > |"
				$this.Text = $this.Text -replace $InvalidChars,''
			}
		}
	}))
	$fm_Rename.Controls.Add(($bt_RenameOK = New-Object Windows.Forms.Button -Property @{
		Bounds = '80,35, 80,25'
		Text = 'OK'
		add_Click = {
			$Success = Rename-Item $fm_Rename.Tag.FileName $tb_RenameFile.Text -PassThru
			If (!$Success) {
				MsgBox "Unable to rename item:`n`n$($error[0].toString())" 0 'Error'
			} Else {
				$fm_Rename.Close()
			}
		}
	}))
	$fm_Rename.AcceptButton = $bt_RenameOK

	$fm_Rename.Controls.Add(($bt_RenameCancel = New-Object Windows.Forms.Button -Property @{
		Bounds = '165,35, 80,25'
		Text = 'Cancel'
		add_Click = { $fm_Rename.Close() }
	}))
	$fm_Rename.CancelButton = $bt_RenameCancel


#endregion SETUP RENAME GUI ##########################################################################################

[System.GC]::Collect() # Garbage collection to lower RAM usage
[Windows.Forms.Application]::Run(($AppContext = New-Object Windows.Forms.ApplicationContext($fm_Browser)))