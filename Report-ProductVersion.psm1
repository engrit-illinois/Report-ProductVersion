# Documentation home: https://github.com/engrit-illinois/Report-ProductVersion
# By mseng3

function Report-ProductVersion {

	param(
		# Array or wildcard query of computers to query
		[Parameter(Position=0,Mandatory=$true,ParameterSetName="Array")]
		[string[]]$Computers,
		
		# Name of collection containing computer to query
		[Parameter(Position=0,Mandatory=$true,ParameterSetName="Collection")]
		[string]$Collection,
		
		# Product to search for
		[Parameter(Mandatory=$true)]
		[string]$Product,
		
		# Full path to log file
		[string]$Log,
		
		# Full path to CSV file
		[string]$Csv,
		
		# Logs the full set of results every time a machine is polled
		# Useful in case things are hanging up, so you get at least some data
		[switch]$LogIncrementalProgress,
		
		# Site code
		[string]$SiteCode="MP0",
		
		# SMS Provider machine name
		[string]$Provider="sccmcas.ad.uillinois.edu",
		
		# ConfigurationManager Powershell module path
		[string]$CMPSModulePath="$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"

	)

	function log($msg) {
		$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss:ffff"
		$msg = "[$ts] $msg"
		Write-Host $msg
		if($Log) {
			Write-Output $msg | Out-File $Log -Append
		}
	}
	
	# Loads the ConfigMgr Powershell module, and connects to the Site.
	function Prepare-SCCM {
		# Customizations
		$initParams = @{}
		#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
		#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

		# Import the ConfigurationManager.psd1 module 
		if((Get-Module ConfigurationManager) -eq $null) {
			Import-Module $CMPSModulePath @initParams -Scope Global
		}

		# Connect to the site's drive if it is not already present
		if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
			New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Provider @initParams
		}

		# Set the current location to be the site code.
		Set-Location "$($SiteCode):\" @initParams
	}
	
	function Get-CompNameList($compNames) {
		$list = ""
		foreach($name in $compNames) {
			$list = "$list, $name"
		}
		$list = $list.Substring(2,$list.length - 2) # Remove leading ", "
		$list
	}
	
	function Get-CompNames {
		log "Getting list of computer names..."
		if($Computers) {
			log "List was given as an array." -l 1 -v 1
			$compNames = @()
			foreach($comp in @($Computers)) {
				$compResults = (Get-ADComputer -Filter "Name -like '$comp'" | Select Name).Name
				foreach($result in @($compResults)) {
					$compNames += @($result)
				}
			}
			$list = Get-CompNameList $compNames
			log "Found $($compNames.count) computers in given array: $list." -l 1
		}
		elseif($Collection) {
			log "List was given as a collection. Getting members of collection: `"$Collection`"..." -l 1 -v 1
			Prepare-SCCM
			$colObj = Get-CMCollection -Name $Collection
			if(!$colObj) {
				log "The given collection was not found!" -l 1
			}
			else {
				# Get comps
				$comps = Get-CMCollectionMember -CollectionName $Collection | Select Name,ClientActiveStatus
				if(!$comps) {
					log "The given collection is empty!" -l 1
				}
				else {
					# Sort by active status, with active clients first, just in case inactive clients might come online later
					# Then sort by name, just for funsies
					$comps = $comps | Sort -Property @{Expression = {$_.ClientActiveStatus}; Descending = $true}, @{Expression = {$_.Name}; Descending = $false}
					
					$compNames = $comps.Name
					$list = Get-CompNameList $compNames
					log "Found $($compNames.count) computers in `"$Collection`" collection: $list." -l 1
				}
			}
		}
		else {
			log "Somehow neither the -Computers, nor -Collection parameter was specified!" -l 1
		}
		
		log "Done getting list of computer names." -v 2
		
		$compNames | Sort
	}

	function Get-Data($comps) {
		# Initialize array
		$data = @()

		# Attempt to poll each machine
		foreach($comp in $comps) {
			log "Polling computer `"$comp`"..."
			$success = $true
			$results = "unknown"
			$productWMI = $Product.Replace("*","%")
			try {
				$result1 = Get-WmiObject -Query "select * from win32_product where name like '$productWMI'" -computername $comp -ErrorAction Stop
				$result2 = Get-WmiObject -Class "Win32Reg_AddRemovePrograms" -Computername $comp -ErrorAction Stop | Where { $_.DisplayName -like $Product }
			}
			catch [System.Runtime.InteropServices.COMException] {
				if($_.Exception.Message.trim() -eq "The RPC server is unavailable.") {
					log "    Could not contact `"$comp`"!"
				}
				else {
					log "    Unknown COMException!"
				}
				$success = $false
			}
			catch {
				log "    Unknown error!"
				$success = $false
			}
			
			if($success) {
				$results = @($result1) + @($result2)
				log "    Success. Found $(@($results).count) matching products."
			
				if(@($results).count -ge 1) {
					foreach($result in @($results)) {
						$result | Add-Member -NotePropertyName "source" -NotePropertyValue "unknown"
						if(($result.DisplayName) -and (!($result.Name))) {
							$result.source = "Win32Reg_AddRemovePrograms"
							$result | Add-Member -NotePropertyName "Name" -NotePropertyValue $result.DisplayName
						}
						elseif(($result.Name) -and (!($result.DisplayName))) {
							$result.source = "Win32_Product"
							$result | Add-Member -NotePropertyName "DisplayName" -NotePropertyValue $result.Name
						}
						else {
							log "Found result that had either both, or neither of Name and DisplayName!"
						}
						
						log "        Source: `"$($result.source)`", Name: `"$($result.Name)`", Version: `"$($result.Version)`"."
					}
				}
			}
			else {
				$results = @([PSCustomObject]@{
					"PSComputerName" = $comp
					"Source" = "Error"
					"Name" = "Error"
					"DisplayName" = "Error"
					"Version" = "Error"
				})
			}
			
			$data += @($results)
			
			#log "    Done."
			
			if($LogIncrementalProgress) {
				log "    $(($data | Format-Table | Out-String).trim())"
			}
		}
		
		$data
	}

	function Output-Data($data) {
		$sdata = $data | Select PSComputerName,Name,Version | Sort PSComputerName
		
		# Output data to CSV
		if($Csv) {
			$sdata | Export-Csv -Path $Csv -NoTypeInformation -Encoding Ascii
		}
		
		# Output data to log and console
		log " "
		log ($sdata | Format-Table -AutoSize -Wrap | Out-String).trim()
	}
	
	
	$myPWD = $pwd.path
	$comps = Get-CompNames
	if($comps -ne $false) {
		$data = Get-Data $comps
		Output-Data $data
	}
	Set-Location $myPWD
	
	log " "
	log "EOF"
	log " "
}
