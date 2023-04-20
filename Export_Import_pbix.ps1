﻿# RUN if error
Set-ExecutionPolicy RemoteSigned

# Run the following command to install it if not already installed: 
Install-Module -Name MicrosoftPowerBIMgmt

# Log in to Power BI Service
Login-PowerBI  –Environment Public


# First, Collect all (or one) of the workspaces in a parameter called PBIWorkspace
#$PBIWorkspace = Get-PowerBIWorkspace 							# Collect all workspaces you have access to
#$PBIWorkspace = Get-PowerBIWorkspace -Name 'My Workspace Name' # Use the -Name parameter to limit to one workspace

$prod = 'Workspace1', 'Workspace2'

$PBIWorkspace = @()
ForEach($ws in $prod)
{
	$PBIWorkspace +=(Get-PowerBIWorkspace -Name $ws)[0]
}

# Now collect todays date
$TodaysDate = Get-Date –Format "yyyyMMdd" 
# Almost finished: Build the outputpath. This Outputpath creates a news map, based on todays date
$OutPutPath = "C:\Users\Kateryna_Nemchenko\Desktop\PowerBIReportsBackup" + $TodaysDate 


# Now loop through the workspaces, hence the ForEach
ForEach($Workspace in $PBIWorkspace)
{
	# For all workspaces there is a new Folder destination: Outputpath + Workspacename
	$Folder = $OutPutPath + "\" + $Workspace.name 
	# If the folder doens't exists, it will be created
	If(!(Test-Path $Folder))
	{
		New-Item –ItemType Directory –Force –Path $Folder
	}
	# At this point, there is a folder structure with a folder for all your workspaces 
	
	
	# Collect all (or one) of the reports from one or all workspaces 
	$PBIReports = Get-PowerBIReport –WorkspaceId $Workspace.Id		# Collect all reports from the workspace we selected.
	#$PBIReports = Get-PowerBIReport -WorkspaceId $Workspace.Id -Name "My Report Name" # Use the -Name parameter to limit to one report
		
		# Now loop through these reports: 
		ForEach($Report in $PBIReports)
		{
			# Your PowerShell comandline will say Downloading Workspacename Reportname
			Write-Host "Downloading "$Workspace.name":" $Report.name 
			
			# The final collection including folder structure + file name is created.
			$OutputFile = $OutPutPath + "\" + $Workspace.name + "\" + $Report.name + ".pbix"
			
			# If the file exists, delete it first; otherwise, the Export-PowerBIReport will fail.
			if (Test-Path $OutputFile)
			{
				Remove-Item $OutputFile
			}
			
			# The pbix is now really getting downloaded
			Export-PowerBIReport –WorkspaceId $Workspace.ID –Id $Report.ID –OutFile $OutputFile
            
			# Processing the file with Python
			Write-Host "Proccessing "$Workspace.name":" $Report.name
            python bookmarks_update.py --directory $OutPutPath --workspace $Workspace.name --report $Report.name
            
            # New-PowerBIReport doesn't like dots in Name parameter
			Write-Host "Sending to PBI portal "$Workspace.name":" $Report.name
            New-PowerBIReport -Path $OutputFile -Name $Report.name -WorkspaceId $Workspace.ID -ConflictAction CreateOrOverwrite
            
		}
}


#Set-ExecutionPolicy Restricted