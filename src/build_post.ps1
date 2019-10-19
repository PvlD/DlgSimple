#
# on_build.ps1
#

param($mode,  $proj_dir)


Write-Host " mode:"$mode
Write-Host " proj_dir:"$proj_dir


$env:Path += ";C:\Program Files\nodejs\"



if ($mode -eq "Debug")
{
npm  run bldd

if ( $lastExitCode -ne  0 )
	{
	exit 1
	}
	#Write-Host " run bldd ExitCode :" $lastExitCode
}
else 
{
npm  run bldr
if ( $lastExitCode -ne  0 )
	{
	exit 1
	}

}
