# Alternate slide.ps1
# By Scott Kasai
#
# Purpose:
# This script is designed to be run from Task Scheduler to do the following:
# * Kill running Powerpoint if Powerpoint Process exists
# * Copy Powerpoint Slide to Local Computer
# * (Re)start Powerpoint with New Slide
 
# Define the file we are checking on Google Team Drive, using the specified path on
# the PC, IE: G:\Team Drives\Department-Drive\Presentation\presentation.ppsx
$file = "<Google Team Drive Document Path on Computer>"
 
 
# Define the file name and location to save locally and call from:
# IE: c:\admin\slide\presentation.ppsx
$file2 = "<Local location on computer>"
 
# Get Process ID of Powerpoint if it is running, otherwise, continue silently
# if there is no process.
$process1 = get-process powerpnt -ErrorAction SilentlyContinue
 
# Check and see if Powerpoint isn't Running
if ($process1 -ne $null)
{
# Stop current Powerpoint Process
  Stop-Process $process1
}
# Copy source file locally and then start presentation
copy-item -Path $file1 -Destination $file2
invoke-item $file2
