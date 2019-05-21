# slide.ps1
# By Scott Kasai
#
# Purpose:
# This script is designed to be run from Task Scheduler to do the following:
# * Check and see if Powerpoint is running (which it should with the display item)
#   If it is not, copy the relevant Powerpoint from Google Team Drive to the computer
#   then run it.
# * If Powerpoint is running, check and see if the file on Google Team Drive's Last
#   Modified is less than our specified check time
#   If it is, stop Powerpoint, copy the new Powerpoint file locally and restart the
#   presentation.
 
 
# Define the file we are checking on Google Team Drive, using the specified path on
# the PC, IE: G:\Team Drives\Department-Drive\Presentation\presentation.ppsx
$file1 = "<Google Team Drive Document Path on Computer>"
 
 
# Define the file name and location to save locally and call from:
# IE: c:\admin\slide\presentation.ppsx
$file2 = "<Local location on computer>"
 
 
# Get current time
$curtime = Get-Date
# Set Check time, in this example, -25 is used to specify time 25 minutes before
# current time.
$checktime = $curtime.AddMinutes(-25)
 
 
# Get last modified time from source file
$lastmod = (get-time $file).lastwritetime
 
 
# Get Process ID of Powerpoint if it is running, otherwise, continue silently
# if there is no process.
$process1 = get-process powerpnt -ErrorAction SilentlyContinue
 
 
# Check and see if Powerpoint isn't Running
if ($process1 -eq $null)
{
# Copy source file locally and then start presentation
  copy-item -Path $file1 -Destination $file2
  invoke-item $file2
}
else
{
# Check and see if source file has changed within the specified checktime
# range
  if ($checktime -lt $lastmod)
  {
# Stop current Powerpoint Process, copy source file locally and then restart presentation
    Stop-Process $process1
    copy-item -Path $file1 -Destination $file2
    invoke-item $file2
  }
}
