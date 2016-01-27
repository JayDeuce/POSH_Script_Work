#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.29
# Created on:   3/9/2014 2:10 AM
# Created by:   Jonathan Durant
# Organization: 
# Filename:     
#========================================================================


[string]$Date = [dateTime]::now.toshortdatestring()
[string]$Date = $Date.Replace("/","-")
[string]$Time = [dateTime]::now.toshorttimestring()
[string]$Time = $Time.Replace(":","-").Replace(" ","-")
[string]$DateTime = $Date + "_" + $Time