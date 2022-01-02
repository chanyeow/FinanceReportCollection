@echo off
@title resource build

if exist financeData (
   echo "resouce is exist"
) else (
svn://1.117.42.10:3690/financeData
)

pause