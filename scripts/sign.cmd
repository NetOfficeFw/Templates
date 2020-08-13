@echo off

vsixsigntool.exe sign /f "%1" /p "%2" /sha1 "AC6DBFFB1BF8B62281DEB8641023A66CDDC5DB57" ^
  /tr http://timestamp.comodoca.com/rfc3161 /td sha256 /v ^
  NetOfficeExtension.vsix
