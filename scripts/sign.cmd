@echo off

vsixsigntool.exe sign /f "%1" /p "%2" /sha1 "CA24B654C5FFAC7D97BD196269BD0E80FADF1841" ^
  /tr http://timestamp.comodoca.com/rfc3161 /td sha256 /v ^
  NetOfficeExtension.vsix
