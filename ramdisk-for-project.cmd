@echo off
set project_name=
for %%i in ("%cd%") do set project_name=%%~ni

rmdir /s /q out
mkdir R:\out.%project_name%
mklink /d out R:\out.%project_name%

rmdir /s /q build
mkdir R:\build.%project_name%
mklink /d build R:\build.%project_name%