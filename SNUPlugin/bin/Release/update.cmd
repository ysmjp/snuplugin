@echo off
call:delete_and_copy "E:\Projects\Github\DoBrainRemaster\Assets\Plugins"
call:delete_and_copy "E:\Projects\Unity\untitled\Assets\Plugins"
exit /b

:delete_and_copy <proj.dir>
set "proj.dir=%~1"
set "plugin.name=SNUPlugin.dll"
del "%proj.dir%\%plugin.name%" >nul 2>&1
del "%proj.dir%\%plugin.name%.meta" >nul 2>&1
copy "%plugin.name%" "%proj.dir%\%plugin.name%"
exit /b