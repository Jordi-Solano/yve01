@echo off
echo ============================================================
echo  Yve.01 - Git Push: Add AI chat copilot
echo ============================================================
cd /d C:\Users\Jo\yve01

echo Eliminando index.lock si existe...
if exist .git\index.lock del /f .git\index.lock
echo.

echo Haciendo git add...
git add dashboard.py SKILL_YVE01_UPDATED.md oracle_auth.py oracle_lector_facturas.py oracle_crear_asientos.py oracle_actualizar_estado.py oracle_pipeline.py lector_drr.py ORACLE_INTEGRATION.md
echo.

echo Estado del repositorio:
git status --short
echo.

echo Haciendo commit...
git commit -m "Add AI chat copilot"
echo.

echo Haciendo push...
git push
echo.

echo ============================================================
echo  Push completado. Revisa los mensajes arriba.
echo ============================================================
pause
