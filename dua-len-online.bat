@echo off
cd /d "%~dp0"
echo Buoc 1: Luu lai cac thay doi con dang do o may nay...
echo.
git add -A
git commit -m "Cap nhat truoc khi dong bo" --allow-empty
echo.
echo Buoc 2: Lay ve nhung thay doi tren GitHub (neu co)...
echo.
git pull origin main --no-edit
echo.
echo ================================================
if errorlevel 1 (
  echo CO LOI o buoc "lay ve" - xem dong chu mau do o tren.
  echo Neu thay chu "CONFLICT" thi DUNG lai, chup man hinh gui cho Claude xem tiep.
  echo ================================================
  pause
  exit /b 1
)
echo Buoc 2 xong. Dang qua Buoc 3: Dua code len GitHub...
echo ================================================
echo.
git push origin main
echo.
echo ================================================
echo Xong roi (hoac neu co loi thi xem dong chu do o tren).
echo Neu co hien cua so trinh duyet hoi dang nhap GitHub thi dang nhap binh thuong.
echo ================================================
pause
