@echo off
:: ========================================================
:: أداة إعداد وتثبيت مكتبة المصطفى التلقائية
:: ========================================================
title اعداد مكتبة المصطفى

:: 1. التحقق من صلاحيات المسؤول (Admin) وطلبها إذا لم تتوفر
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo [!] جاري طلب صلاحيات المسؤول للتشغيل...
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B
)

cls
echo ========================================================
echo         مــكــتــبــة الــمــصــطــفــى - جاري الإعداد...
echo ========================================================

:: 2. التحقق من تثبيت Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [x] لغة Python غير مثبتة على هذا الجهاز!
    echo [i] يرجى تثبيت بايثون أولاً من الموقع الرسمي أو متجر ويندوز، ثم إعادة تشغيل الأداة.
    pause
    exit /B
) else (
    echo [✓] تم التحقق من وجود Python.
)

:: 3. التحقق من تثبيت المكتبات وتنزيلها تلقائياً إن لم تكن موجودة
echo [i] جاري التحقق من المكتبات المطلوبة (python-docx, colorama)...
python -c "import docx, colorama" >nul 2>&1
if %errorlevel% neq 0 (
    echo [i] بعض المكتبات غير موجودة، جاري التثبيت التلقائي عبر PIP...
    python -m pip install --upgrade pip
    python -m pip install python-docx colorama
) else (
    echo [✓] جميع المكتبات المطلوبة مثبتة مسبقاً.
)

:: 4. التحقق من وجود ملف app.py، وإذا لم يوجد، سحبه من مستودع غيت هب
if not exist "%~dp0app.py" (
    echo [i] ملف التشغيل غير موجود محلياً، جاري سحبه من مستودع GitHub...
    git --version >nul 2>&1
    if %errorlevel% eq 0 (
        git clone https://github.com/MUSTAFA-mbx5/ai.git "%TEMP%\mustafa_ai_repo"
        if exist "%TEMP%\mustafa_ai_repo\app.py" (
            copy /y "%TEMP%\mustafa_ai_repo\app.py" "%~dp0app.py"
            echo [✓] تم تنزيل الملف بنجاح.
        )
    ) else (
        echo [!] تحذير: أداة Git غير مثبتة، يرجى التأكد من وجود ملف app.py في نفس المجلد.
    )
)

:: 5. إنشاء اختصار على سطح المكتب باسم "مكتبة المصطفى"
set SCRIPT="%TEMP%\CreateShortcut.vbs"
echo Set oWS = WScript.CreateObject("WScript.Shell") > %SCRIPT%
echo sLinkFile = "%USERPROFILE%\Desktop\مكتبة المصطفى.lnk" >> %SCRIPT%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%
echo oLink.TargetPath = "cmd.exe" >> %SCRIPT%
echo oLink.Arguments = "/c python ""%~dp0app.py""" >> %SCRIPT%
echo oLink.WorkingDirectory = "%~dp0" >> %SCRIPT%
echo oLink.Description = "أداة تنسيق مستندات مكتبة المصطفى" >> %SCRIPT%
echo oLink.Save >> %SCRIPT%
cscript /nologo %SCRIPT%
del %SCRIPT%

echo [✓] تم إنشاء اختصار "مكتبة المصطفى" على سطح المكتب بنجاح!
echo ========================================================
echo جاري تشغيل الأداة الآن...
timeout /t 2 >nul

:: 6. تشغيل سكربت بايثون الرئيسي
if exist "%~dp0app.py" (
    python "%~dp0app.py"
) else (
    echo [x] خطأ: لم يتم العثور على ملف app.py لتشغيله!
    pause
)
