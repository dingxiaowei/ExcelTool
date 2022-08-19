@echo off
cd config

echo "开始拷贝bytes文件"
for %%i in (*.bytes) do (
    echo begin copy... %%i
    copy /y %%~nxi ..\..\hola\unity\hola_unity\Assets\Resources\Config\%%~nxi
    echo copy complate ... %%i
)
echo "bytes文件拷贝完成"

echo "开始拷贝cs文件"
for %%i in (*.cs) do (
    echo begin copy... %%i
    copy /y %%~nxi ..\..\hola\unity\hola_unity\Assets\Src\MMHouse\Backend\Config\%%~nxi
    echo copy complate ... %%i
)
echo "cs文件拷贝完成"

echo "删除生成的文件"
for %%i in (*.bytes) do (
    del %%i
    echo delete complate ... %%i
)
for %%i in (*.cs) do (
    del %%i
    echo delete complate ... %%i
)
echo "删除完成"

pause