set j=0
for /f "delims=""" %%i in (address.txt) do (

taskkill /fi "windowtitle eq %%i"
)

