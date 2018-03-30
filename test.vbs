currentpath = createobject("Scripting.FileSystemObject").GetFolder(".").Path

Dim width,height
width = CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
height = CreateObject("HtmlFile").ParentWindow.Screen.AvailHeight

Dim line_for_each()
Set fso = CreateObject("Scripting.FileSystemObject")
Set file1 = fso.OpenTextFile(currentpath & "\address.txt",1,False)  

Dim i
i = 0 
DO While file1.AtEndOfStream <> True
redim preserve line_for_each(i)
line_for_each(i) = line & file1.ReadLine & vbcrlf
i = i + 1
loop

Dim l_num
Dim c_num

l_num = cint(mid(line_for_each(0),1,1))
c_num = cint(mid(line_for_each(0),3,1))

Dim objWMIService
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

Dim objStartupInfo()
Dim j
j = 0

For il = 1 To l_num
For jl = 1 To c_num
redim preserve objStartupInfo(j)
Set objStartupInfo(j) = objWMIService.Get("Win32_ProcessStartup")
objStartupInfo(j).SpawnInstance_
objStartupInfo(j).Y = cint(height/l_num)*(il-1)
objStartupInfo(j).XSize = 40
objStartupInfo(j).X = cint(width/c_num)*(jl-1)
objStartupInfo(j).YSize = 40
j = j+1
Next
Next

Dim objNewProcess
Set objNewProcess = objWMIService.Get("Win32_Process")

Dim lin
Dim lin_string
Dim col
Dim col_string
lin = cint((height-150)/l_num/15)
col = cint((width-40)/c_num/8)

lin_string = cstr(lin)
col_string = cstr(col)


Dim intPID
Dim errRtn()
For e = 1 to i-1
redim preserve errRtn(e)
errRtn(e) = objNewProcess.Create(currentpath & "\ping.bat " & lin_string & " " & col_string & " " & line_for_each(e), Null, objStartupInfo(e-1), intPID)
Next
