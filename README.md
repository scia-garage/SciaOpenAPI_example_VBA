# SciaOpenAPI_example_VBA
## Prepare the environment
* install MS Office (32bit / 64bit)
* install SCIA Engineer (32bit / 64bit according to MS Office)
    * following steps should be done automatically during SCIA Engineer setup:
        * install .NET FW 4.6.1 or newer
        * find your .NET FW install directory (e.g. c:\Windows\Microsoft.NET\Framework\v4.0.30319 for 32bit, c:\Windows\Microsoft.NET\Framework64\v4.0.30319) 
        * start command line AS ADMINISTRATOR and navigate to that directory and execute following command (if needed, adjust the actual path to SCIA Engineer install directory):
```
 regasm "c:\Program Files (x86)\SCIA\Engineer19\SCIA.OpenAPI.dll" /tlb:"c:\Program Files (x86)\SCIA\Engineer19\SCIA.OpenAPI.tlb" /codebase 
```
* copy the .\res.\excel.config file to Excel.exe location (e.g. for 32bit MS Office c:\Program Files (x86)\Microsoft Office\root\Office16\ )
* copy ESAAtl80Extern.dll and FemBase.dll from Scia Engineer install directory to the Excel.exe location (e.g. c:\Program Files (x86)\Microsoft Office\root\Office16\)


## Start your development in VBA...for instance:
- start Excel AS ADMINISTRATOR
- create new Excel sheet
- File > Options > Cutomize Ribbon
- Under Customize the Ribbon and under Main Tabs, select the Developer check box
- In Excel sheet on Developer tab click the "Visual Basic"
- In Visual Basic editor select Tools>References and:
-- check the "SCIA API for external developers"
-- click Browse and find "c:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
- You can validate that you can see SCIA.OpenAPI.dll classes in View > Object Browser
//- in your script you must change current directory to Scia Engineer install directory using ChDir "c:\Program Files (x86)\SCIA\Engineer19.0\" (because of ESAAtl80Extern.dll)
- You can start your VBA development using the SCIA.OpenAPI.dll functionality
=========================================
