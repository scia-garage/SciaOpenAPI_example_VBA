# SciaOpenAPI_example_VBA
## Prepare the environment
* install MS Office (32bit / 64bit)
* install SCIA Engineer (32bit / 64bit according to MS Office)
    * following steps should be done automatically during SCIA Engineer setup:
        * install .NET FW 4.6.1 or newer
        * start command line AS ADMINISTRATOR and navigate to SEn install directory and run "ep_regsvr32 esa.exe" (for 64bit "ep_regsvr64 esa.exe")
        * find your .NET FW install directory (e.g. c:\Windows\Microsoft.NET\Framework\v4.0.30319 for 32bit, c:\Windows\Microsoft.NET\Framework64\v4.0.30319 for 64bit) 
        * start command line AS ADMINISTRATOR and navigate to that directory and execute following command (if needed, adjust the actual path to SCIA Engineer install directory):
```
for32bit:
regasm "c:\Program Files\SCIA\Engineer19.0 (x86)\SCIA.OpenAPI.dll" /tlb:"c:\Program Files\SCIA\Engineer19.0 (x86)\SCIA.OpenAPI.tlb" /codebase

for64bit:
regasm "c:\Program Files\SCIA\Engineer19.0\SCIA.OpenAPI.dll" /tlb:"c:\Program Files\SCIA\Engineer19.0\SCIA.OpenAPI.tlb" /codebase
```
* run SCIA Engineer to check it works (e.g. set protection, etc.)
* copy the .\res.\excel.config file to Excel.exe location (e.g. for 32bit MS Office c:\Program Files (x86)\Microsoft Office\root\Office16\ )



## Start your development in VBA...for instance:
* start Excel
* create new Excel sheet
* File > Options > Cutomize Ribbon
* In Main Tabs pan, check the Developer checkbox, click OK
* In Excel sheet on Developer tab click the "Visual Basic"
* In Visual Basic editor select Tools>References and:
   * check the "SCIA API for external developers"
* You can validate that you can see SCIA.OpenAPI.dll classes in View > Object Browser
* You can start your VBA development using the SCIA.OpenAPI.dll functionality

## Remarks
* you can get inspiration from enclosed example
* using of several versions of Scia Engineer at once: communication between VBA and Scia Engineer is based on COM technology. During SCIA Engineer installation the SCIA.OpenAPI.dll is registered into windows registry using the c# registration utility REGASM.EXE. In VBA in Tools>Reference you see only currenlty registered version of SCIA.OpenAPI.dll. If you want to use previous version of SCIA.OpenAPI.dll, you must unregister current version (regasm "/unregister") and register desired version into registry.
