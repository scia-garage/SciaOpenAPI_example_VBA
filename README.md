# SciaOpenAPI_example_VBA
## Prepare the environment
* install MS Office
* install SCIA Engineer
    * following steps should be done automatically during SCIA Engineer setup:
        * install .NET FW 4.6.1 or newer
        * find your .NET FW install directory (e.g. c:\Windows\Microsoft.NET\Framework\v4.0.30319) 
        * start command line AS ADMINISTRATOR and navigate to that directory and execute following command (if needed, adjust the actual path to SCIA Engineer install directory):
```
 regasm "c:\Program Files (x86)\SCIA\Engineer19\SCIA.OpenAPI.dll" /tlb:"c:\Program Files (x86)\SCIA\Engineer19\SCIA.OpenAPI.tlb" 
 ```
* copy from c:\Windows\Microsoft.NET\Framework\v4.0.30319 to C:\Program Files (x86)\Microsoft Office\root\vfs\Windows\Microsoft.NET\Framework\v4.0.30319
* follow https://support.microsoft.com/cs-cz/help/2683270/vba-error-handling-may-result-in-search-for-winhelp-ini (my path C:\Program Files (x86)\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA7.1\1029)

## Start your development in VBA...for instance:
- create new Excel sheet
- File > Options > Cutomize Ribbon
- Under Customize the Ribbon and under Main Tabs, select the Developer check box
- In Excel sheet on Developer tab click the "Visual Basic"
- In Visual Basic editor select Tools>References and check the "SCIA API for external developers"
- You can validate that you can see SCIA.OpenAPI.dll classes in View > Object Browser
- You can start your VBA development using the SCIA.OpenAPI.dll functionality
