# SciaOpenAPI_example_VBA
## Prepare the environment
- install SCIA Engineer
- install MS Office
- find your .NET FW install directory (e.g. c:\Windows\Microsoft.NET\Framework\v4.0.30319) 
- start command line AS ADMINISTRATOR and navigate to that directory
- run regasm "c:\Program Files (x86)\SCIA\Engineer19\SCIA.OpenAPI.dll" /tlb:"c:\Program Files (x86)\SCIA\Engineer19\SCIA.OpenAPI.tlb" (if needed, adjust the actual path to SCIA Engineer install directory

## Start your development in VBA...for instance:
- create new Excel sheet
- File > Options > Cutomize Ribbon
- Under Customize the Ribbon and under Main Tabs, select the Developer check box
- In Excel sheet on Developer tab click the "Visual Basic"
- In Visual Basic editor select Tools>References and check the "SCIA API for external developers"
- You can validate that you can see SCIA.OpenAPI.dll classes in View > Object Browser
- You can start your VBA development using the SCIA.OpenAPI.dll functionality
