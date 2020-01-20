# Prerequisities
- installed MS Office (32bit / 64bit)
- installed latest version of SCIA Engineer (19.1 patch 2) (32bit / 64bit)


# SciaOpenAPI_example_VBA
## Prepare the environment
* install MS Office (32bit / 64bit)
* install SCIA Engineer (32bit / 64bit according to MS Office)
    * following steps should be done automatically during SCIA Engineer setup:
        * install .NET FW 4.6.1 or newer
      
        * find your .NET FW install directory (e.g. c:\Windows\Microsoft.NET\Framework\v4.0.30319 for 32bit, c:\Windows\Microsoft.NET\Framework64\v4.0.30319 for 64bit) 
        * start command line AS ADMINISTRATOR and navigate to that directory and execute following command (if needed, adjust the actual path to SCIA Engineer install directory):
```
for32bit:
regasm "c:\Program Files (x86)\SCIA\Engineer19.0\SCIA.OpenAPI.dll" /tlb:"c:\Program Files (x86)\SCIA\Engineer19.0\SCIA.OpenAPI.tlb" /codebase

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
   * Add VBA Microsoft Scripting Runtime reference
* You can validate that you can see SCIA.OpenAPI.dll classes in View > Object Browser
* Implement Sub for deleting Temp folder of SCIA Engineer
```VBA
Private Sub DeleteTemp(TempFolder As String)

    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject

    If FSO.FolderExists(TempFolder) Then
        Call FSO.DeleteFolder(TempFolder, True)
    End If
End Sub
```
* You could also implement Sub and functions for generation of GUIDs
```VBA
   Public Function Get_NewGUID() As String
    'Returns GUID as string 36 characters long

    Randomize

    Dim r1a As Long
    Dim r1b As Long
    Dim r2 As Long
    Dim r3 As Long
    Dim r4 As Long
    Dim r5a As Long
    Dim r5b As Long
    Dim r5c As Long

    'randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
    r1a = RandomBetween(0, 65535)
    r1b = RandomBetween(0, 65535)
    r2 = RandomBetween(0, 65535)
    r3 = RandomBetween(16384, 20479)
    r4 = RandomBetween(32768, 49151)
    r5a = RandomBetween(0, 65535)
    r5b = RandomBetween(0, 65535)
    r5c = RandomBetween(0, 65535)

    Get_NewGUID = (PadHex(r1a, 4) & PadHex(r1b, 4) & "-" & PadHex(r2, 4) & "-" & PadHex(r3, 4) & "-" & PadHex(r4, 4) & "-" & PadHex(r5a, 4) & PadHex(r5b, 4) & PadHex(r5c, 4))

End Function

Public Function Floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    'From: http://www.tek-tips.com/faqs.cfm?fid=5031
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
        Floor = Int(X / Factor) * Factor
End Function

Public Function RandomBetween(ByVal StartRange As Long, ByVal EndRange As Long) As Long
    'Based on https://msdn.microsoft.com/en-us/library/f7s023d2(v=vs.90).aspx
    '         randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
        RandomBetween = CLng(Floor((EndRange - StartRange + 1) * Rnd())) + StartRange
End Function

Public Function PadLeft(text As Variant, totalLength As Integer, padCharacter As String) As String
    'Based on https://stackoverflow.com/questions/12060347/any-method-equivalent-to-padleft-padright
    ' with a little more checking of inputs

    Dim s As String
    Dim inputLength As Integer
    s = CStr(text)
    inputLength = Len(s)

    If padCharacter = "" Then
        padCharacter = " "
    ElseIf Len(padCharacter) > 1 Then
        padCharacter = Left(padCharacter, 1)
    End If

    If inputLength < totalLength Then
        PadLeft = String(totalLength - inputLength, padCharacter) & s
    Else
        PadLeft = s
    End If

End Function

Public Function PadHex(number As Long, length As Integer) As String
    PadHex = PadLeft(Hex(number), 4, "0")
End Function
```
* You can start your VBA development using the SCIA.OpenAPI.dll functionality
* When finishing your work with Scia.OpenAPI in your script, don't forget to call the SCIA.OpenAPI.Environment.Dispose() method for your specific environemnt object!!!

## Remarks
* you can get inspiration from enclosed example
* using of several versions of Scia Engineer at once: communication between VBA and Scia Engineer is based on COM technology. During SCIA Engineer installation the SCIA.OpenAPI.dll is registered into windows registry using the c# registration utility REGASM.EXE. In VBA in Tools>Reference you see only currenlty registered version of SCIA.OpenAPI.dll. If you want to use previous version of SCIA.OpenAPI.dll, you must unregister current version (regasm "/unregister") and register desired version into registry.
