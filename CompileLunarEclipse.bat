:: Compile script for LunarEclipse

:: Copy all relevant data folders from Source to Build Folder
if not exist Build mkdir Build
xcopy /s /e /y .\Source\Resources .\Build\Resources\
xcopy /s /e /y ".\Source\Constraint Data" ".\Build\Constraint Data\"

:: Set common C# compiler resources as variables
SET VarianRes=/r:VMS.TPS.Common.Model.API.dll /r:VMS.TPS.Common.Model.Types.dll
SET WpfRes=/r:WPF\PresentationCore.dll /r:WPF\PresentationFramework.dll^
 /r:WPF\WindowsBase.dll /r:System.Xaml.dll
SET PdfRes=/r:PdfSharp.dll /r:MigraDoc.DocumentObjectModel.dll^
 /r:MigraDoc.Rendering.dll^

:: Compile each C# script as an ESAPI DLL to use in Eclipse
csc.exe /target:library /out:.\Build\CollisionAvoid.esapi.dll .\Source\CollisionAvoid.cs^
 %VarianRes% %WpfRes%

csc.exe /target:library /out:.\Build\DoseEval.esapi.dll .\Source\DoseEval.cs^
 .\Source\ReportDoseEval.cs %VarianRes% %WpfRes% %PdfRes%

csc.exe /target:library /out:.\Build\PlanScan.esapi.dll .\Source\PlanScan.cs^
 %VarianRes% %WpfRes%
