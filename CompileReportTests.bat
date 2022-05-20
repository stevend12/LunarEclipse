:: Compile script for LunarEclipse PDF report test modules

:: Set common C# compiler resources as variables
SET VarianRes=/r:VMS.TPS.Common.Model.API.dll /r:VMS.TPS.Common.Model.Types.dll
SET WpfRes=/r:WPF\PresentationCore.dll /r:WPF\PresentationFramework.dll^
 /r:WPF\WindowsBase.dll /r:System.Xaml.dll
SET PdfRes=/r:PdfSharp.dll /r:MigraDoc.DocumentObjectModel.dll^
 /r:MigraDoc.Rendering.dll^

:: Compile resources
csc.exe /target:exe /out:DoseEvalPrint.exe .\Source\ReportDoseEval.cs^
 .\Source\DoseEval.cs %VarianRes% %WpfRes% %PdfRes%
