Place your locally built SimConnect.NET.dll in this folder.

Expected filename: SimConnect.NET.dll
You can override the path by setting the MSBuild property `SimConnectNetDllPath`, e.g.:

  dotnet build -p:SimConnectNetDllPath=D:\path\to\SimConnect.NET.dll

If the DLL is not found, the project will fall back to the NuGet package reference.
