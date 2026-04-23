[Setup]
AppName=Add-insSupport
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppVersion=1.0.2
AppPublisher=ToolSupport
DefaultDirName={autopf}\Add-insSupport
DefaultGroupName=Add-insSupport
OutputDir=.\installer_output
OutputBaseFilename=Add-insSupport_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
; Khong can quyen Admin, chi dung HKCU
; Tu dong dong Excel truoc khi cai de tranh file bi lock (DLL cu khong the ghi de)
CloseApplications=yes
CloseApplicationsFilter=excel.exe

[Files]
; Copy TOAN BO file tu bin\Release (bao gom cac dependency DLL Private=True
; vi du: Microsoft.Office.Tools.Common.v4.0.Utilities.dll, v.v.)
; Neu thieu bat ky DLL nao, add-in se load loi tham va ribbon khong hien.
Source: "bin\Release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs; Excludes: "*.pdb,*.ps1"

[InstallDelete]
; Xoa toan bo file cu trong thu muc cai dat truoc khi copy file moi
; → Tranh VSTO / CLR load DLL cache tu phien cu
Type: filesandordirs; Name: "{app}"

[Icons]
Name: "{group}\Uninstall Add-insSupport"; Filename: "{uninstallexe}"

[Messages]
FinishedLabel=Cai dat thanh cong. Vui long khoi dong lai Excel de ap dung add-in.

[Code]
// Dung Code section de tao file URI voi forward slash dung chuan
// va tro vao file .vsto (deployment manifest) voi |vstolocal

// Xoa cac entry VSTO SolutionMetadata co the dang cache manifest cu
procedure CleanVstoSolutionMetadata;
var
  MetaKey: String;
  SubKeys: TArrayOfString;
  I: Integer;
  ManifestVal: String;
begin
  MetaKey := 'Software\Microsoft\VSTO\SolutionMetadata';
  if not RegGetSubkeyNames(HKEY_CURRENT_USER, MetaKey, SubKeys) then Exit;
  for I := 0 to GetArrayLength(SubKeys) - 1 do
  begin
    if RegQueryStringValue(HKEY_CURRENT_USER,
        MetaKey + '\' + SubKeys[I], 'Manifest', ManifestVal) then
    begin
      if Pos('Add-insSupport', ManifestVal) > 0 then
        RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER,
            MetaKey + '\' + SubKeys[I]);
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  AppPath: String;
  ManifestUri: String;
  RegKey: String;
begin
  if CurStep = ssPreInstall then
  begin
    // Xoa VSTO metadata cache truoc khi cai moi de force load DLL moi
    CleanVstoSolutionMetadata;
  end;

  if CurStep = ssPostInstall then
  begin
    AppPath := ExpandConstant('{app}');
    // Chuyen backslash thanh forward slash cho file URI
    StringChangeEx(AppPath, '\', '/', True);
    // Tro vao .vsto (deployment manifest) + |vstolocal de bo qua kiem tra cert
    ManifestUri := 'file:///' + AppPath + '/Add-insSupport.vsto|vstolocal';
    RegKey := 'Software\Microsoft\Office\Excel\Addins\Add-insSupport';
    RegWriteStringValue(HKEY_CURRENT_USER, RegKey, 'Description',  'Add-insSupport Excel Add-in');
    RegWriteStringValue(HKEY_CURRENT_USER, RegKey, 'FriendlyName', 'Add-insSupport');
    RegWriteDWordValue (HKEY_CURRENT_USER, RegKey, 'LoadBehavior',  3);
    RegWriteStringValue(HKEY_CURRENT_USER, RegKey, 'Manifest',     ManifestUri);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER,
      'Software\Microsoft\Office\Excel\Addins\Add-insSupport');
    // Xoa VSTO metadata de Excel khong con nho add-in cu
    CleanVstoSolutionMetadata;
  end;
end;
