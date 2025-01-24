program TIC360PCInventario;
{
  PCInventario es un software elaborado por Eduardo Rizo
  Puede usarse libremente bajo los términos y condiciones
  de la licencia GPL versión 3.

  Modo de uso: Editar el archivo config.txt y cambia el valor de Area
  de acuerdo a tus necesidades. Al ejecutar el programa se creará una carpeta
  que se llamará igual que el valor de Area y pondrá dentro un archivo con
  la información recolectada. El nombre del archivo será igual al nombre
  de la computadora y le agregará la extensión 'txt'.

  Ejemplo del contenido del archivo config.txt

  [Settings]
  Area=CASA
}
{$APPTYPE CONSOLE}

uses
  System.SysUtils,
  System.Variants,
  System.Win.ComObj,
  System.IniFiles,
  ActiveX,
  DateUtils,
  Windows,
  System.StrUtils;

var
  OutputFile: TIniFile;
  OutputDir: string;
  ComputerName: string;

procedure InitializeConfig;
var
  Ini: TIniFile;
  Area: string;
  ComputerNameBuffer: array[0..MAX_COMPUTERNAME_LENGTH] of Char;
  Size: DWORD;
  OutputFilePath: string;
begin
  // Leer el archivo config.ini
  if not FileExists(ExtractFilePath(ParamStr(0)) + 'config.txt') then
  begin
    WriteLn('Error: No se encontró el archivo config.txt');
    Halt(1);
  end;

  Ini := TIniFile.Create(ExtractFilePath(ParamStr(0)) + 'config.txt');
  try
    Area := Ini.ReadString('Settings', 'Area', '');
    if Area = '' then
    begin
      WriteLn('Error: La variable AREA no está definida en config.txt');
      Halt(1);
    end;

    // Crear el directorio basado en AREA
    OutputDir := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)) + Area);
    if not DirectoryExists(OutputDir) then
      if not CreateDir(OutputDir) then
      begin
        WriteLn('Error: No se pudo crear el directorio ', OutputDir);
        Halt(1);
      end;

    // Obtener el nombre del equipo para el archivo de salida
    Size := MAX_COMPUTERNAME_LENGTH + 1;
    if Windows.GetComputerName(@ComputerNameBuffer, Size) then
    begin
      ComputerName := StrPas(ComputerNameBuffer);
      OutputFilePath := OutputDir + ComputerName + '.txt';

      // Si el archivo ya existe, eliminarlo
      if FileExists(OutputFilePath) then
        if not DeleteFile(PWideChar(OutputFilePath)) then
        begin
          WriteLn('Error: No se pudo eliminar el archivo existente ', OutputFilePath);
          Halt(1);
        end;

      // Crear el archivo nuevo
      OutputFile := TIniFile.Create(OutputFilePath);
    end
    else
    begin
      WriteLn('Error al obtener el nombre del equipo.');
      Halt(1);
    end;
  finally
    Ini.Free;
  end;
end;

procedure Log(const Section, Key, Value: string);
begin
  OutputFile.WriteString(Section, Key, Value);
end;

procedure GetComputerName;
var
  ComputerNameBuffer: array[0..MAX_COMPUTERNAME_LENGTH] of Char;
  Size: DWORD;
begin
  Size := MAX_COMPUTERNAME_LENGTH + 1;
  if Windows.GetComputerName(@ComputerNameBuffer, Size) then
    Log('General', 'ComputerName', StrPas(ComputerNameBuffer))
  else
    Log('General', 'Error', 'No se pudo obtener el nombre del equipo');
end;

procedure GetDiskDriveInfo;
var
  Locator, WMIService, DiskItem, LogicalItem: OLEVariant;
  EnumDisk, EnumLogical: IEnumVARIANT;
  DiskValue, LogicalValue: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
  SizeGB, FreeSpaceGB, UsedSpaceGB, PercentUsed: Double;
  MediaType: string;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');

    // Obtener información básica del disco
    DiskItem := WMIService.ExecQuery('SELECT Model, Manufacturer, SerialNumber, Size, MediaType FROM Win32_DiskDrive');
    EnumDisk := IUnknown(DiskItem._NewEnum) as IEnumVARIANT;

    while EnumDisk.Next(1, DiskValue, Fetched) = 0 do
    begin
      Inc(Index);
      SizeGB := StrToFloatDef(VarToStr(DiskValue.Size), 0) / (1024 * 1024 * 1024);
      MediaType := VarToStr(DiskValue.MediaType);

      Log('Disk' + IntToStr(Index), 'Model', VarToStr(DiskValue.Model));
      Log('Disk' + IntToStr(Index), 'Manufacturer', VarToStr(DiskValue.Manufacturer));
      Log('Disk' + IntToStr(Index), 'Serial', VarToStr(DiskValue.SerialNumber));
      Log('Disk' + IntToStr(Index), 'SizeGB', FormatFloat('0.00', SizeGB));

      // Verificar si el disco es físico o virtual
      if MediaType = '' then
        MediaType := 'Unknown'
      else if Pos('Virtual', MediaType) > 0 then
        MediaType := 'Virtual Disk'
      else
        MediaType := 'Physical Disk';

      Log('Disk' + IntToStr(Index), 'MediaType', MediaType);

      // Obtener espacio usado de las particiones asociadas al disco
      LogicalItem := WMIService.ExecQuery('SELECT Size, FreeSpace FROM Win32_LogicalDisk WHERE DriveType = 3');
      EnumLogical := IUnknown(LogicalItem._NewEnum) as IEnumVARIANT;

      while EnumLogical.Next(1, LogicalValue, Fetched) = 0 do
      begin
        FreeSpaceGB := StrToFloatDef(VarToStr(LogicalValue.FreeSpace), 0) / (1024 * 1024 * 1024);
        UsedSpaceGB := SizeGB - FreeSpaceGB;

        if SizeGB > 0 then
          PercentUsed := (UsedSpaceGB / SizeGB) * 100
        else
          PercentUsed := 0;

        Log('Disk' + IntToStr(Index), 'FreeSpaceGB', FormatFloat('0.00', FreeSpaceGB));
        Log('Disk' + IntToStr(Index), 'PercentUsed', FormatFloat('0.00', PercentUsed));
      end;
    end;

    if Index = 0 then
      Log('Disks', 'Status', 'No disks found');
  except
    on E: Exception do
      Log('Disks', 'Error', 'Error al obtener información de discos: ' + E.Message);
  end;
end;


procedure GetMemoryInfo;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Manufacturer, PartNumber, SerialNumber, Capacity FROM Win32_PhysicalMemory');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('Memory' + IntToStr(Index), 'Manufacturer', VarToStr(Value.Manufacturer));
      Log('Memory' + IntToStr(Index), 'PartNumber', VarToStr(Value.PartNumber));
      Log('Memory' + IntToStr(Index), 'Serial', VarToStr(Value.SerialNumber));
      Log('Memory' + IntToStr(Index), 'CapacityGB', FormatFloat('0.00', StrToFloatDef(VarToStr(Value.Capacity), 0) / (1024 * 1024 * 1024)));
    end;
  except
    on E: Exception do
      Log('Memory', 'Error', 'Error al obtener información de memoria RAM: ' + E.Message);
  end;
end;

procedure GetCPUInfo;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Name, Manufacturer, ProcessorId, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('CPU' + IntToStr(Index), 'Name', VarToStr(Value.Name));
      Log('CPU' + IntToStr(Index), 'Manufacturer', VarToStr(Value.Manufacturer));
      Log('CPU' + IntToStr(Index), 'ProcessorId', VarToStr(Value.ProcessorId));
      Log('CPU' + IntToStr(Index), 'MaxClockSpeedGHz', FormatFloat('0.00', Value.MaxClockSpeed / 1000));
      Log('CPU' + IntToStr(Index), 'Cores', VarToStr(Value.NumberOfCores));
      Log('CPU' + IntToStr(Index), 'Threads', VarToStr(Value.NumberOfLogicalProcessors));
    end;
  except
    on E: Exception do
      Log('CPU', 'Error', 'Error al obtener información del procesador: ' + E.Message);
  end;
end;

procedure GetNetworkAdapters;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Name, MACAddress, Manufacturer, NetEnabled FROM Win32_NetworkAdapter');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('NetworkAdapter' + IntToStr(Index), 'Name', VarToStr(Value.Name));
      Log('NetworkAdapter' + IntToStr(Index), 'MACAddress', VarToStr(Value.MACAddress));
      Log('NetworkAdapter' + IntToStr(Index), 'Manufacturer', VarToStr(Value.Manufacturer));
      if VarIsNull(Value.NetEnabled) then
        Log('NetworkAdapter' + IntToStr(Index), 'State', 'Unknown')
      else if Value.NetEnabled then
        Log('NetworkAdapter' + IntToStr(Index), 'State', 'Enabled')
      else
        Log('NetworkAdapter' + IntToStr(Index), 'State', 'Disabled');
    end;
  except
    on E: Exception do
      Log('NetworkAdapters', 'Error', 'Error al obtener información de las tarjetas de red: ' + E.Message);
  end;
end;

procedure GetDisplayAdapters;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Name, AdapterRAM, DriverVersion FROM Win32_VideoController');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('DisplayAdapter' + IntToStr(Index), 'Name', VarToStr(Value.Name));
      Log('DisplayAdapter' + IntToStr(Index), 'AdapterRAMBytes', VarToStr(Value.AdapterRAM));
      Log('DisplayAdapter' + IntToStr(Index), 'DriverVersion', VarToStr(Value.DriverVersion));
    end;
  except
    on E: Exception do
      Log('DisplayAdapters', 'Error', 'Error al obtener información de los adaptadores de pantalla: ' + E.Message);
  end;
end;

procedure GetInstalledCameras;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Name, Status, DeviceID FROM Win32_PnPEntity WHERE Description LIKE "%Camera%"');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('Camera' + IntToStr(Index), 'Name', VarToStr(Value.Name));
      Log('Camera' + IntToStr(Index), 'Status', VarToStr(Value.Status));
      Log('Camera' + IntToStr(Index), 'DeviceID', VarToStr(Value.DeviceID));
    end;
  except
    on E: Exception do
      Log('Cameras', 'Error', 'Error al obtener información de las cámaras: ' + E.Message);
  end;
end;

procedure GetAudioDevices;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Name, Status, DeviceID FROM Win32_SoundDevice');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('AudioDevice' + IntToStr(Index), 'Name', VarToStr(Value.Name));
      Log('AudioDevice' + IntToStr(Index), 'Status', VarToStr(Value.Status));
      Log('AudioDevice' + IntToStr(Index), 'DeviceID', VarToStr(Value.DeviceID));
    end;
  except
    on E: Exception do
      Log('AudioDevices', 'Error', 'Error al obtener información de los dispositivos de audio: ' + E.Message);
  end;
end;

procedure GetOSInfo;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
begin
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Caption, Version, BuildNumber, OSArchitecture FROM Win32_OperatingSystem');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Log('OperatingSystem', 'Caption', VarToStr(Value.Caption));
      Log('OperatingSystem', 'Version', VarToStr(Value.Version));
      Log('OperatingSystem', 'BuildNumber', VarToStr(Value.BuildNumber));
      Log('OperatingSystem', 'Architecture', VarToStr(Value.OSArchitecture));
    end;
  except
    on E: Exception do
      Log('OperatingSystem', 'Error', 'Error al obtener información del sistema operativo: ' + E.Message);
  end;
end;

procedure GetInstalledSoftware;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Name, Version FROM Win32_Product');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('Software' + IntToStr(Index), 'Name', VarToStr(Value.Name));
      Log('Software' + IntToStr(Index), 'Version', VarToStr(Value.Version));
    end;
  except
    on E: Exception do
      Log('InstalledSoftware', 'Error', 'Error al obtener información del software instalado: ' + E.Message);
  end;
end;

procedure GetWindowsLicenseInfo;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  LicenseStatus: Integer;
  Namespace: string;
begin
  Namespace := 'root\CIMV2'; // Intento inicial
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    try
      // Intentar primer espacio de nombres
      WMIService := Locator.ConnectServer('.', Namespace);
    except
      on E: Exception do
      begin
        Namespace := 'root\SoftwareLicensingService'; // Cambiar espacio de nombres
        WMIService := Locator.ConnectServer('.', Namespace);
      end;
    end;

    HWItem := WMIService.ExecQuery(
      'SELECT LicenseStatus, Description FROM SoftwareLicensingProduct ' +
      'WHERE PartialProductKey IS NOT NULL AND ApplicationID="55c92734-d682-4d71-983e-d6ec3f16059f"'
    );

    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    if Enum.Next(1, Value, Fetched) = 0 then
    begin
      LicenseStatus := Value.LicenseStatus;

      case LicenseStatus of
        0: Log('WindowsLicense', 'Status', 'Unlicensed');
        1: Log('WindowsLicense', 'Status', 'Licensed');
        2: Log('WindowsLicense', 'Status', 'Out-of-Box Grace Period');
        3: Log('WindowsLicense', 'Status', 'Out-of-Tolerance Grace Period');
        4: Log('WindowsLicense', 'Status', 'Non-Genuine Grace Period');
        5: Log('WindowsLicense', 'Status', 'Notification Mode');
        6: Log('WindowsLicense', 'Status', 'Extended Grace Period');
      else
        Log('WindowsLicense', 'Status', 'Unknown');
      end;

      // Registrar también la descripción del producto (opcional)
      Log('WindowsLicense', 'Description', VarToStr(Value.Description));
    end
    else
      Log('WindowsLicense', 'Status', 'No license information found');
  except
    on E: Exception do
      Log('WindowsLicense', 'Error', 'Error al obtener información de la licencia de Windows: ' + E.Message);
  end;
end;


procedure GetSecurityStatus;
var
  FirewallMgr: OleVariant;
  IsEnabled: Boolean;
  AntivirusFound: Boolean;
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
begin
  AntivirusFound := False;

  try
    // Verificar Antivirus
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\SecurityCenter2');
    HWItem := WMIService.ExecQuery('SELECT displayName FROM AntiVirusProduct');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    if Enum.Next(1, Value, Fetched) = 0 then
    begin
      AntivirusFound := True;
      Log('Security', 'Antivirus', VarToStr(Value.displayName));
    end
    else
      Log('Security', 'Antivirus', 'No antivirus detected');

    // Verificar Firewall usando HNetCfg.FwMgr
    try
      FirewallMgr := CreateOleObject('HNetCfg.FwMgr');
      IsEnabled := FirewallMgr.LocalPolicy.CurrentProfile.FirewallEnabled;
      if IsEnabled then
        Log('Security', 'Firewall', 'Enabled')
      else
        Log('Security', 'Firewall', 'Disabled');
    except
      on E: Exception do
        Log('Security', 'FirewallError', 'Error al verificar el Firewall: ' + E.Message);
    end;

  except
    on E: Exception do
    begin
      if not AntivirusFound then
        Log('Security', 'AntivirusError', 'Error al verificar antivirus: ' + E.Message);
      Log('Security', 'FirewallError', 'Error al verificar Firewall: ' + E.Message);
    end;
  end;
end;

procedure GetBIOSInfo;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
begin
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Manufacturer, Version, SerialNumber FROM Win32_BIOS');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    if Enum.Next(1, Value, Fetched) = 0 then
    begin
      Log('BIOS', 'Manufacturer', VarToStr(Value.Manufacturer));
      Log('BIOS', 'Version', VarToStr(Value.Version));
      Log('BIOS', 'SerialNumber', VarToStr(Value.SerialNumber));
    end
    else
      Log('BIOS', 'Error', 'No se pudo obtener información del BIOS');
  except
    on E: Exception do
      Log('BIOS', 'Error', 'Error al obtener información del BIOS: ' + E.Message);
  end;
end;

procedure GetWindowsUpdatePolicies;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT HotFixID, InstalledOn FROM Win32_QuickFixEngineering');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('WindowsUpdates' + IntToStr(Index), 'HotFixID', VarToStr(Value.HotFixID));
      Log('WindowsUpdates' + IntToStr(Index), 'InstalledOn', VarToStr(Value.InstalledOn));
    end;

    if Index = 0 then
      Log('WindowsUpdates', 'Status', 'No updates found');
  except
    on E: Exception do
      Log('WindowsUpdates', 'Error', 'Error al obtener información de actualizaciones: ' + E.Message);
  end;
end;

procedure GetNetworkConfiguration;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
  IPAddresses, Gateways, DNS: string;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Description, MACAddress, IPAddress, DefaultIPGateway, DNSServerSearchOrder, DHCPEnabled FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      IPAddresses := VarToStrDef(Value.IPAddress[0], 'N/A');
      Gateways := VarToStrDef(Value.DefaultIPGateway[0], 'N/A');
      DNS := VarToStrDef(Value.DNSServerSearchOrder[0], 'N/A');

      Log('NetworkConfig' + IntToStr(Index), 'Description', VarToStr(Value.Description));
      Log('NetworkConfig' + IntToStr(Index), 'MACAddress', VarToStr(Value.MACAddress));
      Log('NetworkConfig' + IntToStr(Index), 'IPAddress', IPAddresses);
      Log('NetworkConfig' + IntToStr(Index), 'Gateway', Gateways);
      Log('NetworkConfig' + IntToStr(Index), 'DNS', DNS);
      Log('NetworkConfig' + IntToStr(Index), 'DHCPEnabled', IfThen(Value.DHCPEnabled, 'Yes', 'No'));
    end;

    if Index = 0 then
      Log('NetworkConfig', 'Status', 'No active network adapters found');
  except
    on E: Exception do
      Log('NetworkConfig', 'Error', 'Error al obtener información de red: ' + E.Message);
  end;
end;


procedure GetLocalUsersAndGroups;
var
  Locator, WMIService, Users, Groups: OLEVariant;
  EnumUsers, EnumGroups: IEnumVARIANT;
  UserValue, GroupValue: OLEVariant;
  Fetched: LongWord;
  UserIndex, GroupIndex: Integer;
begin
  UserIndex := 0;
  GroupIndex := 0;

  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');

    // Obtener información de usuarios locales
    Users := WMIService.ExecQuery('SELECT Name, Status, Disabled FROM Win32_UserAccount WHERE LocalAccount = TRUE');
    EnumUsers := IUnknown(Users._NewEnum) as IEnumVARIANT;

    while EnumUsers.Next(1, UserValue, Fetched) = 0 do
    begin
      Inc(UserIndex);
      Log('LocalUser' + IntToStr(UserIndex), 'Name', VarToStr(UserValue.Name));
      Log('LocalUser' + IntToStr(UserIndex), 'Status', VarToStr(UserValue.Status));
      if UserValue.Disabled then
        Log('LocalUser' + IntToStr(UserIndex), 'Disabled', 'Yes')
      else
        Log('LocalUser' + IntToStr(UserIndex), 'Disabled', 'No');
    end;

    if UserIndex = 0 then
      Log('LocalUsers', 'Status', 'No local users found');

    // Obtener información de grupos locales
    Groups := WMIService.ExecQuery('SELECT Name, Caption FROM Win32_Group WHERE LocalAccount = TRUE');
    EnumGroups := IUnknown(Groups._NewEnum) as IEnumVARIANT;

    while EnumGroups.Next(1, GroupValue, Fetched) = 0 do
    begin
      Inc(GroupIndex);
      Log('LocalGroup' + IntToStr(GroupIndex), 'Name', VarToStr(GroupValue.Name));
      Log('LocalGroup' + IntToStr(GroupIndex), 'Caption', VarToStr(GroupValue.Caption));
    end;

    if GroupIndex = 0 then
      Log('LocalGroups', 'Status', 'No local groups found');
  except
    on E: Exception do
    begin
      Log('LocalUsers', 'Error', 'Error al obtener información de usuarios: ' + E.Message);
      Log('LocalGroups', 'Error', 'Error al obtener información de grupos: ' + E.Message);
    end;
  end;
end;

procedure GetComponentTemperatures;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
  TemperatureC: Double;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\\WMI');
    HWItem := WMIService.ExecQuery('SELECT CurrentTemperature, InstanceName FROM MSAcpi_ThermalZoneTemperature');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);

      // Convertir la temperatura de Kelvin a Celsius
      TemperatureC := (Value.CurrentTemperature - 2732) / 10.0;

      Log('Temperature' + IntToStr(Index), 'Zone', VarToStr(Value.InstanceName));
      Log('Temperature' + IntToStr(Index), 'TemperatureC', FormatFloat('0.0', TemperatureC));
    end;

    if Index = 0 then
      Log('Temperatures', 'Status', 'No temperature sensors found');
  except
    on E: Exception do
      Log('Temperatures', 'Error', 'Error al obtener información de temperaturas: ' + E.Message);
  end;
end;

procedure GetMotherboardDetails;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
begin
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Manufacturer, Product, SerialNumber, Version FROM Win32_BaseBoard');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    if Enum.Next(1, Value, Fetched) = 0 then
    begin
      Log('Motherboard', 'Manufacturer', VarToStr(Value.Manufacturer));
      Log('Motherboard', 'Model', VarToStr(Value.Product));
      Log('Motherboard', 'SerialNumber', VarToStr(Value.SerialNumber));
      Log('Motherboard', 'Version', VarToStr(Value.Version));
    end
    else
      Log('Motherboard', 'Status', 'No motherboard information found');
  except
    on E: Exception do
      Log('Motherboard', 'Error', 'Error al obtener información de la tarjeta madre: ' + E.Message);
  end;
end;

procedure GetMemoryUsage;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  TotalMemoryGB, FreeMemoryGB, UsedMemoryGB, PercentUsed: Double;
begin
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    if Enum.Next(1, Value, Fetched) = 0 then
    begin
      // Convertir valores de kilobytes a gigabytes
      TotalMemoryGB := StrToFloatDef(VarToStr(Value.TotalVisibleMemorySize), 0) / (1024 * 1024);
      FreeMemoryGB := StrToFloatDef(VarToStr(Value.FreePhysicalMemory), 0) / (1024 * 1024);
      UsedMemoryGB := TotalMemoryGB - FreeMemoryGB;

      // Calcular el porcentaje de memoria usada
      if TotalMemoryGB > 0 then
        PercentUsed := (UsedMemoryGB / TotalMemoryGB) * 100
      else
        PercentUsed := 0;

      Log('MemoryUsage', 'TotalMemoryGB', FormatFloat('0.00', TotalMemoryGB));
      Log('MemoryUsage', 'FreeMemoryGB', FormatFloat('0.00', FreeMemoryGB));
      Log('MemoryUsage', 'UsedMemoryGB', FormatFloat('0.00', UsedMemoryGB));
      Log('MemoryUsage', 'PercentUsed', FormatFloat('0.00', PercentUsed));
    end
    else
      Log('MemoryUsage', 'Status', 'No memory information found');
  except
    on E: Exception do
      Log('MemoryUsage', 'Error', 'Error al obtener información de memoria: ' + E.Message);
  end;
end;

procedure GetCPUUsage;
var
  Locator, WMIService, HWItem: OLEVariant;
  Enum: IEnumVARIANT;
  Value: OLEVariant;
  Fetched: LongWord;
  Index: Integer;
  CPUUsage: Double;
begin
  Index := 0;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := Locator.ConnectServer('.', 'root\CIMV2');
    HWItem := WMIService.ExecQuery('SELECT Name, PercentProcessorTime FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name = "_Total"');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      CPUUsage := StrToFloatDef(VarToStr(Value.PercentProcessorTime), 0);

      Log('CPUUsage' + IntToStr(Index), 'Processor', VarToStr(Value.Name));
      Log('CPUUsage' + IntToStr(Index), 'PercentUsed', FormatFloat('0.00', CPUUsage));
    end;

    if Index = 0 then
      Log('CPUUsage', 'Status', 'No CPU usage information found');
  except
    on E: Exception do
      Log('CPUUsage', 'Error', 'Error al obtener información de uso del CPU: ' + E.Message);
  end;
end;

procedure GetCurrentDateTime;
var
  CurrentDate: string;
  CurrentTime: string;
begin
  try
    // Obtener la fecha y hora actuales
    CurrentDate := FormatDateTime('yyyy-mm-dd', Now); // Solo la fecha
    CurrentTime := FormatDateTime('hh:nn:ss', Now);   // Solo la hora

    // Registrar la fecha y la hora por separado
    Log('DateTime', 'CurrentDate', CurrentDate);
    Log('DateTime', 'CurrentTime', CurrentTime);
  except
    on E: Exception do
      Log('DateTime', 'Error', 'Error al obtener la fecha/hora: ' + E.Message);
  end;
end;

begin
  try
    CoInitialize(nil); // Inicializar COM
    try
      InitializeConfig;
      WriteLn('Iniciando TIC360 PCInventario... Por favor espere');

      GetComputerName;
      GetCurrentDateTime;
      GetMotherboardDetails;
      GetBIOSInfo;
      GetDiskDriveInfo;
      GetMemoryInfo;
      GetMemoryUsage;
      GetCPUInfo;
      GetCPUUsage;
      GetNetworkAdapters;
      GetNetworkConfiguration;
      GetDisplayAdapters;
      GetInstalledCameras;
      GetAudioDevices;
      //GetComponentTemperatures;
      GetOSInfo;
      GetWindowsUpdatePolicies;
      //GetInstalledSoftware;
      //GetLocalUsersAndGroups;
      GetWindowsLicenseInfo;
      GetSecurityStatus;

      Log('General', 'Status', 'Proceso completado');
    finally
      CoUninitialize; // Finalizar COM
      OutputFile.Free;
    end;
  except
    on E: Exception do
    begin
      WriteLn('Error inesperado: ', E.Message);
      if Assigned(OutputFile) then
        OutputFile.Free;
    end;
  end;
end.


