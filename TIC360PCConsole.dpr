program TIC360PCConsole;

{$APPTYPE CONSOLE}

uses
  System.SysUtils,
  System.Variants,
  System.Win.ComObj,
  System.IniFiles,
  ActiveX,
  DateUtils,
  Windows;

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
    HWItem := WMIService.ExecQuery('SELECT Model, Manufacturer, SerialNumber, Size FROM Win32_DiskDrive');
    Enum := IUnknown(HWItem._NewEnum) as IEnumVARIANT;

    while Enum.Next(1, Value, Fetched) = 0 do
    begin
      Inc(Index);
      Log('Disk' + IntToStr(Index), 'Model', VarToStr(Value.Model));
      Log('Disk' + IntToStr(Index), 'Manufacturer', VarToStr(Value.Manufacturer));
      Log('Disk' + IntToStr(Index), 'Serial', VarToStr(Value.SerialNumber));
      Log('Disk' + IntToStr(Index), 'SizeGB', FormatFloat('0.00', StrToFloatDef(VarToStr(Value.Size), 0) / (1024 * 1024 * 1024)));
    end;
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

begin
  try
    CoInitialize(nil); // Inicializar COM
    try
      InitializeConfig;
      Log('General', 'Status', 'Iniciando PCMonitorConsole');

      GetComputerName;
      GetDiskDriveInfo;
      GetMemoryInfo;
      GetCPUInfo;
      GetNetworkAdapters;
      GetDisplayAdapters;
      GetInstalledCameras;
      GetAudioDevices;

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

