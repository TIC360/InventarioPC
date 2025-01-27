# PCInventario
Software que recolecta información del hardware y software de computadoras con sistema operativo Microsoft Windows

Modo de uso: Editar el archivo config.txt y cambia el valor de Area
de acuerdo a tus preferencias. Al ejecutar el programa se creará una carpeta
con nombre igual que el valor de Area y pondrá dentro un archivo con
la información recolectada. El nombre del archivo será igual al nombre
de la computadora y le agregará la extensión 'txt'.

Ejemplo del contenido del archivo config.txt

[Settings]

Area=CASA

También hay un archivo llamado config.ini, en este archivo se define el tipo de
información que se desea recolectar, se activa con SI y se ignora con NO. 

[Procedures]
GetComputerName=SI

GetCurrentDateTime=SI

GetMotherboardDetails=SI

GetBIOSInfo=SI

GetDiskDriveInfo=SI

GetMemoryInfo=SI

GetMemoryUsage=SI

GetCPUInfo=SI

GetCPUUsage=SI

GetNetworkAdapters=SI

GetNetworkConfiguration=SI

GetDisplayAdapters=SI

GetInstalledCameras=SI

GetAudioDevices=SI

GetComponentTemperatures=SI

GetOSInfo=SI

GetWindowsUpdatePolicies=SI

GetInstalledSoftware=SI

GetLocalUsersAndGroups=SI

GetWindowsLicenseInfo=SI

GetSecurityStatus=SI
