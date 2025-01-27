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

## [Procedures]
### GetComputerName=SI
Nombre de la computadora.

### GetCurrentDateTime=SI
Fecha y hora.

### GetMotherboardDetails=SI
Fabricante, modelo, número de serie, versión.

### GetBIOSInfo=SI
Fabricante, versión, número de serie.

### GetDiskDriveInfo=SI
Fabricante, número de serie, modelo, tamaño en GB, tipo de medio (físico o virtual), espacio libre en GB, porcentaje usado.

### GetMemoryInfo=SI
Fabricante, número de parte, número de serie, capacidad en GB

### GetMemoryUsage=SI
Memoria total, memoria libre, memoria usada, porcentaje de memoria usada

### GetCPUInfo=SI
Nombre, fabricante, Id del procesador, velocidad máxima, núcleos, hilos

### GetCPUUsage=SI
Porcentaje de uso

### GetNetworkAdapters=SI
Nombre, dirección MAC, fabricante, estado

### GetNetworkConfiguration=SI
Descripción, dirección MAC, dirección IP, puerta de enlace, servidor DNS, DHCP habilitado

### GetDisplayAdapters=SI
Nombre, cantidad de memoria RAM, versión del driver.

### GetInstalledCameras=SI
Nombre, estatus, ID del dispositivo

### GetAudioDevices=SI
Nombre, estatus, ID del dispositivo

### GetOSInfo=SI
Descripcion, versión, número de compilación, arquitectura

### GetWindowsUpdatePolicies=SI
Clave de actualización, fecha de actualización

### GetInstalledSoftware=SI
Nombre, versión

### GetLocalUsersAndGroups=SI
Nombre, estado, desactivado (si o no)

### GetWindowsLicenseInfo=SI
Nombre, subtítulo

### GetSecurityStatus=SI
Nombre del antivirus, estado del firewall
