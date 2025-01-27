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
De manera predeterminada, están desactivadas las opciones relativas a: Configuración de Red, Usuarios y Grupos locales, así como el software instalado.

## [Procedures]
### GetComputerName=SI | NO
Nombre de la computadora.

### GetCurrentDateTime=SI | NO
Fecha y hora.

### GetMotherboardDetails=SI | NO
Fabricante, modelo, número de serie, versión.

### GetBIOSInfo=SI | NO
Fabricante, versión, número de serie.

### GetDiskDriveInfo=SI | NO
Fabricante, número de serie, modelo, tamaño en GB, tipo de medio (físico o virtual), espacio libre en GB, porcentaje usado.

### GetMemoryInfo=SI | NO
Fabricante, número de parte, número de serie, capacidad en GB

### GetMemoryUsage=SI | NO
Memoria total, memoria libre, memoria usada, porcentaje de memoria usada

### GetCPUInfo=SI | NO
Nombre, fabricante, Id del procesador, velocidad máxima, núcleos, hilos

### GetCPUUsage=SI | NO
Porcentaje de uso

### GetNetworkAdapters=SI | NO
Nombre, dirección MAC, fabricante, estado

### GetNetworkConfiguration=NO | SI
Descripción, dirección MAC, dirección IP, puerta de enlace, servidor DNS, DHCP habilitado

### GetDisplayAdapters=SI | NO
Nombre, cantidad de memoria RAM, versión del driver.

### GetInstalledCameras=SI | NO
Nombre, estatus, ID del dispositivo

### GetAudioDevices=SI | NO
Nombre, estatus, ID del dispositivo

### GetOSInfo=SI | NO
Descripcion, versión, número de compilación, arquitectura

### GetWindowsUpdatePolicies=SI | NO
Clave de actualización, fecha de actualización

### GetInstalledSoftware=NO | SI
Nombre, versión

### GetLocalUsersAndGroups=NO | SI
Nombre, estado, desactivado (si o no)

### GetWindowsLicenseInfo=SI | NO
Nombre, subtítulo

### GetSecurityStatus=SI | NO
Nombre del antivirus, estado del firewall
