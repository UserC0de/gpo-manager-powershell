Gestor de GPOs en PowerShell
Gestor interactivo para administrar Group Policy Objects (GPO) en entornos Windows: listar, buscar, respaldar, generar reportes de enlaces (HTML/XML) y eliminar con confirmación y buenas prácticas de ejecución segura en PowerShell.

Características
Menú interactivo: listar todas las GPOs, buscar por nombre parcial y ver enlaces activos por GPO o globalmente.

Backups: por GPO o de todas, con carpeta por ejecución y reporte XML adjunto para trazabilidad de enlaces/SOM.

Eliminación segura: confirmación explícita, opción de backup previo y uso de -Confirm al borrar.

Robustez: Set-StrictMode y errores terminantes para detectar fallos tempranamente y endurecer el script.

Requisitos
Windows con RSAT/GPMC instalado (módulo GroupPolicy) y permisos sobre AD/GPOs.

PowerShell con capacidad para ejecutar scripts locales según tu política de ejecución.

Instalación
Clonar o descargar este repositorio desde GitHub a una carpeta local de trabajo.

Abrir PowerShell con un usuario que tenga permisos suficientes sobre las GPOs del dominio.

Uso
Ejecutar el script en la carpeta del repositorio: .\Gestor-GPO.ps1.

Opciones del menú:

Ver todas las GPOs y sus GUIDs.

Ver enlaces de una GPO: muestra SOMPath en consola y genera un HTML en el escritorio.

Hacer backup de una GPO o de todas, guardando también un reporte XML.

Eliminar GPO (con confirmación y backup opcional) o eliminar todas las GPOs sin enlaces.

Consejo: los informes HTML/XML no deben subirse al repositorio; mantén los backups fuera de control de versiones.

Seguridad y buenas prácticas
Activa Set-StrictMode y usa errores terminantes para evitar estados silenciosos y variables no inicializadas.

No subas artefactos de entorno o backups reales; usa .gitignore para excluir C:\BackupGPOs/ y reportes.

Protege credenciales de GitHub con 2FA y PAT si automatizas despliegues o pushes.
