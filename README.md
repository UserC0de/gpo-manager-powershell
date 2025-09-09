# Gestor de GPOs en PowerShell

Gestor interactivo para administrar Group Policy Objects (GPO) en Windows: listar, buscar, respaldar, generar reportes de enlaces (HTML/XML) y eliminar con confirmación, siguiendo buenas prácticas de ejecución segura en PowerShell.

## ✨ Características

- Menú interactivo:
  - Listar todas las GPOs y sus GUIDs.
  - Buscar por nombre parcial.
  - Ver enlaces de una GPO (SOMPath) y abrir reporte HTML.
  - Ver resumen de enlaces de todas las GPOs.
  - Backup de una GPO o de todas (con reporte XML por GPO).
  - Eliminar GPO con confirmación y backup previo opcional.
  - Eliminar GPOs sin enlaces.

- Robustez:
  - `Set-StrictMode` y errores terminantes.
  - Saneado de nombres en rutas de backup.
  - Reporte XML junto al backup para trazabilidad de enlaces.

## 🧰 Requisitos

- Windows con RSAT/GPMC instalado (módulo `GroupPolicy`).
- Permisos adecuados sobre el dominio/OU/GPOs.
- PowerShell con permisos para ejecutar scripts locales.

## 🚀 Instalación

1. Clona o descarga este repositorio.
2. Abre PowerShell con una cuenta con permisos sobre GPOs.
3. (Opcional) Ajusta la ExecutionPolicy según la política de tu organización.

## ▶️ Uso

Ejecuta el script desde la carpeta del repositorio:

