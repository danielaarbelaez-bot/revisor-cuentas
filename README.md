# Revisor web privado — versión inicial

Esta versión abre en navegador sin instalar nada en los computadores de los usuarios, **después de que un administrador la despliegue en un servidor**. No está publicada ni conectada a Outlook.

## Despliegue institucional (personal de TI)

1. Aprobar servidor y tratamiento de datos personales; restringir acceso mediante VPN o red institucional, HTTPS, políticas de retención y auditoría.
2. Construir imagen: `docker build -t revisor-cuentas .`
3. Configurar variables `REVISOR_USER` y `REVISOR_PASSWORD` en un gestor de secretos del servidor (nunca en código ni en repositorios). Para una producción real, sustituir el login de demostración por SSO institucional con MFA.
4. Ejecutar el contenedor detrás de proxy HTTPS, con controles de acceso, límites de carga, escaneo de archivos y aislamiento.
5. Ingresar desde navegador a la URL aprobada por TI. Cargar la base maestra manualmente en cada sesión de revisión.

## Alcance actual

- Lee PDF con texto, DOCX, XLSX/XLSM y ZIP con restricciones básicas.
- Compara campos seleccionados con la base maestra que cargue el usuario.
- Genera alertas preliminares y borrador de correo.
- No hay conexión con Outlook, IA externa ni publicación web.
- No confirma firmas, macros, anonimización total, valores de seguridad social ni correspondencia exhaustiva de evidencias.
- El control de acceso incluido es **solo demostrativo**: no tiene MFA, roles, bloqueo por intentos, registros de auditoría ni control de sesiones centralizado. No es apto para datos reales sensibles sin endurecimiento y aprobación institucional.
- No se incluye ninguna base de contratistas ni documentos reales dentro de este paquete.

## Próximas fases

1. SSO institucional (Microsoft Entra ID), MFA y permisos por rol.
2. Integración autorizada de Outlook mediante Microsoft Graph, con consentimiento y permisos mínimos.
3. Motor de validaciones configurables para F62, F68 V3, PILA, matrices y soportes.
4. Detección y revisión humana de datos sensibles, incluso en escaneos y fotografías.
5. Historial, trazabilidad, exportación de informes y controles de retención.
