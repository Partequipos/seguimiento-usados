# 4. Tecnologías y Dependencias

## Tecnologías Principales

### React 18
**¿Qué es?** Framework de JavaScript para construir interfaces de usuario.

**¿Para qué se usa?** Para crear todos los componentes visuales de la aplicación (botones, formularios, tablas, etc.)

**Versión:** 18.3.1

### TypeScript
**¿Qué es?** JavaScript con tipos. Ayuda a detectar errores antes de ejecutar el código.

**¿Para qué se usa?** Para escribir código más seguro y mantenible. Todo el código está en TypeScript.

**Versión:** 5.5.3

### Vite
**¿Qué es?** Herramienta para desarrollar y compilar aplicaciones web modernas.

**¿Para qué se usa?** 
- Servidor de desarrollo rápido
- Compilar el código para producción
- Optimizar los archivos finales

**Versión:** 5.4.2

### Tailwind CSS
**¿Qué es?** Framework de CSS que permite estilizar usando clases en el HTML.

**¿Para qué se usa?** Para dar estilo a la aplicación (colores, espaciados, diseño responsive)

**Versión:** 3.4.1

## Librerías de Autenticación y SharePoint

### @azure/msal-browser
**¿Qué es?** Librería oficial de Microsoft para autenticación en el navegador.

**¿Para qué se usa?** Para manejar el proceso de login con Microsoft Azure AD.

**Versión:** 3.27.0

### axios
**¿Qué es?** Cliente HTTP para hacer peticiones a APIs.

**¿Para qué se usa?** Para comunicarse con Microsoft Graph API y obtener/enviar datos de SharePoint.

**Versión:** 1.7.9

## Librerías de Interfaz de Usuario

### lucide-react
**¿Qué es?** Colección de iconos modernos.

**¿Para qué se usa?** Para los iconos en botones, menús, etc.

**Versión:** 0.344.0

### recharts
**¿Qué es?** Librería para crear gráficos y visualizaciones de datos.

**¿Para qué se usa?** Para los gráficos en el dashboard (barras, líneas, etc.)

**Versión:** 3.2.1

### react-hook-form
**¿Qué es?** Librería para manejar formularios de forma eficiente.

**¿Para qué se usa?** Para los formularios de agregar/editar equipos.

**Versión:** 7.64.0

### date-fns
**¿Qué es?** Librería para trabajar con fechas.

**¿Para qué se usa?** Para formatear y manipular fechas en la aplicación.

**Versión:** 4.1.0

## Herramientas de Desarrollo

### ESLint
**¿Qué es?** Herramienta para detectar errores y problemas en el código.

**¿Para qué se usa?** Para mantener el código limpio y sin errores.

**Versión:** 9.9.1

### TypeScript ESLint
**¿Qué es?** Extensión de ESLint para TypeScript.

**¿Para qué se usa?** Para detectar errores específicos de TypeScript.

**Versión:** 8.3.0

### PostCSS y Autoprefixer
**¿Qué es?** Herramientas para procesar CSS.

**¿Para qué se usa?** Para que los estilos funcionen en todos los navegadores.

## Resumen de Versiones

| Tecnología | Versión | Tipo |
|------------|---------|------|
| React | 18.3.1 | Dependencia principal |
| TypeScript | 5.5.3 | Dependencia de desarrollo |
| Vite | 5.4.2 | Dependencia de desarrollo |
| Tailwind CSS | 3.4.1 | Dependencia de desarrollo |
| MSAL Browser | 3.27.0 | Dependencia principal |
| Axios | 1.7.9 | Dependencia principal |
| Recharts | 3.2.1 | Dependencia principal |
| React Hook Form | 7.64.0 | Dependencia principal |
| Lucide React | 0.344.0 | Dependencia principal |
| date-fns | 4.1.0 | Dependencia principal |

## ¿Por qué estas tecnologías?

### React
- ✅ Muy popular y con gran comunidad
- ✅ Componentes reutilizables
- ✅ Buen rendimiento
- ✅ Fácil de mantener

### TypeScript
- ✅ Detecta errores antes de ejecutar
- ✅ Mejor autocompletado en editores
- ✅ Código más fácil de entender
- ✅ Refactorización más segura

### Vite
- ✅ Muy rápido en desarrollo
- ✅ Compilación rápida
- ✅ Configuración simple
- ✅ Optimizaciones automáticas

### Tailwind CSS
- ✅ Estilos rápidos sin escribir CSS personalizado
- ✅ Diseño responsive fácil
- ✅ Archivos finales pequeños (solo usa lo necesario)
- ✅ Consistencia visual

### MSAL
- ✅ Librería oficial de Microsoft
- ✅ Maneja tokens automáticamente
- ✅ Seguro y confiable
- ✅ Renovación automática de tokens

## Compatibilidad de Navegadores

La aplicación funciona en:
- ✅ Chrome (últimas 2 versiones)
- ✅ Edge (últimas 2 versiones)
- ✅ Firefox (últimas 2 versiones)
- ✅ Safari (últimas 2 versiones)

**No funciona en:**
- ❌ Internet Explorer (ya no se soporta)

## Requisitos del Sistema

### Para Desarrollo:
- Node.js 18 o superior
- npm 9 o superior
- 4 GB de RAM mínimo
- Conexión a internet

### Para Usuarios Finales:
- Navegador moderno (ver arriba)
- Conexión a internet
- Cuenta de Microsoft 365 con acceso a SharePoint

## Actualización de Dependencias

Para actualizar las dependencias:

```bash
# Ver qué paquetes tienen actualizaciones
npm outdated

# Actualizar todos los paquetes (cuidado: puede romper cosas)
npm update

# Actualizar un paquete específico
npm install nombre-del-paquete@version-nueva
```

⚠️ **Recomendación:** Después de actualizar, probar bien la aplicación para asegurarse de que todo funciona.

---

**Siguiente:** [Estructura del Código](./05-Estructura-Codigo.md)

