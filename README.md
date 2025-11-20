# Sistema RPA - Natural Conexión

Sistema de Automatización Robótica de Procesos (RPA) desarrollado para Natural Conexión, diseñado para automatizar el procesamiento, validación y registro de pedidos.

## Descripción

Esta aplicación web permite procesar archivos Excel/CSV con pedidos, validar automáticamente la información, simular el registro en el sistema SAG y generar reportes completos con estadísticas y dashboards visuales.

## Características Principales

### 1. Validación Automática de Pedidos
- Verificación de campos obligatorios
- Validación de formatos de correo electrónico
- Comprobación de valores numéricos
- Validación de cantidades y montos

### 2. Simulación de Registro en SAG
- Registro automático de pedidos validados
- Generación de estados y fechas de registro
- Simulación de fallos aleatorios para pruebas

### 3. Sistema de Notificaciones
- Log detallado de correos electrónicos simulados
- Registro de envíos exitosos y fallidos
- Trazabilidad completa de comunicaciones

### 4. Reportes y Dashboard
- Generación automática de archivos Excel con resultados
- Gráficos visuales (pie charts y bar charts)
- KPIs completos del procesamiento
- Dashboard web interactivo

## Estructura del Proyecto

```
project/
├── app.py                      # Aplicación Flask principal
├── rpa_engine.py              # Motor RPA con toda la lógica
├── requirements.txt           # Dependencias Python
├── templates/
│   ├── index.html            # Página de carga de archivos
│   └── resultado.html        # Dashboard de resultados
├── data/                      # Archivos subidos por usuarios
├── output/                    # Archivos generados por RPA
│   ├── PedidosValidados.xlsx
│   ├── ErroresRPA.xlsx
│   ├── PedidosRegistradosSAG.xlsx
│   ├── ReporteProcesados.xlsx
│   ├── DashboardRPA.xlsx
│   └── LogCorreos.txt
└── ejemplo_pedidos.xlsx       # Archivo de ejemplo para pruebas
```

## Instalación y Ejecución en Replit

El proyecto está configurado para ejecutarse automáticamente en Replit. Solo necesitas:

1. Hacer clic en el botón **Run** o presionar el botón de inicio
2. El servidor se iniciará automáticamente en el puerto 5000
3. Accede a la interfaz web desde el panel de vista previa

### Instalación Manual (si es necesario)

```bash
pip install -r requirements.txt
python app.py
```

## Uso del Sistema

### Paso 1: Preparar el Archivo

Tu archivo Excel/CSV debe contener las siguientes columnas:

- **ID_Pedido**: Identificador único del pedido
- **Fecha_Pedido**: Fecha de realización
- **Nombre_Cliente**: Nombre completo del cliente
- **Correo_Cliente**: Email válido
- **Direccion_Envio**: Dirección de entrega
- **SKU**: Código del producto
- **Nombre_Producto**: Nombre del producto
- **Cantidad**: Cantidad ordenada (numérico)
- **Precio_Unitario**: Precio por unidad (numérico)
- **Valor_Total**: Total del pedido (numérico)

### Paso 2: Cargar el Archivo

1. Accede a la página principal
2. Haz clic en el área de carga o arrastra tu archivo
3. Formatos permitidos: `.xlsx`, `.xls`, `.csv`
4. Haz clic en **Procesar Pedidos**

### Paso 3: Revisar Resultados

El sistema mostrará:
- Total de pedidos procesados
- Pedidos válidos e inválidos
- Tasa de validación
- Registros exitosos en SAG
- Correos enviados
- Tiempo de procesamiento

### Paso 4: Descargar Archivos

Descarga los archivos generados:
- **PedidosValidados.xlsx**: Pedidos que pasaron la validación
- **ErroresRPA.xlsx**: Pedidos con errores y sus motivos
- **PedidosRegistradosSAG.xlsx**: Pedidos registrados en SAG
- **ReporteProcesados.xlsx**: Reporte detallado
- **DashboardRPA.xlsx**: Dashboard con KPIs
- **LogCorreos.txt**: Log de notificaciones

## Archivo de Ejemplo

Se incluye `ejemplo_pedidos.xlsx` con datos de prueba que contiene:
- 8 pedidos de ejemplo
- Algunos registros válidos
- Algunos con errores intencionales (para demostrar validación)
- Diferentes casos de uso

## Funcionalidades del Motor RPA

### Validaciones Implementadas

1. **Campos obligatorios**: Verifica que no estén vacíos
2. **Email válido**: Formato correcto de correo electrónico
3. **Valores numéricos**: Cantidad y Valor_Total deben ser números
4. **Valores positivos**: Valor_Total debe ser mayor a 0

### Archivos Generados

| Archivo | Descripción |
|---------|-------------|
| PedidosValidados.xlsx | Pedidos que superaron todas las validaciones |
| ErroresRPA.xlsx | Detalle de errores encontrados |
| PedidosRegistradosSAG.xlsx | Pedidos con estado de registro SAG |
| ReporteProcesados.xlsx | Métricas y estadísticas del proceso |
| DashboardRPA.xlsx | Dashboard con KPIs y formato profesional |
| LogCorreos.txt | Registro de notificaciones enviadas |
| graficos_dashboard.png | Gráficos visuales (pie/bar charts) |

## Tecnologías Utilizadas

- **Backend**: Flask 3.0.0
- **Procesamiento de Datos**: Pandas 2.1.4
- **Archivos Excel**: OpenPyXL 3.1.2
- **Gráficos**: Matplotlib 3.8.2
- **Frontend**: HTML5, CSS3, Bootstrap 5, JavaScript
- **Servidor**: Werkzeug 3.0.1

## Características de Seguridad

- Validación de tipos de archivo
- Límite de tamaño de archivo (16 MB)
- Nombres de archivo seguros (secure_filename)
- Variables de entorno para secretos
- Validación de entrada de datos

## Próximas Mejoras

1. Integración real con sistema SAG mediante API
2. Envío real de correos electrónicos
3. Autenticación de usuarios
4. Historial de procesamiento
5. Exportación a múltiples formatos
6. Procesamiento en cola para archivos grandes

## Soporte

Para preguntas o problemas, contacta al equipo de desarrollo de Natural Conexión.

## Licencia

Desarrollado para Natural Conexión - 2025
