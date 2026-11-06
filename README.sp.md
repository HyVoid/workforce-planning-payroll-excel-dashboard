[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Seguimiento de nómina y planificación de capacidad de personal Plantilla de Excel | Calculadora de horas extra, FTE y costo de mano de obra
<img width="1920" height="790" alt="image" src="https://github.com/user-attachments/assets/f36d608b-7a02-4a4b-b72a-ec7702d528f5" />


![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Browser%20%2B%20Microsoft%20Excel-success)
![Tool](https://img.shields.io/badge/Tool-Payroll%20%26%20Workforce%20Planning-blue)

**Esta plantilla de Excel para el seguimiento de nómina lista para usar y esta hoja de cálculo de planificación de capacidad de personal ayudan a los dueños de pequeñas empresas, a los generalistas de recursos humanos y a los gerentes de operaciones a calcular los costos mensuales de mano de obra, registrar las horas extra de los empleados y monitorear la capacidad de dotación en equivalentes a tiempo completo (FTE) en menos de 1 minuto. Construida con fórmulas de matriz dinámica automatizadas para Microsoft 365 y Excel 2021+, este registro de horas de empleados y calculadora de costo de mano de obra listo para usar no requiere configurar macros, arrastrar fórmulas ni tener conocimientos de análisis de datos: basta con copiar y pegar los datos brutos de tus hojas de horas para generar un panel ejecutivo de nómina y capacidad de personal.**

Prueba la calculadora web interactiva en tu navegador sin costo, o descarga la hoja de cálculo de Excel desbloqueada para tus ciclos mensuales de nómina recurrentes (con garantía de devolución del dinero de 30 días).

> 🌐 **Calculadora gratuita para el navegador** → [Abre la calculadora interactiva gratuita de nómina y planificación de capacidad de personal en tu navegador (aplicación web HTML, sin inicio de sesión)](https://hyvoid.github.io/workforce-planning-payroll-excel-dashboard/)
>
> 📥 **Plantilla de Excel lista para usar** → [Descarga la plantilla de Excel automatizada de seguimiento de nómina empresarial y planificación de capacidad de personal (.xlsx para Excel 2021+ / Microsoft 365)](https://www.theseusworkshop.com/l/ufscrp?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=payroll-workforce-planning)

---

## ¿Qué preguntas sobre costo de mano de obra, horas extra y capacidad de personal responde esta hoja de cálculo?

Procesar los salarios brutos mensuales y las deducciones fiscales solo te dice qué dinero ya salió de tu cuenta bancaria. Para evitar desviaciones presupuestarias y cuellos de botella en los turnos del próximo mes, los gerentes necesitan una **hoja de cálculo unificada de análisis de costo de mano de obra y planificación de plantilla** que conecte los gastos de nómina con las horas de trabajo restantes:

- **Capacidad FTE y ritmo de consumo de carga de trabajo:** ¿Qué departamentos están agotando sus **horas contratadas anuales en equivalentes a tiempo completo (FTE)** más rápido de lo presupuestado?
- **Fuga de pago por horas extra (multiplicadores de 1,5x y 2,0x):** ¿El aumento del **gasto mensual en horas extra** se debe a un pico estacional temporal o a una escasez permanente de personal?
- **Cumplimiento de horas de trabajo y prevención del agotamiento:** ¿Qué empleados concretos alcanzarán sus **límites anuales de cumplimiento de horas de trabajo** antes de la temporada alta del cuarto trimestre?
- **Desglose de la variación de la nómina bruta:** ¿El aumento de los **costos mensuales de mano de obra** se debe a subidas del salario base, a horas extra no aprobadas o a complementos por turnos?
- **Capacidad del proyecto y asignación de recursos:** ¿Puede tu equipo de operaciones actual asumir nuevos pedidos de clientes este mes con la **capacidad de personal no utilizada** existente, o necesitas contratar?
- **Previsión de cuellos de botella por departamento:** ¿Qué unidad de negocio se quedará sin horas de mano de obra disponibles primero durante los meses de máxima producción?

El software de nómina estándar exporta registros históricos en bruto. Este **panel de gestión de la fuerza laboral (WFM) en Excel** convierte esas filas brutas de hojas de horas en un **sistema automatizado de apoyo a la toma de decisiones operativas (DSS)**. Como cada cálculo se ejecuta con funciones nativas de matriz dinámica de Excel (`LET`, `XLOOKUP`, `SUMIFS`, `FILTER`) sin macros VBA, el archivo se abre sin problemas en cualquier equipo de escritorio moderno, sin bloqueos de seguridad ni scripts rotos.

---

## Qué te ayuda a controlar: correspondencia entre problemas de RR. HH. y operaciones y sus soluciones

Utiliza esta tabla de referencia estructurada para ver cómo los **problemas específicos de seguimiento de nómina, auditoría de horas extra y capacidad de personal** se corresponden directamente con hojas automatizadas dentro del libro de trabajo:

| Punto de dolor operativo de alta frecuencia (lo que te cuesta trabajo) | Qué rastrea esta plantilla de Excel (métricas y fórmulas automatizadas) | Solución de gestión inmediata (resultado de decisión en 1 minuto) |
| :--- | :--- | :--- |
| **Desviaciones presupuestarias mensuales imprevistas en horas extra** | Rastrea las horas regulares frente a las **horas extra de 1,5x y 2,0x**, las proporciones entre pago por horas extra y salario regular, y los picos de pagos premium por departamento. | **El panel de monitoreo de horas extra en tiempo real** aísla si el costo de las horas extra es una anomalía de una semana o una escasez crónica de plantilla que exige una nueva contratación. |
| **Agotamiento de los empleados e incumplimiento del límite anual de horas** | Compara las **horas reales trabajadas acumuladas en el año (YTD)** con las horas FTE contratadas anuales y con umbrales personalizados de alerta de utilización (por ejemplo, un tope del 85%). | **El rastreador de horas contratadas anuales** señala a los empleados de alto riesgo meses antes de que superen los límites legales laborales o sufran agotamiento durante las temporadas altas. |
| **Consolidación caótica de hojas de horas de varios departamentos** | Consolida el **salario base, los salarios por horas extra y los complementos** de hasta 7 pestañas departamentales en un único libro mayor maestro. | **El motor automatizado de resumen de nómina** elimina los errores de copiar y pegar manualmente y los enlaces `VLOOKUP` rotos entre las hojas de cálculo de los supervisores de turno. |
| **Improvisación al programar nuevos proyectos o turnos** | Calcula las **horas anuales disponibles restantes** por empleado y las **tasas netas de utilización de mano de obra (%)** por departamento. | **El pronosticador de capacidad de personal** muestra la capacidad libre exacta de cada equipo para que puedas asignar nuevas cargas de trabajo a los equipos subutilizados sin costo adicional de contratación. |
| **Incapacidad para explicar los picos mensuales de costo de mano de obra** | Separa los **costos fijos de salario base** de los **costos variables de comportamiento** (multiplicadores de horas extra, diferenciales de turno, bonificaciones). | **La vista de variación del costo de mano de obra por departamento** ofrece a los responsables de finanzas y operaciones un desglose instantáneo para las revisiones presupuestarias mensuales. |

---

## Tutorial de inicio rápido: cómo ejecutar tu análisis de nómina y capacidad en 1 minuto (4 pasos sencillos)

No necesitas habilidades avanzadas de Excel, formación en tablas dinámicas ni experiencia en modelado financiero para usar esta herramienta. Los cálculos de personal subyacentes son complejos, **pero todas las reglas de búsqueda, los multiplicadores de horas extra y las fórmulas de matriz dinámica ya están predefinidos y se expanden automáticamente**. Ni siquiera tienes que hacer doble clic en las esquinas de las celdas ni arrastrar fórmulas hacia abajo: pega tus datos y el panel se actualiza al instante.

### Paso 1: define una sola vez tus reglas globales de nómina y horas extra
Abre la pestaña `Settings` para introducir los parámetros de referencia de tu empresa:
- Mes de inicio del ejercicio fiscal y horizonte de planificación
- Multiplicadores estándar de pago de horas extra (por ejemplo, `1.5x` para horas extra estándar, `2.0x` para tarifa de festivos y fines de semana)
- Umbral de alerta de utilización de la capacidad de los empleados (por ejemplo, activar la advertencia al `85%` de las horas contratadas anuales)
- Símbolo de la moneda local y nombres de los departamentos

*Resultado:* cada gráfico del panel, cada alerta de cumplimiento y cada tabla de nómina se actualizan automáticamente en todo el libro de trabajo.

### Paso 2: copia y pega tu plantilla de empleados y las hojas de horas mensuales
Pega tu lista de personal en la pestaña `Employee_Master` y luego pega los registros mensuales de fichaje en las hojas departamentales:
- ID del empleado, nombre, departamento y horas contratadas anuales
- Horas de trabajo regulares registradas en el período de pago
- Horas extra aprobadas (`1.5x` / `2.0x`)
- Complementos o bonificaciones por turno pagados

*Resultado:* copia y pega directamente las exportaciones CSV en bruto de **ADP, Gusto, QuickBooks Payroll, Workday, BambooHR o relojes de fichaje manuales**. Las fórmulas dinámicas predefinidas relacionan automáticamente los ID de los empleados y se expanden para ajustarse a tu número de filas: no hace falta arrastrar ninguna fórmula.

### Paso 3: consulta tu panel instantáneo de costo de mano de obra y capacidad de personal
Haz clic en `Payroll_Summary`, `Annual_Hours_Tracker` o `Dashboard` para leer tu análisis completo:
- Gasto mensual total de nómina bruta (salario base frente a horas extra frente a complementos)
- Horas contratadas anuales restantes por empleado
- Porcentaje de utilización FTE por departamento y equilibrio de la carga de trabajo
- Indicadores automáticos de cumplimiento en rojo y ámbar para el personal con exceso de trabajo

*Resultado:* identifica tus cuellos de botella operativos exactos en **60 segundos** y captura o imprime los gráficos directamente para tu reunión semanal de dirección.

### Paso 4: reutiliza mes a mes y ahorra horas en temporada alta (obtén el archivo de Excel)
Cuando cierre el período de pago del mes siguiente, pega las filas de la hoja de horas del nuevo mes. La plantilla conserva los meses históricos y actualiza automáticamente tu trayectoria acumulada del año (YTD).

> 💡 **Convierte horas de lucha con hojas de cálculo en una rutina de copiar y pegar de 1 minuto:**
> ¿Por qué perder las noches de temporada alta corrigiendo fórmulas rotas o creando tablas dinámicas desde cero? Aunque no tengas ninguna experiencia en análisis de datos, este modelo de Excel ya diseñado hace por ti el trabajo matemático pesado. Pega tus cifras brutas de nómina, detecta tus decisiones de gestión clave en 1 minuto y vuelve a ocuparte de tu negocio principal.
>
> 👉 **[Descarga aquí la plantilla completa y reutilizable de Excel para el seguimiento de nómina y la planificación de capacidad de personal (.xlsx)](https://www.theseusworkshop.com/l/ufscrp?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=workforce-planning-payroll-dashboard)** para guardar los datos de tu empresa localmente, ejecutar análisis de nómina mensuales ilimitados sin conexión y pagar una sola vez sin cuotas recurrentes de SaaS.

---

## Comparativa: ¿por qué usar esta plantilla de Excel automatizada en lugar de hojas de cálculo manuales o software de RR. HH. heredado?

| Escenario operativo y métrica | Hojas de cálculo de Excel manuales y exportaciones básicas de nómina (sin esta herramienta) | Kit de Excel automatizado de capacidad de personal y nómina (con esta herramienta) |
| :--- | :--- | :--- |
| **Auditoría mensual de nómina y horas extra** | Los contables auditan los salarios brutos solo *después* de que cierran los períodos contables y de que el dinero ya se gastó, sin margen para evitar desviaciones en el presupuesto de horas extra. | El seguimiento continuo por período de pago del **salario base, las primas por horas extra de 1,5x/2,0x y los complementos** pone de relieve los picos de costo antes de que se agote el presupuesto trimestral. |
| **Mantenimiento de fórmulas y entrada de datos** | Los usuarios pasan horas arrastrando fórmulas hacia abajo, corrigiendo errores `#REF!` y `VLOOKUP` y actualizando manualmente los rangos de los gráficos cada mes. | Las **fórmulas nativas de matriz dinámica** predefinidas relacionan automáticamente las filas pegadas y actualizan todos los gráficos de KPI: **sin hacer doble clic ni arrastrar fórmulas**. |
| **Horas contratadas anuales y cumplimiento** | Los supervisores de turno controlan las horas FTE en notas sueltas y desconectadas, y solo descubren los incumplimientos de las horas de trabajo cuando los empleados se agotan o renuncian. | El cálculo automático de las **horas trabajadas acumuladas en el año frente a los límites contratados anuales** activa alertas visuales instantáneas cuando el personal supera tu umbral de utilización del 85%. |
| **Consolidación de datos de varios departamentos** | RR. HH. y Finanzas pierden de 1 a 2 días al mes fusionando pestañas de hojas de horas inconsistentes de entre 5 y 7 responsables de departamento distintos. | Hasta **7 pestañas departamentales de entrada estandarizadas** alimentan automáticamente un único panel maestro centralizado al pegar los datos brutos de fichaje en CSV. |
| **Momento de decidir la plantilla y las contrataciones** | Los gerentes solicitan nuevas contrataciones por intuición o por picos de una sola semana, sin comprobar si los equipos vecinos tienen horas disponibles. | Las **tasas de utilización FTE por departamento** comparadas en paralelo demuestran si necesitas nueva plantilla o si puedes reasignar tareas a equipos con capacidad restante. |

---

## ¿Para quién es esta plantilla de Excel de nómina y planificación de capacidad? (por rol y escenario de búsqueda)

Esta hoja de cálculo lista para usar está pensada para gerentes, dueños de negocios y analistas que tienen exportaciones brutas de hojas de horas o de nómina (de ADP, Gusto, Paychex, QuickBooks o Rippling) y necesitan **control de costos y visibilidad de la carga de trabajo** inmediatos sin pagar $5,000+/año por software empresarial de gestión de la fuerza laboral (WFM):

- **Dueños de pequeñas empresas y fundadores de startups que buscan una `Simple Monthly Payroll & Overtime Tracker Excel Sheet`:** deja de preguntarte por qué la nómina mensual no deja de subir. Pega las cifras de tu proceso de pago para ver el salario regular exacto frente al pago por horas extra de todo tu equipo en 1 minuto.
- **Gerentes de RR. HH. y responsables de People Ops que buscan una `Annualized Hours & Employee Utilization Tracking Template`:** monitorea las horas trabajadas acumuladas en el año frente a los contratos FTE anuales, previene el agotamiento del personal y anticípate a los límites máximos de horas de la legislación laboral.
- **Administradores de nómina y contables que necesitan una `Multi-Department Timesheet Consolidation Spreadsheet`:** agrega las horas regulares, los multiplicadores de horas extra y los complementos por turno de hasta 7 departamentos en una única tabla de resumen clara sin escribir ni una sola macro.
- **Gerentes de operaciones, almacén y turnos que buscan una `Staffing Capacity & Workload Planning Calculator`:** comprueba cuántas horas de mano de obra utilizables le quedan a cada equipo antes de la temporada alta para equilibrar los horarios de los turnos y evitar las urgencias de última hora por horas extra.
- **Directores financieros, contadores y CFO fraccionales que necesitan un `Labor Cost Variance & Headcount Budget Model`:** audita el ritmo de consumo de mano de obra por departamento, separa el crecimiento estructural de los salarios de la fuga por horas extra y entrega a tus clientes gráficos de personal listos para el consejo.

---

## ¿Cuándo necesitas una hoja de cálculo de horas de empleados y costo de mano de obra lista para usar? (escenarios habituales)

Si reconoces alguna de estas situaciones operativas cotidianas, esta plantilla te ofrece una solución inmediata hoy mismo:

1. **"Mi factura mensual de nómina subió un 18% y no sé qué departamento o qué turno de horas extra lo provocó."** → Pega tus fichajes mensuales para clasificar al instante los departamentos por proporción de horas extra y gasto en complementos.
2. **"Estamos entrando en la temporada alta del tercer y cuarto trimestre y no sé si mi personal actual puede asumir los pedidos adicionales."** → Consulta la columna **Capacidad anual restante** para ver las horas disponibles exactas de cada equipo antes de comprometerte con nuevas contrataciones.
3. **"Odio pasar la tarde del viernes corrigiendo fórmulas de Excel rotas y actualizando tablas dinámicas para la reunión de operaciones del lunes."** → Pega tus datos brutos en las tablas de entrada ya formateadas y deja que el panel automatizado genere los visuales de tu reunión en 60 segundos.
4. **"Nuestro servicio de nómina procesa los cheques de pago sin problema, pero sus paneles de informes no sirven para planificar la capacidad."** → Exporta el archivo CSV bruto de fichaje de tu proveedor de nómina, cárgalo en este libro de trabajo y obtén la visión estratégica de capacidad que tu software de nómina omite.

---

## La lógica de negocio y la metodología: 3 puntos ciegos ocultos de la nómina explicados de forma sencilla

No necesitas un MBA ni un título en finanzas para aplicar análisis inteligentes de la fuerza laboral. La mayoría de las empresas pierde dinero con la mano de obra porque los informes contables estándar ocultan tres verdades operativas muy simples. Esta es la lógica de negocio, en lenguaje claro, integrada en esta plantilla de Excel, y cómo protege tu margen de beneficio:

### 1. Punto ciego n.º 1: controlar los "dólares gastados" oculta las "horas restantes" (tratar la mano de obra como un inventario de horas)
* **El hecho objetivo:** el software de nómina estándar solo contabiliza el dinero *después* de que lo gastas. Trata la mano de obra como una factura mensual abierta en lugar de como un conjunto fijo de horas que compraste para el año.
* **Por qué esto te cuesta dinero:** supongamos que un empleado tiene un contrato de 2.000 horas de trabajo al año. Si una mala programación de turnos consume 1.300 de esas horas en junio, un informe contable mensual estándar seguirá pareciendo normal, hasta que llegue octubre y el empleado alcance su techo legal o contractual de horas en pleno trimestre de mayores ventas. Entonces te ves obligado a detener pedidos, pagar tarifas de horas extra de 1,5x a 2,0x o asumir los honorarios de un reclutador para contratar personal temporal de emergencia.
* **Cómo lo resuelve la plantilla automáticamente:** el libro de trabajo deduce cada hora registrada del fondo de **horas contratadas anuales** de cada empleado. Ves tu **inventario exacto de horas restantes** cada mes, lo que te da meses de margen para reequilibrar las cargas de trabajo antes de que llegue una crisis de personal.

### 2. Punto ciego n.º 2: una sola cifra de "nómina total" oculta la fuga por horas extra (salario base frente a horas extra variables)
* **El hecho objetivo:** mirar el "pago bruto mensual total" mezcla las horas base productivas con horas de penalización costosas (horas extra a 1,5x, pago de festivos a 2,0x y complementos puntuales).
* **Por qué esto te cuesta dinero:** pagar horas extra a 1,5x en un equipo de 10 personas durante tres meses seguidos suele costar más en efectivo puro que contratar a una persona adicional a tiempo completo con un salario base estándar de 1,0x, además de agotar físicamente a tus mejores trabajadores y aumentar la rotación.
* **Cómo lo resuelve la plantilla automáticamente:** el motor de cálculo divide cada dólar en **salario base estructural**, **primas por horas extra** y **complementos**. De un vistazo sabes si los altos costos de mano de obra provienen de un pico temporal de dos semanas (mantén la plantilla actual) o de una escasez permanente de carga de trabajo (resulta más barato contratar a un nuevo miembro del equipo).

### 3. Punto ciego n.º 3: pagar horas extra en un equipo mientras otro tiene capacidad libre (programación en silos)
* **El hecho objetivo:** cuando los gerentes de departamento controlan los horarios en hojas de cálculo separadas, el equipo A aprueba habitualmente horas extra costosas o solicita nuevas contrataciones mientras el equipo B está al 65% de su capacidad.
* **Por qué esto te cuesta dinero:** cada hora extra pagada mientras otro equipo cualificado tiene horas contratadas sin usar es una deducción directa y evitable del beneficio neto de tu empresa.
* **Cómo lo resuelve la plantilla automáticamente:** al reunir hasta 7 departamentos en una comparación paralela de **tasa de utilización de mano de obra (%)**, detectas de inmediato qué equipos están sobrecargados (>85% de capacidad) y qué equipos tienen capacidad libre (<70% de capacidad) para absorber el trabajo próximo sin costo adicional de mano de obra.

---

## Especificaciones técnicas y compatibilidad de archivos

- **Formato:** libro de trabajo nativo `.xlsx` + vista previa HTML interactiva en el navegador
- **Compatibilidad con Excel:** Microsoft Excel 2021+, Microsoft 365 (Windows y Mac) mediante funciones nativas de matriz dinámica
- **Macros / VBA:** 100 % sin macros (formato `.xlsx`: sin advertencias de seguridad, bloqueos de TI ni scripts VBA rotos)
- **Capacidad:** preconfigurado para la consolidación de varios departamentos (hasta 7 pestañas departamentales localizadas que alimentan 1 panel ejecutivo centralizado)

---

## Acerca de Theseus Workshop | Herramientas de decisión empresarial listas para usar

**Theseus Workshop** (*Pure utility. Zero fluff.*) crea modelos financieros de Excel ligeros y listos para usar, estimadores de costos paramétricos y paneles operativos de decisión para fundadores, responsables de operaciones y equipos financieros. Sustituimos el software empresarial sobrecargado y las frágiles hojas de cálculo improvisadas por herramientas limpias de Microsoft 365 sin macros que aportan claridad de gestión inmediata en menos de 1 minuto.

- 🛒 **Explora nuestro catálogo completo de herramientas en Gumroad:** [Descubre modelos financieros de Excel, paneles de flujo de caja y plantillas de estimación listos para usar en Theseus Workshop (theseusworkshop.com)](https://www.theseusworkshop.com/?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20footer&utm_content=payroll-workforce-planning)
- 📥 **Obtén este kit de personal y nómina:** [Descarga la plantilla completa de Excel de seguimiento de nómina empresarial y planificación de capacidad de personal (.xlsx)](https://www.theseusworkshop.com/l/ufscrp?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20footer&utm_content=payroll-workforce-planning)
- 🌐 **Prueba la versión en vivo para el navegador:** [Abre el panel gratuito e interactivo de capacidad de personal y nómina (HTML)](https://hyvoid.github.io/workforce-planning-payroll-excel-dashboard/)


---

## Arquitectura Técnica y Referencia de Fórmulas

<details>
<summary>Para desarrolladores de Excel, modeladores financieros y analistas de datos</summary>

### Arquitectura Relacional del Libro de Trabajo

La hoja de cálculo opera bajo una estricta separación de responsabilidades: Configuración, Datos Maestros, Entrada Transaccional y Presentación Agregada. Esto asegura alta integridad de datos y cero sobrescritura de fórmulas durante las actualizaciones mensuales.

| Capa Arquitectónica | Hojas Asignadas | Función Técnica Central |
|---------------------|---------------------|-------------------------|
| **Parámetros Globales** | `Settings` | Variables centralizadas (Año Fiscal, Multiplicador de Horas Extra, Umbrales de KPI). |
| **Gestión de Datos Maestros**| `Employee_Master` | Base de datos de clave primaria para IDs de Empleado, Tarifas Base y Objetivos FTE Anuales. |
| **Entrada Transaccional** | `Dept_1` – `Dept_7` | Entidades aisladas de entrada de tarjetas de tiempo mensuales para gerentes de primera línea. |
| **Agregación de Datos** | `Payroll_Summary`, `Annual_Hours_Tracker` | Modelado de datos entre hojas utilizando arreglos dinámicos, `SUMIFS` y algoritmos de utilización. |
| **UI de Presentación** | `Dashboard` | Resúmenes ejecutivos de solo lectura, formato condicional y trazado de gráficos. |

### Lógica de Cálculo Central

La herramienta utiliza estrictamente funciones nativas de Microsoft 365 / Excel 2021+. **Cero dependencias de macros (VBA) o Power Query.**

#### Recuperación de Datos de Empleados (XLOOKUP)
```excel
=XLOOKUP(Employee_ID, Employee_Master[Employee_ID], Employee_Master[Hourly_Rate], 0, 0)

```

*Garantiza que los libros de nómina posteriores reflejen instantáneamente los cambios de tarifa realizados en el conjunto de datos Maestro.*

#### Agregación de Nómina Bruta entre Departamentos

```excel
=SUMIFS(Dept_1[Gross_Pay], Dept_1[Employee_ID], [@Employee_ID]) + SUMIFS(Dept_2...

```

*Crea un roll-up empresarial sin problemas sin requerir actualizaciones complejas de tablas dinámicas.*

#### Motor de Clasificación de Riesgo de Capacidad

```excel
=IFS(
    Utilization_Rate >= Settings!High_Threshold, "High Risk - Burnout",
    Utilization_Rate <= Settings!Low_Threshold, "Under-Utilized - Idle Capacity",
    TRUE, "Optimal Allocation"
)

```

*Impulsa el formato condicional en el panel ejecutivo para resaltar las intervenciones requeridas.*

</details>

---

## Explora más plantillas de operaciones de negocio

Me especializo en diseñar modelos de apoyo a la toma de decisiones en Excel ligeros y de alto impacto para flujos de trabajo operativos complejos. Explora el ecosistema completo de herramientas:

* **[project-time-tracking-excel-template](https://github.com/HyVoid/project-time-tracking-excel-template)** — ¿Qué proyectos se están desviando en costo de tiempo o margen? — Analiza la asignación de costos de mano de obra por proyecto y la rentabilidad de las carteras de proyectos activos.
* **[Caf-audit-tool-for-efficient-reviews](https://github.com/HyVoid/Caf-audit-tool-for-efficient-reviews)** — ¿Qué dimensión operativa está consumiendo la mayor parte de los ingresos y la fuga financiera exige una intervención inmediata? — Convierte observaciones estructuradas de cafeterías en una puntuación de diagnóstico ponderada, una estimación de la fuga financiera y una lista priorizada de acciones correctivas.
* **[Restaurant-Menu-Configuration-Modifier-Pricing-Toolkit](https://github.com/HyVoid/Restaurant-Menu-Configuration-Modifier-Pricing-Toolkit)** — ¿El menú tiene precios coherentes en todas las combinaciones de modificadores y tamaños, o hay errores de precios ocultos a punto de aparecer? — Centraliza las reglas de precios de modificadores y configuración del menú de restaurante y los datos del menú listos para el TPV en una única fuente de verdad gobernada.
* **[wedding-planner](https://github.com/HyVoid/wedding-planner)** — ¿Qué tareas, costos y plazos de la planificación de bodas necesitan coordinación? — Coordina la gestión de proveedores, el seguimiento de tareas y el control presupuestario en los flujos de trabajo de planificación de bodas.
* **[manufacturing-capacity-planning-template](https://github.com/HyVoid/manufacturing-capacity-planning-template)** — ¿Se está utilizando eficientemente la capacidad de mano de obra o la operación necesita otra dotación de personal u otra programación? — Transforma los planes de producción y los datos de mano de obra en decisiones de utilización de capacidad, asignación de costos de mano de obra y dotación de personal.
* **[cin7-excel-reporting-template](https://github.com/HyVoid/cin7-excel-reporting-template)** — ¿Qué dicen los datos operativos de Cin7 sobre el rendimiento del negocio? — Traduce los datos de inventario y operaciones de Cin7 en información accionable sobre el rendimiento del negocio.
* **[employee-performance-dashboard-excel](https://github.com/HyVoid/employee-performance-dashboard-excel)** — ¿Qué empleados y proyectos corren riesgo de incumplir la entrega y dónde debe intervenir la dirección antes de que venzan los plazos? — Conecta la planificación anual del trabajo, el seguimiento de la ejecución, la evaluación ponderada del desempeño y la elaboración de informes de dirección en un flujo de trabajo operativo continuo.
* **[rental-property-management-excel-template](https://github.com/HyVoid/rental-property-management-excel-template)** — ¿Qué unidades están perdiendo ingresos por vacantes no detectadas y qué rotaciones requieren acción inmediata? — Transforma los registros de reservas en estado de ocupación en vivo, inteligencia sobre vacantes y alertas operativas en toda una cartera de alquileres.
* **[warehouse-billing-excel-template](https://github.com/HyVoid/warehouse-billing-excel-template)** — ¿Los cargos de facturación del almacén son precisos y rastreables hasta los registros operativos subyacentes? — Calcula los cargos de almacenamiento, manipulación y transporte del almacén y conviértelos en estados de facturación auditables para el cliente.
* **[commercial-hvac-plumbing-bid-estimator](https://github.com/HyVoid/commercial-hvac-plumbing-bid-estimator)** — ¿Cuál es un precio de licitación defendible y auditable para este proyecto comercial de instalaciones mecánicas, eléctricas y de fontanería (MEP)? — Estima y controla los costos de proyectos MEP comerciales desde la medición de cantidades hasta la presentación final de la oferta.
* **[landed-cost-calculator-excel](https://github.com/HyVoid/landed-cost-calculator-excel)** — ¿Por qué el costo landed a nivel de SKU de mi ERP difiere del de la API del transitario y qué método de asignación está causando la erosión del margen? — Concilia y compara los cálculos de costo landed con una plantilla de referencia para rastrear las diferencias de asignación de costos a nivel de SKU hasta su lógica contable subyacente.
* **[hoa-budget-accounting-spreadsheet](https://github.com/HyVoid/hoa-budget-accounting-spreadsheet)** — ¿Se están asignando correctamente los fondos de nuestra asociación de propietarios y vamos según lo previsto con el presupuesto? — Clasifica las transacciones y genera informes de variación presupuestaria para la gestión financiera de la asociación de propietarios.
* **[homeowner-management-excel-template](https://github.com/HyVoid/homeowner-management-excel-template)** — ¿Cuál es mi costo total real de propiedad de la vivienda, incluidos todos los gastos ocultos? — Calcula el costo total de propiedad de la vivienda integrando la hipoteca, los servicios públicos, el mantenimiento y las mejoras de capital.
* **[crew-scheduling-excel-template](https://github.com/HyVoid/crew-scheduling-excel-template)** — ¿Podrá mi equipo actual asumir la carga de trabajo cuando las fechas de los proyectos se desplacen y se solapen? — Pronostica la demanda de personal y las brechas de dotación en un horizonte de 12 meses para respaldar las decisiones de planificación de capacidad.
* **[Schedule-Driven-Construction-Cash-Curve-Workbook](https://github.com/HyVoid/Schedule-Driven-Construction-Cash-Curve-Workbook)** — ¿Cómo evolucionarán con el tiempo el flujo de caja y el avance planificado de mi proyecto de construcción según el cronograma? — Convierte cronogramas de construcción y registros de partidas de licitación en curvas S automatizadas, previsiones de costos por fases temporales y rastreadores del avance planificado.
* **[shop-floor-mrp-kanban-excel-system](https://github.com/HyVoid/shop-floor-mrp-kanban-excel-system)** — ¿Nuestro programa de soldadura por turnos se ejecutará sin desabastecimientos que detengan la línea? — Calcula las necesidades de materiales derivadas de la BOM y las señales de reposición Kanban a partir de los programas de producción por turnos.

---

## Licencia

Licenciado bajo la **Licencia Apache 2.0**.

Eres libre de usar, modificar y distribuir este marco de software tanto para uso comercial como organizacional privado, de acuerdo con los términos de la Licencia Apache 2.0.

Copyright © 2026.

