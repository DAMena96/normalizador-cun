# Normalizador CUN – versión con JSON

Esta versión ya **no depende de la hoja `VAL` dentro del Excel Sin Gestión**.

El Excel que cargas puede traer solo la base de leads. Los catálogos de normalización viven en:

```text
assets/data/asesores.json
assets/data/programas.json
assets/data/ciudades.json
assets/data/llaves.json
assets/data/areas.json
```

## Cómo editar un asesor

Abre:

```text
assets/data/asesores.json
```

Ejemplo:

```json
{
  "Vinculaciones Carlos Andres Burgos Carvajal": "Carlos Burgos",
  "Contact Allison Campos Suarez": "Allison Campos"
}
```

La parte izquierda es **como llega en el Excel**.
La parte derecha es **como quieres que aparezca normalizado**.

## Cómo editar un programa

Abre:

```text
assets/data/programas.json
```

Ejemplo:

```json
{
  "Ingeniría de sistemas": "INGENIERIA DE SISTEMAS",
  "ingeniería_de_sistemas": "INGENIERIA DE SISTEMAS"
}
```

## Cómo editar áreas

Abre:

```text
assets/data/areas.json
```

Ejemplo:

```json
{
  "INGENIERIA DE SISTEMAS": "MIXTO",
  "DERECHO": "PRESENCIAL",
  "ESPECIALIZACION EN MARKETING DIGITAL": "POSGRADO"
}
```

## Validación final

La regla está en:

```text
assets/js/03-normalizer.js
```

Función:

```javascript
calcularValidacionFinal(llave, area)
```

Replica esta fórmula:

```excel
=SI(Y(AP2="-";AN2="PRESENCIAL");"PRESENCIAL";
 SI(Y(AP2="-";O(AN2="MIXTO";AN2="OTRO";AN2="VIRTUAL"));"VIRTUAL";
 SI(AP2="PRESENCIAL";"PRESENCIAL";
 SI(AN2="PRESENCIAL";"PRESENCIAL";"POSGRADO"))))
```

## Importante

Debes abrir el proyecto con **Live Server** en VS Code.
No lo abras con doble clic, porque el navegador puede bloquear la carga de archivos JSON.


## Interfaz App Normalizador Contact CUN

Esta versión incluye una pantalla inicial con colores corporativos CUN:

- Brand colors: `#84BD00`, `#007A33`, `#898D8D`
- Complementary colors: `#0C2340`, `#00B388`, `#00859B`, `#005EB8`, `#1B365D`, `#FFCD00`

Módulos disponibles:

1. Base Sin Gestión: activo y conectado al flujo actual.
2. Base de Interesados: preparado para implementación.
3. Base No Contactado: preparado para implementación.

Recuerda abrir el proyecto con Live Server para que carguen correctamente los archivos JSON.


## Módulo Base de Interesados

Esta versión incluye el módulo **Base de Interesados** con dos salidas:

1. **Predictivo**:
   - CONTACTID
   - Prioridad
   - email
   - First_name
   - telefono
   - Programa
   - number 1
   - AgentName

2. **Base de Interesados normalizada**:
   - PROGRAMA2
   - CIUDAD
   - AREA
   - AREA VALIDA
   - VAL
   - SUPERVISOR
   - FECHA CREACION

Filtros incluidos:
- AREA VALIDA
- VAL
- FECHA CREACIÓN

Nota: el módulo reutiliza los catálogos JSON existentes (`programas.json`, `ciudades.json`, `areas.json`, `llaves.json`, `asesores.json`) para reemplazar las búsquedas que antes se hacían en la hoja VAL.


Actualización: el filtro de fecha ahora usa **FECHA CREACION** y el predictivo se exporta con los mismos filtros activos de la base de interesados.


## Ajuste Área válida y Distribución Interesados

- Área válida ahora usa la lógica:
  - Si AREA es `MIXTO`, revisa si `VAL` contiene `V`.
  - Si no, revisa si `utm_term` contiene `V`.
  - Si no encuentra `V`, queda `PRESENCIAL`.
  - Si AREA no es `MIXTO`, conserva el área original.

Esto ayuda a acercar los totales a la validación original del Excel.

- Se agregó pestaña **Distribución** en Base de Interesados.
- La distribución usa los registros filtrados y exporta:
  - `ID de registro`
  - `CORREO ASESOR ASIGNADO`


## Corrección Área válida — Interesados

Regla aplicada exactamente como se indicó:

```excel
=SI(AREA="MIXTO";
   SI(CONTAR.SI(VAL;"*V*")>0;"VIRTUAL";
      SI(CONTAR.SI(conv;"*V*")>0;"VIRTUAL";"PRESENCIAL"));
   AREA)
```

Detalles:
- `PROGRAMA2` se normaliza desde el programa de la base de interesados.
- `AREA` se toma desde `areas.json` usando `PROGRAMA2`.
- `VAL` se toma desde `llaves.json` usando `Campaña mercadeo`.
- `AREA VALIDA` mira primero `VAL` y después `conv`.


## Corrección final Programa / Área — Interesados

Flujo aplicado:

1. Se arma el **Programa** igual que en Predictivo:
   - Si la columna `Programa` viene vacía, usa `Programa de interes_`.
   - Si también está vacío, queda `Sin programa`.
   - Se aplica formato tipo nombre propio.

2. Ese **Programa** se cruza con `programas.json` para normalizarlo.

3. El resultado normalizado se cruza con `areas.json` para obtener `AREA`.

4. `AREA VALIDA` aplica:
   - Si `AREA = MIXTO`, mira primero `VAL`.
   - Si `VAL` contiene `V`, queda `VIRTUAL`.
   - Si no, mira `conv`.
   - Si `conv` contiene `V`, queda `VIRTUAL`.
   - Si no, queda `PRESENCIAL`.
   - Si `AREA` no es `MIXTO`, conserva `AREA`.

La distribución de interesados quedó con el mismo flujo visual que Sin Gestión:
cargar Excel de asesores con columna `Correo` o `Email`, distribuir y exportar.


## Corrección UI y Distribución Interesados

- Los filtros de Base de Interesados ahora se cierran después de seleccionar una opción.
- La pestaña Distribución de Base de Interesados quedó con el mismo flujo de Base Sin Gestión:
  1. Leads a distribuir desde el filtro activo.
  2. Cargar Excel de asesores con columna `Correo` o `Email`.
  3. Distribuir, exportar y copiar IDs.
- Se agregó botón `Copiar IDs` en la tabla de Base Interesados.


## Ajuste Área válida — conv y Periodo

Regla actualizada:

- Solo aplica si `AREA` es `MIXTO`.
- Primero mira `conv`:
  - Si `conv` contiene `V`, queda `VIRTUAL`.
  - Si `conv` tiene valor pero no contiene `V` —por ejemplo `2023C`— queda `PRESENCIAL`.
- Si `conv` está vacío, mira `Periodo`:
  - Si `Periodo` contiene `V` —por ejemplo `26V03`— queda `VIRTUAL`.
  - Si no contiene `V`, queda `PRESENCIAL`.
- Si `AREA` no es `MIXTO`, conserva el valor original del área.


## Corrección filtros y VAL

- Los filtros ahora se cierran al hacer clic fuera del selector.
- El filtro `VAL` ya no usa la campaña como respaldo; solo muestra valores cruzados desde `llaves.json`.
- Los valores de `AREA` y `AREA VALIDA` se normalizan en mayúsculas para evitar duplicados como `Otro` y `OTRO`.


## Corrección VAL y módulo Base No Contactado

### VAL
La columna `VAL` ahora se calcula cruzando el programa normalizado contra `areas.json`.
El resultado esperado es:
- `MIXTO`
- `PRESENCIAL`
- `VIRTUAL`
- `POSGRADO`
- `OTRO`

Luego `AREA VALIDA` aplica la regla:
- Si `VAL = MIXTO`, mira `conv`.
- Si `conv` está vacío, mira `Periodo`.
- Si contiene `V`, queda `VIRTUAL`.
- Si no contiene `V`, queda `PRESENCIAL`.
- Si `VAL` no es `MIXTO`, conserva el valor de `VAL`.

### Base No Contactado
Se agregó el módulo completo:
- Predictivo.
- Base No Contactado normalizada.
- Filtros por Área válida, VAL y Fecha creación.
- Distribución de leads igual a Base Sin Gestión.
- Exportaciones y copiar IDs.


## No Contactado — Distribución CARI AI y Notas

Se agregaron dos pestañas nuevas dentro de Base No Contactado:

### Distribución CARI AI

Exporta:
- `numero_telefono`
- `nombre_aspirante`
- `carrera_interes`
- `identificacion`
- `correo_electronico`
- `telefono_adicional`
- `periodo`
- `campania`

Reglas:
- `numero_telefono`: teléfono normalizado con prefijo 57 y sin signo `+`.
- `carrera_interes`: Programa de la base.
- `periodo`: si `AREA VALIDA` es `POSGRADO`, queda `Virtual/Especialización`; en los demás casos queda `Virtual/pregrado`.
- `campania`: Campaña mercadeo.

### Notas CARI AI

Exporta:
- `ID`
- `Notes`

La nota fija siempre es: `Se intenta contacto con el aspirante via CARI AI`.


## Ajuste CARI AI

- `carrera_interes` ahora usa `PROGRAMA NORMALIZADO`.
- `identificacion` usa la columna original `Número de Documento`; si está vacía queda `Sin documento`.
- `telefono_adicional` queda igual a `numero_telefono`.
- Se agregaron dos registros fijos que siempre salen en CARI AI, aunque se apliquen filtros:
  - Esteban
  - Diguar
- Los registros fijos no salen en Notas CARI AI.


## Módulo Recibos — Solo CARE AI y Notas

Se agregó el módulo **Recibos CARE AI**.

Columnas esperadas de la base:
- ID de registro
- Hora de creación
- Correo electrónico personal
- Estado del interesado
- Etapa del Registro
- Propietario de Interesado
- Nombre completo
- Campaña mercadeo
- Teléfono
- Número de Documento
- Programa de interes
- Correo institucional
- Número de gestiones
- Modalidad
- Periodo
- Tipo de homologación
- Sub.estado
- Sub.estado II
- Hora de la última actividad
- Estado a tipificar
- Creado por
- Ciudad de residencia
- Fecha última nota
- utm_campaign
- Programa de interés_
- Fecha pago oportuno
- Hora de modificación
- Programa

### Salida CARE AI

Exporta:
- numero_telefono
- nombre_aspirante
- carrera_interes
- identificacion
- correo_electronico
- telefono_adicional
- periodo
- campania

Reglas:
- `numero_telefono`: teléfono normalizado con 57, sin signo `+`.
- `telefono_adicional`: igual a `numero_telefono`.
- `carrera_interes`: programa normalizado.
- `identificacion`: Número de Documento; si está vacío queda `Sin documento`.
- Validación teléfono:
  - vacío o 0: `Sin número`
  - menos de 10 dígitos: `Número errado`
  - primeros 2 dígitos menor a 30 o mayor a 35: `Número errado`
  - si no: `Ok`
- Incluye siempre los dos registros fijos: Esteban y Diguar.

### Notas

Exporta:
- ID
- Notes

Nota fija:
`Se intenta contacto con el aspirante via CARI AI`


## Distribución balanceada por Hora de creación

Se actualizó la lógica de **Distribución de Leads** en:
- Base Sin Gestión
- Base de Interesados
- Base No Contactado

Ahora los leads se ordenan por `Hora de creación` antes de repartirlos entre asesores, para que cada asesor reciba una mezcla proporcional de leads antiguos y nuevos.


## Base Sin Gestión — Distribución CARI AI y Notas

Se agregaron dos pestañas dentro de Base Sin Gestión:

- Distribución CARI AI
- Notas CARI AI

La salida CARI AI exporta:
- numero_telefono
- nombre_aspirante
- carrera_interes
- identificacion
- correo_electronico
- telefono_adicional
- periodo
- campania

Incluye siempre los dos registros fijos Esteban y Diguar. Las notas exportan ID y la nota fija:
`Se intenta contacto con el aspirante via CARI AI`.

## Cambios CARI AI y Detalle asesor

- En Base Sin Gestión y Base No Contactado, la columna `campania` de CARI AI ahora queda siempre como `Organico`.
- Recibos no se modificó en esta regla.
- Se agregó **Detalle asesor por mentor** después de la distribución de leads en:
  - Base Sin Gestión
  - Base de Interesados
  - Base No Contactado
- El detalle muestra cuántos leads recibió cada asesor agrupado por mentor y permite filtrar por mentor.


## Fix Detalle asesor

Se corrigió la visibilidad de la tarjeta **Detalle asesor por mentor** para que aparezca automáticamente después de hacer clic en Distribuir. Ahora también muestra tarjetas resumen arriba:
- Leads distribuidos
- Asesores con leads
- Mentores


## Detalle asesor visible

Se movió la tarjeta **Detalle asesor por mentor** para que quede visible dentro de la pestaña Distribución antes de la tabla. Ahora se ve desde el inicio con un mensaje y, al hacer clic en Distribuir, se llena automáticamente con los datos.


## Fix final Detalle asesor

Se agregaron funciones de respaldo para que el resumen se llene sí o sí después de Distribuir en:
- Base Sin Gestión
- Base de Interesados
- Base No Contactado


## Supervisor.json — Detalle asesor por mentor

Se agregó el catálogo `assets/data/Supervisor.json`.

La tarjeta **Detalle asesor por mentor** ahora toma el mentor/supervisor desde el asesor asignado:
- Cruza `CORREO ASESOR ASIGNADO` contra `Supervisor.json`.
- Usa el supervisor/mentor del JSON para agrupar.
- También muestra el área del asesor.
- Aplica en Base Sin Gestión, Base de Interesados y Base No Contactado.
- No aplica para Recibos.


## Corrección final Supervisor, Detalle y Distribución

- `Supervisor.json` queda integrado en `assets/data/Supervisor.json`.
- El supervisor se toma del nombre de la primera columna del JSON.
- El asesor se toma del valor de la primera columna del JSON.
- El área se toma de la segunda columna del JSON.
- El grupo/tipo se toma de la tercera columna del JSON.
- El filtro de Detalle ahora muestra supervisores reales, no áreas ni nombres de asesores.
- La distribución ahora sí mezcla antiguos y nuevos: ordena por `Hora de creación` y asigna round-robin a los asesores.
- Se agregó reset visual al volver al inicio para que no quede pegada la última base en otro módulo.


## Fix definitivo Detalle asesor por supervisor

Se agregó `assets/js/13-detalle-supervisor-final.js`, cargado al final del proyecto, para garantizar que:
- El filtro muestra supervisores reales desde `Supervisor.json`.
- El detalle se llena al dar clic en Distribuir.
- La agrupación es por supervisor y dentro muestra asesores con cantidad de leads.
- La distribución mezcla leads antiguos y nuevos ordenando por `Hora de creación` y asignando en round-robin.


## Detalle asesor simplificado y funcional

Se eliminó el filtro visual de Mentor para evitar confusión.
Ahora el detalle se genera directamente después de distribuir y muestra:
- Supervisor
- Asesor
- Área
- Grupo
- Cantidad de leads

También incluye botón para exportar el detalle a Excel.

La distribución ordena por `Hora de creación` y reparte round-robin entre asesores para mezclar leads antiguos y nuevos.


## Corrección final pedida

- La tabla de distribución queda arriba.
- Debajo de esa tabla queda la sección **Detalle asesor por supervisor**.
- Se eliminó visualmente el filtro de mentor.
- El detalle agrupa por Supervisor → Asesor → Cantidad de leads.
- Se agregó exportación de detalle asesor.
- En filtros de Sin Gestión, el filtro MENTOR ahora se llama Supervisor.
- El campo MENTOR se llena cruzando el Asesor con Supervisor.json.
- La distribución mezcla antiguos y nuevos ordenando por Hora de creación y asignando round-robin.


## Fix definitivo Detalle asesor

El detalle ahora se genera leyendo directamente la tabla de distribución visible:
- `ID de registro`
- `CORREO ASESOR ASIGNADO`

Luego cruza `CORREO ASESOR ASIGNADO` contra `Supervisor.json` con formato:
`{ "asesor@cun.edu.co": "Supervisor" }`.

Esto evita fallas con variables internas y garantiza que el detalle salga después de distribuir.


## Fix Detalle asesor por supervisor y botón volver

- Se corrigió Base Sin Gestión: la tabla de distribución ahora conserva el id `dist-tbl`.
- El detalle lee la tabla visible y cruza `CORREO ASESOR ASIGNADO` contra `Supervisor.json`.
- El botón Volver al inicio de Sin Gestión ahora usa el mismo estilo de los otros módulos.


## Fix final detalle y botón

- El detalle ahora se pinta usando directamente `distribuidoData`, `interDistribuidoData` y `ncDistribuidoData`.
- Ya no depende de leer la tabla HTML.
- Se refresca automáticamente después de hacer clic en Distribuir.
- El botón Volver al inicio de Base Sin Gestión queda forzado con el mismo estilo visual.


## Ajuste topbar Sin Gestión

Se agregó `sin-gestion-topbar` para que Base Sin Gestión tenga el mismo encabezado visual que Base de Interesados y No Contactado.

## Detalle asesor

El detalle NO depende ya del HTML de la tabla. El último script `18-detalle-render-directo.js` toma directamente:
- `distribuidoData`
- `interDistribuidoData`
- `ncDistribuidoData`

y cruza `CORREO ASESOR ASIGNADO` contra `assets/data/Supervisor.json`.

## Corrección final integrada
- El detalle por supervisor ahora está integrado en `08-distribution.js`, cargado antes de Interesados y No Contactado.
- Usa un mapa embebido del nuevo Supervisor.json, formato `{ asesor: supervisor }`.
- Las filas distribuidas se enriquecen con `SUPERVISOR/MENTOR` al momento de distribuir.
- La topbar de Base Sin Gestión usa la misma estructura que los otros módulos.

## Fix visible definitivo
El detalle por supervisor se pinta dentro del mismo `stat-box` de Distribuir y exportar, debajo de los chips de asesores. No depende de la tarjeta 4 ni de leer HTML.

## Ajuste Supervisor_simple
- Se elimina el detalle dentro del resumen.
- El detalle queda debajo de la tabla de distribución.
- El cruce usa únicamente `Supervisor_simple.json` / `Supervisor.json` simple.
- Se eliminan dependencias del JSON anterior.


## Supervisor_simple definitivo

La lógica ahora usa únicamente `assets/data/Supervisor_simple.json`.

Formato correcto:
```json
{
  "correo_o_asesor_izquierda": "supervisor_derecha"
}
```

Cruce aplicado:
- Columna `Asesor` → Supervisor para filtros/normalización.
- `CORREO ASESOR ASIGNADO` → Supervisor para detalle después de distribuir.


## Supervisor_simple corregido

Se regeneró `assets/data/Supervisor_simple.json` con el formato correcto:

```json
{
  "correo_asesor@cun.edu.co": "Nombre del supervisor"
}
```

Ahora el título del grupo en Detalle asesor muestra el **nombre del supervisor**, no Presencial/Otro.
También se corrigió el botón **Exportar detalle asesor**.


## Supervisor_simple exacto

Se eliminó `assets/data/Supervisor.json`.

El proyecto usa únicamente:
`assets/data/Supervisor_simple.json`

La lógica toma:
- izquierda = asesor/correo
- derecha = supervisor

No se regenera ni se transforma el JSON.

## Cambio aplicado — Distribución Sin Gestión con MENTOR

En la pestaña **Base Sin Gestión → Distribución de Leads**, la tabla de distribución y la exportación ahora incluyen la columna:

- `MENTOR`

La columna se toma de `MENTOR`, `SUPERVISOR` o del cruce del asesor asignado contra `Supervisor_simple.json`.
