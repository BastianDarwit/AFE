# AFE · Presentaciones de avance del TFG

Rama única para los hitos de avance. Cada hito tiene su carpeta y sus versiones se numeran dentro de ella.

**Trabajo:** Propuesta de un sistema de gestión de riesgos para los macroprocesos de Gestión de Personas (MP8) y Servicios Generales (MP21) de Celulosa Arauco y Constitución S.A.
**Estudiante:** Bastián Darwit Vera · **Profesor guía:** Carlos Muñoz P. · **Profesor co-guía:** Manuel Echiburú Fuenzalida.

| Carpeta | Contenido | Estado |
|---|---|---|
| `AVANCE_1/` | Presentación v5 y guion v2: versión final presentada el 12/09/2026 (semana 5) | Cerrado |
| `AVANCE_2/` | Presentación v4 (vigente) con su guion de 10 minutos; v3, v2 y v1 como historial, sobre la plantilla oficial «Avance 2 AFE 2017» (semana 10) | En revisión |

Las versiones intermedias del Avance 1 (v1 a v4) quedan en la rama `PRESENTACION_TFG`.

## Avance 2 · v4

Misma estructura y datos que la v3 (19 láminas), con dos criterios de presentación:

- **Títulos de resultado.** Cada lámina de resultados se titula «Resultado: …» con lo que muestra (por ejemplo, «Resultado: un ciclo de gestión único en MP21»), y lleva sobre el título la etiqueta del objetivo: «Desarrollo y resultados del OE1 · Arquitectura de procesos» u «OE2 · Evaluación de riesgos».
- **Sin referencias a versiones.** Se presenta una única versión, así que la presentación no compara con versiones anteriores de sí misma ni de los manuales. La lámina de comparación de manuales se reemplazó por la apertura de MP21 con lo que hace cada proceso, y la tabla de equivalencias del ciclo por «El contrato dentro del ciclo», que ubica cada hito del contrato y marca las interfaces con MP19.

Guion: `Guion_Avance_2_v4.md`. Generador: `gen_av4.py`.

## Avance 2 · v3

19 láminas. Incorpora el manual **MP21 v2** (rama `MANUAL-MP21` del repositorio `MANUALES`). El OE1 pasa de 4 a 10 láminas:

| OE1 | Lámina |
|---|---|
| 1 | Elementos por nivel MGO, con MP21 completo en sus cinco niveles |
| 2 y 3 | MP8: ejemplo C.1 y SIPOC (sin cambios) |
| 4 | **Ajuste metodológico** en MP21: estructura acordada con el dueño en lugar del flujo único, con equivalencias APQC PCF 8.0 |
| 5 | Reestructuración v1 → v2: de la vida de un contrato a un ciclo de gestión del servicio |
| 6 | Ciclo A → B → MP19 → C → D con retroalimentación, y dónde quedó cada elemento del v1 |
| 7 | SIPOC de MP21 v2 y alcance en las cinco plantas, incluida Licancel |
| 8 y 9 | Ejemplos C.1 Aseguramiento del servicio y C.2 Control y mejora |
| 10 | Estado del contenido: MP8 frente a MP21 |

También se actualizan el OE2 (universo de 248 tareas, ejemplo de fila de MP21) y la carta Gantt (A1.4.1 y A1.4.2 completadas; 5 de 8 actividades del OE1 completadas).

| Nivel | MP8 v1 | MP21 v2 |
|---:|---:|---:|
| Procesos | 5 | 4 |
| Subprocesos | 17 | 8 |
| Actividades (levantadas / propuestas) | 59 (37 / 22) | 24 (11 / 13) |
| Tareas (levantadas / propuestas) | 192 (171 / 21) | 56 (13 / 43) |
| Tareas con responsable | 83 | 53 |

Regenerar: `python3 stats.py mp21_v2.pptx stats_mp21_v2.json`, `python3 capturas_v3.py ../capturas` y `python3 gen_av3.py salida.pptx` (usa `gen_av2.py` para las láminas sin cambios).

## Avance 2 · v2

Portada, título, problema (2), objetivos (1), desarrollo y resultados del OE1 (4), del OE2 (2), carta Gantt con nivel de cumplimiento y cierre.

Cambios respecto de la v1: el marco teórico se reemplaza por una lámina única de objetivos, porque el trabajo ya está aprobado; el SIPOC de MP8 y la ficha de MP21 pasan de anexos a resultados del OE1, con señalizaciones.

Los datos provienen de los manuales del repositorio `MANUALES` (ramas `MANUAL-MP8` y `MANUAL-MP21`, v1):

| Nivel | MP8 | MP21 |
|---|---:|---:|
| Procesos | 5 | 4 |
| Subprocesos | 17 | 11 |
| Actividades | 59 | por levantar |
| Tareas | 192 | por levantar |

En MP8, 171 tareas están levantadas y 21 son propuestas por validar (texto naranjo en el manual), y 83 tienen responsable identificado.

## Regenerar

```
cd AVANCE_2/generador
python3 stats.py mp8.pptx stats_mp8.json    # y mp21
python3 capturas.py ../capturas             # capturas de los manuales
python3 gen_av2.py salida.pptx
python3 render.py salida.pptx html && node check.js html   # control de desbordes y solapes
```

Los scripts esperan, en la carpeta raíz de trabajo, la plantilla `plantilla_av2.pptx`, los manuales `mp8.pptx` y `mp21.pptx` y la carpeta `capturas/`.
