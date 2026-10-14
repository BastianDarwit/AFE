# AFE · Presentaciones de avance del TFG

Rama única para los hitos de avance. Cada hito tiene su carpeta y sus versiones se numeran dentro de ella.

**Trabajo:** Propuesta de un sistema de gestión de riesgos para los macroprocesos de Gestión de Personas (MP8) y Servicios Generales (MP21) de Celulosa Arauco y Constitución S.A.
**Estudiante:** Bastián Darwit Vera · **Profesor guía:** Carlos Muñoz P. · **Profesor co-guía:** Manuel Echiburú Fuenzalida.

| Carpeta | Contenido | Estado |
|---|---|---|
| `AVANCE_1/` | Presentación v5 y guion v2: versión final presentada el 12/09/2026 (semana 5) | Cerrado |
| `AVANCE_2/` | Presentación v2 (vigente) y v1, sobre la plantilla oficial «Avance 2 AFE 2017» (semana 10) | En revisión |

Las versiones intermedias del Avance 1 (v1 a v4) quedan en la rama `PRESENTACION_TFG`.

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
