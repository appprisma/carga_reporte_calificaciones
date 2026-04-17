# Estructura de Hojas — Google Sheets

Esta es la guía de las hojas que deben existir en el Spreadsheet antes de usar la plataforma.

---

## BASE MAESTROS

Contiene el catálogo de profesores.

| Columna | Header exacto | Ejemplo |
|---------|--------------|---------|
| A | `MATRICULA PROFESOR` | `12345` |
| B | `NOMBRE DEL PROFESOR` | `GARCIA LOPEZ JUAN` |

---

## BASE ALUMNOS

Catálogo de alumnos del instituto.

| Columna | Contenido | Ejemplo |
|---------|-----------|---------|
| A | Matrícula alumno | `A001` |
| B | Nombre completo | `Pérez Torres Ana` |
| C | Tutor | `Pérez García Roberto` |
| D | Grupo español | `6A` |
| E | Grupo inglés | `KEY1` |
| F | Correo institucional | `ana.perez@ibime.edu.mx` |

---

## GRUPOS

Lista de grupos activos.

| Columna A: `GRUPO ESPAÑOL` | Columna B: `GRUPO INGLES` |
|---------------------------|--------------------------|
| 1A | KEY1 |
| 1B | KEY2 |
| 2A | PET1 |
| … | … |

---

## ASIGNATURA

Catálogo de materias por idioma.

| Columna A: `ESPAÑOL` | Columna B: `INGLES` |
|---------------------|---------------------|
| Matemáticas | English Grammar |
| Español | Reading Comprehension |
| Ciencias Naturales | Writing |
| … | … |

---

## Matriz_Resumen

Es la hoja central de asignaciones. Define qué grupos y materias le corresponden a cada profesor.

- **Fila 1**: Nombre de cada profesor (igual que en `BASE MAESTROS > NOMBRE DEL PROFESOR`, en minúsculas está bien)
- **Columna A** (desde fila 2): Nombre de cada grupo
- **Celdas interiores**: Nombre de la materia (o materias, separadas por salto de línea `Alt+Enter`)

Ejemplo:

|  | Garcia Lopez Juan | Ramirez Mendez Maria |
|--|------------------|---------------------|
| 1A | Matemáticas | |
| 1B | | Español |
| 2A | Matemáticas | Ciencias Naturales |

---

## Hojas generadas automáticamente

Estas hojas **no necesitas crearlas** — el sistema las genera en el primer uso:

| Hoja | Cuándo se crea | Qué contiene |
|------|---------------|--------------|
| `Maestro` | Primer import | Todos los registros de calificaciones importados |
| `TABLERO_AVANCE` | Primer import | Progreso de entregas por docente y grupo |
| `DASHBOARD_ENTREGAS` | Primer import | Estatus de cada docente (✅ / ⏳ / ❌) |
| `BITACORA_TIEMPOS` | Primer import | Historial de cada importación con tiempos |

---

## Formato del Sheet de Classroom (exportado por el profesor)

El Sheet que el profesor pega como link debe seguir el formato estándar de exportación de Google Classroom:

| Fila | Contenido |
|------|-----------|
| 1 | Fechas de cada actividad |
| 2 | Nombre de cada actividad |
| 3 | Calificación máxima de cada actividad (10 o 100) |
| 4+ | Una fila por alumno: nombre, correo `@ibime.edu.mx`, calificaciones |

El sistema detecta automáticamente si la escala es 10 o 100 y convierte todo a escala 10.
