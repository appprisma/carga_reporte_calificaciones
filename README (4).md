# IBIME — Sistema de Importación de Calificaciones

Web App construida con Google Apps Script que permite a los profesores importar calificaciones desde Google Classroom hacia un Spreadsheet maestro y sincronizarlas con Firebase Realtime Database.

---

## Arquitectura

```
Google Classroom (Sheets exportados)
         │
         ▼
  WebApp.html  ←→  Code.gs          ← GAS Web App
                     │
                     ├─ Sheets (Maestro / BASE MAESTROS / Matriz_Resumen)
                     │        (TABLERO_AVANCE / DASHBOARD_ENTREGAS)
                     │
                     └─ SyncToFirebase.gs → Firebase Realtime DB
```

---

## Estructura del repositorio

```
ibime-calificaciones/
├── src/
│   └── gas/
│       ├── appsscript.json   # Manifiesto GAS
│       ├── Code.gs           # Lógica principal + Web App
│       ├── SyncToFirebase.gs # Sincronización horaria a Firebase
│       └── WebApp.html       # Interfaz de usuario
├── docs/
│   └── estructura-sheets.md  # Descripción de las hojas requeridas
├── .clasp.json               # Config de la CLI clasp (gitignore si repo público)
├── .gitignore
└── README.md
```

---

## Requisitos previos

- Cuenta de Google Workspace (o Gmail para pruebas)
- [Node.js](https://nodejs.org) instalado (para usar `clasp`)
- Proyecto en Firebase con Realtime Database habilitada
- Google Spreadsheet con las hojas descritas en `docs/estructura-sheets.md`

---

## Setup inicial

### 1. Instalar clasp

```bash
npm install -g @google/clasp
clasp login
```

### 2. Clonar este repo y configurar clasp

```bash
git clone https://github.com/TU_USUARIO/ibime-calificaciones.git
cd ibime-calificaciones
```

Edita `.clasp.json` y reemplaza `TU_SCRIPT_ID_AQUI` con el Script ID de tu proyecto de Apps Script.  
Lo encuentras en: tu Spreadsheet → **Extensions > Apps Script > Project Settings > Script ID**

### 3. Guardar las credenciales de Firebase (solo una vez)

Desde el editor de GAS, ejecuta la función `setup_guardarSecrets()` en `Code.gs` con tus valores reales.  
**Después de ejecutarla, puedes dejarla comentada.** Las credenciales quedan guardadas en `PropertiesService` y **nunca se suben al repositorio**.

```javascript
function setup_guardarSecrets() {
  PropertiesService.getScriptProperties().setProperties({
    FIREBASE_URL   : 'https://TU-PROYECTO-default-rtdb.firebaseio.com/',
    FIREBASE_SECRET: 'TU_SECRET_AQUI'
  });
}
```

### 4. Subir el código al proyecto de GAS

```bash
cd ibime-calificaciones
clasp push
```

### 5. Deploy como Web App

Desde el editor de GAS:
1. **Deploy > New deployment**
2. Tipo: **Web App**
3. Execute as: **Me**
4. Who has access: **Anyone** (o Anyone within your organization)
5. Copiar la URL generada — esa es la URL que comparten los profesores

### 6. Configurar sincronización automática a Firebase

Desde el editor de GAS, ejecuta la función `crearTrigger()` en `SyncToFirebase.gs`.  
Esto creará un trigger que sincroniza los datos cada hora.

---

## Hojas requeridas en Google Sheets

| Hoja | Descripción |
|------|-------------|
| `BASE MAESTROS` | Columnas: `MATRICULA PROFESOR`, `NOMBRE DEL PROFESOR` |
| `BASE ALUMNOS` | Columnas: matrícula, nombre, tutor, grupo español, grupo inglés, correo |
| `GRUPOS` | Columnas: `GRUPO ESPAÑOL`, `GRUPO INGLES` |
| `ASIGNATURA` | Columnas: `ESPAÑOL`, `INGLES` |
| `Matriz_Resumen` | Fila 1 = nombres de profesores, Columna A = grupos, celdas = materias |
| `Maestro` | Hoja de destino — se crea automáticamente |
| `TABLERO_AVANCE` | Se crea automáticamente al primer import |
| `DASHBOARD_ENTREGAS` | Se crea automáticamente al primer import |
| `BITACORA_TIEMPOS` | Se crea automáticamente al primer import |

---

## Flujo de uso para el profesor

1. Abrir la URL de la Web App
2. Ingresar su **matrícula** — el sistema carga automáticamente sus grupos y materias asignadas
3. Pegar el **link de Google Sheet** de Classroom para cada grupo (deben tener acceso compartido)
4. Clic en **IMPORTAR DATOS**
5. El sistema detecta automáticamente si la escala es 10 o 100 y normaliza

---

## Actualizaciones frecuentes

Para actualizar el código después de cambios:

```bash
clasp push
# Si cambiaste el HTML o la lógica, crea un nuevo deployment en GAS
```

Para ver logs de errores:
```bash
clasp logs
```

---

## Seguridad

- Las credenciales de Firebase se almacenan en `PropertiesService` de GAS, nunca en el código
- El `.gitignore` excluye archivos de secrets
- Si el repositorio es público, agrega `.clasp.json` al `.gitignore` para no exponer el `scriptId`

---

## Roadmap sugerido

- [ ] Portal web de consulta para alumnos (consume Firebase)
- [ ] Notificaciones por WhatsApp al cargar calificaciones
- [ ] App móvil (Android/iOS) con acceso por correo institucional
- [ ] Reportes exportables en PDF por grupo
