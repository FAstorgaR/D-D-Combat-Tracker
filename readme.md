# D&D Combat Tracker

Aplicación desarrollada en Python y Streamlit para gestionar encuentros de Dungeons & Dragons 5e.

Permite administrar jugadores y enemigos, controlar iniciativa, HP, condiciones, rondas de combate y consultar información detallada de criaturas desde un bestiario en Excel.

---

## Características

### Gestión de iniciativa

* Orden automático por iniciativa.
* Modificadores de iniciativa configurables.
* Reordenamiento manual para resolver empates.
* Seguimiento del turno activo.
* Control de rondas.

### Gestión de criaturas

* Agregar jugadores manualmente.
* Agregar enemigos desde un bestiario.
* Duplicar criaturas.
* Eliminar participantes.
* Visualización diferenciada para jugadores y enemigos.

### Gestión de combate

* Daño directo.
* Curación.
* Puntos de golpe temporales.
* Indicadores visuales de estado.
* Barra de vida dinámica.
* Marcado automático de criaturas caídas.

### Condiciones

Incluye soporte para condiciones comunes de D&D 5e:

* Grappled
* Restrained
* Frightened
* Poisoned
* Stunned
* Charmed
* Paralized
* Exhaustion
* y otras.

Las condiciones tienen duración configurable y disminuyen automáticamente al finalizar turnos.

### Bestiario

Las criaturas se cargan desde un archivo Excel:

```text
Bestiario v2.xlsx
```

Información soportada:

* Nombre
* CA
* HP
* Velocidad
* Iniciativa
* Habilidades
* Resistencias
* Sentidos
* Rasgos
* Acciones
* Acciones adicionales
* Reacciones
* Acciones legendarias
* Atributos

### Guardado

* Autoguardado automático.
* Guardado manual de encuentros.
* Restauración de autoguardado.
* Carga de encuentros previos.

---

## Requisitos

Python 3.10+

Dependencias principales:

```bash
streamlit
pandas
openpyxl
```

Instalación:

```bash
pip install streamlit pandas openpyxl
```

---

## Ejecución

NO ejecutar con:

```bash
python Combat_Tracker.py
```

Este proyecto utiliza Streamlit.

Ejecutar con:

```bash
streamlit run Combat_Tracker.py
```

Después abrir el navegador en:

```text
http://localhost:8501
```

---

## Estructura esperada

```text
CombatTracker/
│
├── Combat_Tracker.py
├── Bestiario v2.xlsx
│
└── encuentros/
    ├── _autosave.json
    └── *.json
```

---

## Flujo de uso

### Preparar encuentro

1. Seleccionar CR.
2. Buscar criatura.
3. Elegir cantidad.
4. Agregar enemigos.

### Agregar jugadores

1. Nombre.
2. CA.
3. HP.
4. Iniciativa.

### Iniciar combate

1. Ordenar iniciativa.
2. Revisar turno activo.
3. Aplicar daño o curación.
4. Avanzar turno.

### Guardar encuentro

Utilizar:

```text
Guardar encuentro
```

para almacenar el estado completo de la sesión.

---

## Tecnologías utilizadas

* Python
* Streamlit
* Pandas
* OpenPyXL
* JSON

---

## Autor

Proyecto personal para gestión de encuentros de Dungeons & Dragons 5e.

Desarrollado para uso en mesa presencial y apoyo a campañas con grandes cantidades de criaturas.

