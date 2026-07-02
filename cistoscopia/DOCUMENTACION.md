# Aplicación de Dictado de Cistoscopía
## Dr. Juan Carlos Riera Medina — Urología Ambulatoria, Hospital Clínico Viña del Mar

---

## 1. Descripción General

Aplicación web de una sola página (`index.html`) que permite dictar un informe de uretrocistoscopía completamente **manos libres**, diseñada para uso intraoperatorio con guantes estériles. Reconoce el campo destino por voz, rellena automáticamente la plantilla institucional y genera el informe final en formato `.docx` listo para imprimir o archivar.

**No requiere instalación, servidor ni conexión a internet.**  
Basta con abrir `index.html` en **Google Chrome** o **Microsoft Edge**.

---

## 2. Arquitectura

```
cistoscopia/
├── index.html               ← App completa (autocontenida)
├── plantilla_template.docx  ← Plantilla con marcadores {{CAMPO}}
├── plantilla_original.docx  ← Plantilla original del Dr. Riera (referencia)
└── plantilla_raw/           ← Plantilla original descomprimida (referencia)
    └── word/
        ├── document.xml
        └── media/           ← Logo e imágenes del encabezado
```

### Dependencias embebidas en index.html
| Librería | Versión | Uso |
|---|---|---|
| JSZip | 3.10.1 | Leer y escribir archivos .docx (ZIP) |
| FileSaver.js | 2.0.5 | Descargar el archivo generado |

Ambas están **incrustadas directamente** en el HTML (no CDN) para funcionamiento offline.

---

## 3. Tecnologías

| Tecnología | Rol |
|---|---|
| HTML5 / CSS3 / JavaScript (vanilla) | Interfaz y lógica de la app |
| Web Speech API (`SpeechRecognition`) | Reconocimiento de voz en tiempo real |
| OOXML / ZIP | Formato interno de los archivos .docx |
| Python 3 + python-docx | Generación de la plantilla `.docx` con marcadores |

---

## 4. Flujo de Uso

```
1. Abrir index.html en Chrome/Edge
2. Permitir acceso al micrófono cuando el navegador lo solicite
3. Presionar "Iniciar dictado"
4. Dictar cada campo en formato:
      "[nombre del campo] [contenido]"
   Ejemplo: "uretra sin alteración mucosa normal"
5. El campo se rellena automáticamente (destello verde de confirmación)
6. Repetir para cada sección del informe
7. Presionar "Descargar informe .docx"
8. El archivo descargado tiene el layout institucional completo
```

---

## 5. Campos del Informe

| Campo | ID interno | Palabras clave reconocidas |
|---|---|---|
| Nombre del paciente | `nombre` | nombre, paciente |
| RUT | `rut` | rut, run, ruth, root, runt, ruta |
| Edad | `edad` | edad |
| Fecha | `fecha` | fecha |
| Indicación / Diagnóstico ingreso | `indicacion` | indicación, indicacion, motivo, diagnóstico de ingreso |
| Uretra anterior | `uretra` | uretra, uretra anterior |
| Uretra posterior / Próstata | `prostata` | próstata, prostata, uretra posterior |
| Cuello vesical | `cuello` | cuello, cuello vesical |
| Vejiga | `vejiga` | vejiga |
| Diagnóstico endoscópico | `diagnostico` | diagnóstico, diagnostico, diagnóstico endoscópico, conclusión |
| Conducta / Sugerencia | `conducta` | conducta, sugerencia, plan, tratamiento |

---

## 6. Detección de Campos por Voz

### Algoritmo (`detectarCampo`)

```javascript
function detectarCampo(texto) {
  // 1. Normalizar: minúsculas, sin tildes, sin caracteres especiales
  const norm = normalizar(texto);

  // 2. Buscar todas las palabras clave que sean prefijo de la frase
  const candidatos = [];
  for (const campo of CAMPO_MAP) {
    for (const key of campo.keys) {
      if (norm.startsWith(normKey + ' ') || norm === normKey) {
        candidatos.push({ campo, len: normKey.length });
      }
    }
  }

  // 3. Elegir el match más largo (evita conflictos entre claves similares)
  candidatos.sort((a, b) => b.len - a.len);
  const best = candidatos[0];

  // 4. Extraer el contenido (todo lo que sigue al nombre del campo)
  const contenido = texto.slice(best.len).replace(/^[\s:,]+/, '').trim();
  return { id: best.campo.id, contenido };
}
```

### Normalización
```javascript
function normalizar(s) {
  return s.toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')   // eliminar tildes/diacríticos
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
```

### Manejo de resultados de voz
```javascript
recognition.onresult = (e) => {
  for (let i = e.resultIndex; i < e.results.length; i++) {
    const t = e.results[i][0].transcript;
    if (e.results[i].isFinal) {
      // IMPORTANTE: usar solo t (el resultado final completo)
      // NO concatenar con texto interino — causa duplicación
      const frase = t.trim();
      const det = detectarCampo(frase);
      if (det && det.contenido) {
        document.getElementById(det.id).value = det.contenido;
        flashCampo(det.id);  // destello verde
      }
    } else {
      // Mostrar texto interino en el panel de visualización
      txDiv.textContent = t;
    }
  }
};
```

**Bug corregido (v3):** versiones anteriores usaban `fraseActual + t` para el resultado final, lo que causaba duplicación porque `t` ya contiene el texto completo y `fraseActual` tenía el mismo texto como interino.

---

## 7. Generación del .docx

### Cómo funciona
1. La `plantilla_template.docx` está codificada en **Base64** e incrustada en el HTML.
2. Al descargar, JSZip la decodifica y abre como ZIP.
3. Se extrae `word/document.xml` y se hacen reemplazos de texto:
   ```javascript
   xml = xml.replaceAll('{{NOMBRE}}', esc(datos.nombre));
   xml = xml.replaceAll('{{RUT}}',    esc(datos.rut));
   // ... etc para cada campo
   ```
4. El XML modificado se re-empaqueta como ZIP y se descarga como `.docx`.

### Función de escape XML
```javascript
function esc(s) {
  return (s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
```

### Nombre del archivo generado
```
cistoscopia_[Apellido1_Apellido2]_[YYYYMMDD].docx
```

---

## 8. Plantilla (.docx)

La plantilla fue construida con **python-docx** desde cero, replicando el estilo visual de la plantilla original del Dr. Riera pero usando **párrafos y tablas normales** (no cuadros de texto flotantes), lo que garantiza el relleno limpio sin duplicaciones.

### Estructura del documento generado

```
┌─────────────────────────────────────────────────┐
│  [LOGO]    HOSPITAL CLINICO VIÑA DEL MAR        │
│            UNIDAD DE UROLOGIA AMBULATORIA       │
│            DR. JUAN CARLOS RIERA MEDINA         │
│            UROLOGO                              │
├─────────────────────────────────────────────────┤
│ Fecha: {{FECHA}}                                │
│                                                 │
│         INFORME URETROCISTOSCOPIA               │
│                                                 │
│  Nombre │ {{NOMBRE}}                            │
│  RUT    │ {{RUT}}                               │
│  Edad   │ {{EDAD}}                              │
├─────────────────────────────────────────────────┤
│ INDICACIÓN                                      │
│   {{INDICACION}}                                │
│                                                 │
│ Anestesia: Intrauretral (sí). Equipo: ...       │
├─────────────────────────────────────────────────┤
│ URETRA ANTERIOR                                 │
│   {{URETRA}}                                    │
│ URETRA POSTERIOR / PRÓSTATA                     │
│   {{PROSTATA}}                                  │
│ CUELLO VESICAL                                  │
│   {{CUELLO}}                                    │
│ VEJIGA                                          │
│   {{VEJIGA}}                                    │
├─────────────────────────────────────────────────┤
│ DIAGNÓSTICO ENDOSCÓPICO: {{DIAGNOSTICO}}        │
│ SUGERENCIA: {{CONDUCTA}}                        │
├─────────────────────────────────────────────────┤
│                    Dr. Juan Carlos Riera Medina │
│                                       Urólogo  │
└─────────────────────────────────────────────────┘
```

### Script de generación de la plantilla
```python
# Requiere: pip install python-docx
from docx import Document
from docx.shared import Pt, Cm, RGBColor

doc = Document()

# Márgenes: 2cm lados, 1.5cm arriba, 1.8cm abajo
section = doc.sections[0]
section.left_margin = section.right_margin = Cm(2.0)
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.8)

# Encabezado: tabla [logo | info hospital]
hdr = doc.add_table(rows=1, cols=2)
hdr.cell(0,0).paragraphs[0].add_run().add_picture('logo.png', width=Cm(4.5))
for linea in ['HOSPITAL CLINICO VIÑA DEL MAR', 'UNIDAD DE UROLOGIA AMBULATORIA',
              'DR. JUAN CARLOS RIERA MEDINA', 'UROLOGO']:
    p = hdr.cell(0,1).add_paragraph(linea)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.runs[0].bold = True

# Campos con marcadores
doc.add_paragraph().add_run('Fecha: ').bold = True
doc.add_paragraph('{{FECHA}}')
# ... etc.

doc.save('plantilla_template.docx')
```

---

## 9. Configuración del Reconocimiento de Voz

```javascript
recognition = new SpeechRecognition();
recognition.lang = 'es-CL';        // Español de Chile
recognition.continuous = true;     // No se detiene entre frases
recognition.interimResults = true; // Muestra texto mientras se habla
recognition.maxAlternatives = 1;   // Solo la alternativa más probable
```

**Compatibilidad:** Chrome 33+, Edge 79+. No funciona en Firefox ni Safari.  
**Reinicio automático:** si el reconocedor se detiene (timeout), se reinicia automáticamente mientras el micrófono esté activo.

---

## 10. Historial de Versiones

| Versión | Commit | Cambio |
|---|---|---|
| v1 | `1fb1edc` | App inicial con dictado → selector manual → aplicar |
| v2 | `1d7a759` | Plantilla original integrada como base64 |
| v3 | `dbe692d` | JSZip y FileSaver embebidos (sin CDN, funciona offline) |
| v4 | `463c2a2` | Detección automática de campo por voz (manos libres) |
| v5 | `388f949` | Plantilla nueva desde cero + corrección "Ruth"→RUT |
| v6 | `6c2ed91` | **Corrección duplicación:** usar `t.trim()` no `fraseActual+t` |

---

## 11. Problemas Conocidos y Soluciones

### Duplicación de texto
**Causa:** `fraseActual + t` en el handler `onresult`. El motor de voz entrega primero el texto como *interino* y luego como *final* — ambos con el mismo contenido. Sumarlos duplica el texto.  
**Solución:** usar solo `t.trim()` para resultados finales.

### RUT reconocido como "Ruth" / "root" / "zona horaria"
**Causa:** el motor de voz en español no reconoce el acrónimo "RUT".  
**Solución:** mapa de alias → `['rut','run','ruth','root','runt','ruta']`.

### Texto desalineado en el .docx
**Causa:** la plantilla original usaba cuadros de texto flotantes (`wps:txbx`) superpuestos sobre párrafos normales. Al reemplazar solo los cuadros, los párrafos quedaban con los valores anteriores, causando duplicación visual y desalineamiento.  
**Solución:** plantilla nueva con tablas y párrafos estándar.

### App no funciona en Firefox / Safari
**Causa:** `SpeechRecognition` no está implementada en esos navegadores.  
**Solución:** usar Chrome o Edge. Los campos se pueden rellenar manualmente escribiendo.

---

## 12. Cómo Actualizar el Logo

1. Reemplazar `plantilla_raw/word/media/image1.png` con el nuevo logo.
2. Ejecutar el script de generación de plantilla (sección 8).
3. Codificar el nuevo `.docx` en base64 e incrustar en `index.html`:
   ```python
   import base64
   with open('plantilla_template.docx', 'rb') as f:
       b64 = base64.b64encode(f.read()).decode()
   # Pegar b64 en TEMPLATE_B64 dentro de index.html
   ```

---

## 13. Cómo Agregar un Nuevo Campo

**En el HTML** — agregar el `<textarea>` o `<input>` con el ID del nuevo campo.

**En el JavaScript** — agregar al array `CAMPO_MAP`:
```javascript
{ id: 'nuevo_campo', keys: ['palabra clave', 'alias'] },
```

**En la plantilla Python** — agregar el marcador `{{NUEVO_CAMPO}}` en el lugar correcto del documento.

**En la función de generación** — agregar la entrada al objeto de reemplazos:
```javascript
'{{NUEVO_CAMPO}}': esc(d.nuevo_campo),
```

---

## 14. Repositorio

- **Repo:** `jcrieram/ProjectsJCRM`
- **Branch de desarrollo:** `claude/cystoscopy-dictation-app-lcsw7f`
- **Archivos principales:**
  - `cistoscopia/index.html` — aplicación completa
  - `cistoscopia/plantilla_template.docx` — plantilla con marcadores
  - `cistoscopia/plantilla_original.docx` — original de referencia
