# Bright Studio — Sales Toolkit

Portal de marca con generador de presentaciones .pptx.

## Subir a GitHub + Vercel (10 minutos)

### Paso 1 — Sube el proyecto a GitHub

1. Ve a github.com → New repository
2. Nómbralo `bright-toolkit` → Create repository
3. En tu ordenador, abre Terminal y ejecuta:

```bash
cd bright-toolkit
git init
git add .
git commit -m "primer commit"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/bright-toolkit.git
git push -u origin main
```

### Paso 2 — Conecta con Vercel

1. Ve a vercel.com → New Project
2. Importa el repositorio `bright-toolkit` desde GitHub
3. En "Framework Preset" selecciona **Other**
4. Deja todo lo demás por defecto → Deploy

### Paso 3 — Ya está

Vercel te da una URL tipo `bright-toolkit.vercel.app`. 
Cada vez que hagas push a GitHub, Vercel actualiza automáticamente.

## Estructura del proyecto

```
bright-toolkit/
├── api/
│   ├── generate.py      ← función que genera el .pptx
│   └── template.pptx    ← template de Gabi
├── public/
│   └── index.html       ← el portal
├── vercel.json          ← configuración de Vercel
└── requirements.txt     ← dependencias Python
```

## Actualizar el template de Gabi

Cuando Gabi tenga una nueva versión del template:
1. Reemplaza `api/template.pptx` con el nuevo archivo
2. Haz `git add . && git commit -m "nuevo template" && git push`
3. Vercel lo despliega automáticamente en 30 segundos
