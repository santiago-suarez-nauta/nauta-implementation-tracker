# Nauta Implementation Tracker

Static dashboard temporal — deploy en GitHub Pages.

**URL pública:** https://santiago-suarez-nauta.github.io/nauta-implementation-tracker/

## Cómo actualizar

```bash
# 1. Exportar el tracker de Google Sheets como .xlsx
# 2. Reemplazar data/Nauta_Implementation_Tracker.xlsx
# 3. Generar HTML + push:
bash deploy.sh
```

GitHub Pages publica en ~30-60 segundos después del push.

## Estructura

```
nauta-dashboard/
├── generate.py                 # lee xlsx → index.html (filtra owner=ss)
├── deploy.sh                   # genera + git push
├── index.html                  # output (servido por GitHub Pages)
├── data/
│   └── Nauta_Implementation_Tracker.xlsx
└── assets/
```

## Filtros aplicados

- Solo clientes con `owner = "ss"` o `"SS"` en la hoja **Resumen Clientes**.
- Todas las etapas (1-8) se incluyen.
- Tareas se asocian al cliente vía nombre exacto.

## Brand

- Tipografía: **Manrope** (Google Fonts, weights 400/500/600/700/800).
- Color primario: **#4A6CF7** (Nauta Blue).
- Color dark: **#0F1B3D**.
