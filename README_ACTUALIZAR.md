# PescAOL Web 1.2

## Botón Instalar PescAOL
En Chrome y Edge aparecerá cuando el navegador emita el evento de instalación. En iPhone, el botón explica la ruta Safari > Compartir > Añadir a pantalla de inicio.

## Sustituir archivos en GitHub
Sube `index.html`, `manifest.webmanifest`, `service-worker.js` e `icon.svg` a la raíz del repositorio, reemplazando los actuales. Mantén `.nojekyll`.

## Consolidar un Excel nuevo
Con Python y pandas instalados: `python consolidar_excel.py archivo.xlsx salida_pescaol`. El proceso genera dirección maestra, variantes y cola de geocodificación.

## Marcadores
Esta versión no usa agrupación por zoom. Cada local es un `L.circleMarker` vectorial persistente; acercar o alejar no elimina puntos. Los puntos cercanos pueden superponerse visualmente, pero no desaparecen del mapa.
