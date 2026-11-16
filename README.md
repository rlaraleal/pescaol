# PescAOL Web 1.1

Paquete listo para publicación gratuita mediante GitHub Pages.

## Publicar gratis

1. Cree una cuenta en https://github.com/ si no dispone de una.
2. Cree un repositorio nuevo llamado `pescaol` y seleccione **Public**.
3. Descomprima este ZIP.
4. En el repositorio, use **Add file > Upload files** y cargue **todo el contenido**, incluida la carpeta `.github`.
5. Confirme los archivos en la rama `main`.
6. Abra **Settings > Pages**.
7. En **Build and deployment > Source**, seleccione **GitHub Actions**.
8. Abra la pestaña **Actions** y espere a que termine “Publicar PescAOL en GitHub Pages”.
9. La URL quedará con esta forma: `https://USUARIO.github.io/pescaol/`.

## Teléfonos y computadores

- Abra la URL pública en Chrome, Edge, Safari o Firefox.
- Android/Chrome: menú `⋮` > **Agregar a pantalla principal** o **Instalar aplicación**.
- iPhone/Safari: **Compartir** > **Añadir a pantalla de inicio**.
- Windows/Edge o Chrome: icono **Instalar aplicación** en la barra de direcciones.

## Actualizar con un Excel nuevo

Cuando se genere una nueva versión desde un Excel AOL:

1. Reemplace `index.html` en el mismo repositorio.
2. Confirme el cambio en `main`.
3. GitHub Pages publicará automáticamente la actualización en la misma URL.
4. Los usuarios deben recargar la página. El service worker usa una versión de caché que debe incrementarse en cada entrega.

## Seguridad de publicación

La versión pública no debe contener RUT en texto claro, patentes, choferes, documentos tributarios ni observaciones internas. En este paquete, los identificadores usados por la lógica fueron seudonimizados.

## Limitación cartográfica

La interfaz usa capas gratuitas de calles. Los proveedores gratuitos pueden aplicar límites o bloqueos por tráfico. Para producción institucional de alto uso, debe utilizarse una capa autorizada o un servicio cartográfico contratado.
