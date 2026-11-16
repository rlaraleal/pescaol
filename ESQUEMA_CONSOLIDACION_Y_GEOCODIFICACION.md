# Consolidación de direcciones y geocodificación

## Dirección maestra por destinatario
La clave interna es el RUT normalizado; para publicación se usa `establecimientoId` seudonimizado. Las variantes se normalizan, agrupan y puntúan por presencia de número, tipo de vía, coincidencia de comuna, frecuencia y fecha reciente.

## Regla de precedencia
1. Coordenada verificada en terreno.
2. Coordenada validada administrativamente.
3. Dirección consolidada geocodificada con coincidencia de comuna.
4. Si no hay coincidencia confiable: `PENDIENTE`, sin publicar punto.

## Cola de geocodificación
`cola_geocodificacion.csv` contiene una consulta por destinatario: `dirección consolidada + comuna + Región de Ñuble + Chile`.

## Validación de respuesta
Aceptar solo cuando: país=Chile; región=Ñuble; comuna coincide; latitud entre -37.5 y -35.7; longitud entre -73.2 y -71.2; y el resultado no sea solo nivel región/comuna. Guardar proveedor, dirección formateada, latitud, longitud, precisión, fecha y estado.

## Google Maps
Google Geocoding o Address Validation pueden aproximar los resultados de Google Maps, pero requieren proyecto, clave, facturación y cumplimiento de sus términos. No se deben copiar coordenadas manualmente de Google Maps ni dejar una clave privada dentro de un repositorio público.

## Publicación
La aplicación debe consumir solo `establecimientoId`, nombre público, dirección consolidada, comuna, latitud, longitud, precisión y la AOL visible. Nunca publicar RUT, chofer, patente ni documentos tributarios.
