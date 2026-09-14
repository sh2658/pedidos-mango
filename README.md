# Pedidos Mango

Aplicación web liviana para recibir pedidos en línea, con catálogo y operación administrados desde Google Sheets.

## Funcionalidades

- Catálogo dinámico de productos disponibles.
- Carrito persistente en el navegador.
- Registro y reutilización de datos del cliente.
- Captura opcional de ubicación para organizar entregas.
- Validación de precios y productos en el servidor.
- Código único y consulta del estado del pedido.
- Apertura/cierre de pedidos, monto mínimo y WhatsApp configurables.
- Resumen operativo y ruta optimizada desde Apps Script.

## Estructura

- `index.html`: interfaz pública, responsive y sin dependencias externas.
- `logo-ceviche-mango.jpeg`: identidad visual.
- `backend/Codigo.gs`: código fuente del servicio de Google Apps Script.

## Configuración operativa

La hoja `Config` admite estas claves: `FechaCorte`, `ZonaLat`, `ZonaLng`, `RadioKm`, `MontoMinimo`, `EmailNotificacion`, `WhatsAppNegocio` y `PedidosAbiertos`.

La clave de Google Maps, si se usa la ruta optimizada, debe guardarse en las propiedades del script con el nombre `MAPS_API_KEY`; no debe subirse al repositorio.

