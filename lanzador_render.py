#!/usr/bin/env python3
"""Punto de entrada para publicar el Planeador V2 en Render.

No modifica la logica existente del servidor. Importa el Handler y los
recursos de servidor_dashboard_v2.py, pero escucha en 0.0.0.0 y en el puerto
que Render entrega mediante la variable de entorno PORT.
"""
import os
from http.server import ThreadingHTTPServer

import servidor_dashboard_v2 as app

HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "10000"))


def main():
    # Conserva las rutas, Excel, importacion, exportacion y recalculo del
    # servidor actual. Solamente cambia la interfaz y el puerto de escucha.
    server = ThreadingHTTPServer((HOST, PORT), app.H)
    print(f"Planeador V2 disponible en {HOST}:{PORT}", flush=True)
    try:
        server.serve_forever()
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
