#!/usr/bin/env python3
from servidor_dashboard_supabase import Handler, HOST, PORT
from http.server import ThreadingHTTPServer

if __name__ == "__main__":
    print(f"Planeador Supabase disponible en {HOST}:{PORT}", flush=True)
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    try:
        server.serve_forever()
    finally:
        server.server_close()
