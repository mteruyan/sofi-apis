def safe_read_csv(url, tries=3, delay=2):
    import subprocess
    import time

    for intento in range(1, tries + 1):
        try:
            print(f"[CSV] Intento {intento}/{tries}")
            print(f"[CSV] URL: {url}")

            resultado = subprocess.run(
                [
                    "curl",
                    "-L",
                    "--max-time", "60",
                    "-sS",
                    url
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace"
            )

            if resultado.returncode == 0 and resultado.stdout:
                lines = resultado.stdout.splitlines()

                print(f"[CSV] Descarga OK")
                print(f"[CSV] Líneas recibidas: {len(lines)}")
                print(f"[CSV] Tamaño: {len(resultado.stdout)} caracteres")

                return lines

            print(
                f"[CSV] Error curl. Código: "
                f"{resultado.returncode}"
            )

            if resultado.stderr:
                print(
                    f"[CSV] STDERR: {resultado.stderr}"
                )

        except Exception as e:
            print(
                f"[CSV] Error inesperado: "
                f"{type(e).__name__}: {e}"
            )

        if intento < tries:
            print(f"[CSV] Esperando {delay} segundos...")
            time.sleep(delay)

    print("[CSV] No se pudo descargar el CSV.")

    return None

url = (
    "https://docs.google.com/spreadsheets/d/"
    "1dJdYBQqzjilP7SXLsb6OSgi_wOqn8DrxfiQ6-OtNcM8"
    "/gviz/tq?tqx=out:csv"
    "&gid=577147525"
)

lines = safe_read_csv(url)

if lines:
    print("FUNCIONA")
    print("Total líneas:", len(lines))
    print("Primera línea:", lines[0])
else:
    print("NO FUNCIONA")
