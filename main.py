import os
import re
import requests
import msal
from config import CLIENT_ID, TENANT_ID, SCOPES, FECHA_INICIO, FECHA_FIN, CARPETA_DESCARGAS


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def limpiar_nombre_archivo(nombre):
    """
    Limpia caracteres inválidos para nombres de archivo/carpeta en Windows.
    También evita nombres terminados en punto o espacio.
    """
    nombre = re.sub(r'[<>:"/\\|?*]', "_", nombre)

    # Quitar saltos de línea y espacios dobles
    nombre = nombre.replace("\n", " ").replace("\r", " ")
    nombre = re.sub(r"\s+", " ", nombre)

    # Windows no permite bien nombres que terminen en punto o espacio
    nombre = nombre.strip().rstrip(".")

    # Evitar nombres vacíos
    if not nombre:
        nombre = "sin_nombre"

    return nombre


def crear_carpeta_si_no_existe(ruta):
    if not os.path.exists(ruta):
        os.makedirs(ruta)


def obtener_token():
    """
    Autenticación con Microsoft Graph usando device code flow.
    El programa mostrará un enlace y un código.
    Tú inicias sesión desde el navegador.
    """
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=authority
    )

    accounts = app.get_accounts()

    if accounts:
        resultado = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        resultado = None

    if not resultado:
        flujo = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flujo:
            raise Exception("No se pudo crear el flujo de autenticación.")

        print("\n=== INICIO DE SESIÓN MICROSOFT ===")
        print(flujo["message"])
        print("==================================\n")

        resultado = app.acquire_token_by_device_flow(flujo)

    if "access_token" in resultado:
        return resultado["access_token"]

    raise Exception(f"Error obteniendo token: {resultado}")


def consultar_correos_con_adjuntos(token):
    """
    Consulta correos entre las fechas indicadas y solo los que tienen adjuntos.
    Maneja paginación automáticamente.
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    filtro = (
        f"receivedDateTime ge {FECHA_INICIO} "
        f"and receivedDateTime le {FECHA_FIN} "
        f"and hasAttachments eq true"
    )

    url = (
        f"{GRAPH_BASE_URL}/me/messages"
        f"?$select=id,subject,receivedDateTime,from,hasAttachments"
        f"&$filter={filtro}"
        f"&$orderby=receivedDateTime asc"
        f"&$top=50"
    )

    correos = []

    while url:
        respuesta = requests.get(url, headers=headers)

        if respuesta.status_code != 200:
            print("Error consultando correos:")
            print(respuesta.status_code)
            print(respuesta.text)
            break

        data = respuesta.json()

        correos.extend(data.get("value", []))

        url = data.get("@odata.nextLink")

    return correos


def obtener_adjuntos(token, message_id):
    """
    Obtiene la lista de adjuntos de un correo.
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/attachments"
    adjuntos = []

    while url:
        respuesta = requests.get(url, headers=headers)

        if respuesta.status_code != 200:
            print(f"Error consultando adjuntos del correo {message_id}:")
            print(respuesta.status_code)
            print(respuesta.text)
            break

        data = respuesta.json()
        adjuntos.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return adjuntos


def descargar_adjunto(token, message_id, attachment_id, nombre_archivo, carpeta_destino):
    """
    Descarga un adjunto usando /$value.
    Microsoft Graph permite obtener el contenido bruto del archivo adjunto con /$value.
    """
    headers = {
        "Authorization": f"Bearer {token}"
    }

    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/attachments/{attachment_id}/$value"

    respuesta = requests.get(url, headers=headers)

    if respuesta.status_code != 200:
        print(f"No se pudo descargar: {nombre_archivo}")
        print(respuesta.status_code)
        print(respuesta.text)
        return False

    nombre_limpio = limpiar_nombre_archivo(nombre_archivo)
    ruta_archivo = os.path.join(carpeta_destino, nombre_limpio)

    contador = 1
    nombre_base, extension = os.path.splitext(nombre_limpio)

    while os.path.exists(ruta_archivo):
        nuevo_nombre = f"{nombre_base}_{contador}{extension}"
        ruta_archivo = os.path.join(carpeta_destino, nuevo_nombre)
        contador += 1

    with open(ruta_archivo, "wb") as archivo:
        archivo.write(respuesta.content)

    print(f"Descargado: {ruta_archivo}")
    return True


def main():
    crear_carpeta_si_no_existe(CARPETA_DESCARGAS)

    print("Obteniendo token de acceso...")
    token = obtener_token()

    print("Consultando correos con adjuntos...")
    correos = consultar_correos_con_adjuntos(token)

    print(f"\nCorreos encontrados con adjuntos: {len(correos)}\n")

    total_adjuntos_descargados = 0

    for index, correo in enumerate(correos, start=1):
        message_id = correo["id"]
        asunto = correo.get("subject", "Sin asunto")
        fecha = correo.get("receivedDateTime", "Sin fecha")

        print(f"\n[{index}/{len(correos)}] Correo:")
        print(f"Asunto: {asunto}")
        print(f"Fecha: {fecha}")

        carpeta_correo = limpiar_nombre_archivo(f"{fecha[:10]} - {asunto[:60]}")
        ruta_carpeta_correo = os.path.join(CARPETA_DESCARGAS, carpeta_correo)
        crear_carpeta_si_no_existe(ruta_carpeta_correo)     

        adjuntos = obtener_adjuntos(token, message_id)

        print(f"Adjuntos encontrados: {len(adjuntos)}")

        for adjunto in adjuntos:
            tipo_adjunto = adjunto.get("@odata.type", "")
            nombre = adjunto.get("name", "archivo_sin_nombre")
            attachment_id = adjunto.get("id")

            if tipo_adjunto != "#microsoft.graph.fileAttachment":
                print(f"Saltando adjunto no descargable directamente: {nombre}")
                continue

            descargado = descargar_adjunto(
                token=token,
                message_id=message_id,
                attachment_id=attachment_id,
                nombre_archivo=nombre,
                carpeta_destino=ruta_carpeta_correo
            )

            if descargado:
                total_adjuntos_descargados += 1

    print("\n===================================")
    print("PROCESO FINALIZADO")
    print(f"Correos procesados: {len(correos)}")
    print(f"Adjuntos descargados: {total_adjuntos_descargados}")
    print("===================================")


if __name__ == "__main__":
    main()