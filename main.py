# main.py - VERSIÓN CON CORRECCIÓN FINAL
import os
import io
import json
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from PIL import Image
import google.generativeai as genai
import report_generator


# --- CONFIGURACIÓN ---
app = Flask(__name__)
CORS(app)
informe_data = {}
# Configuración de APIs
try:
    GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY')
    genai.configure(api_key=GEMINI_API_KEY)
    print("✅ API de Gemini configurada.")
except Exception as e:
    print(f"❌ Error configurando la API de Gemini: {e}")

SCOPES = ['https://www.googleapis.com/auth/drive']

# --- FUNCIONES ---
def authenticate_google_drive():
    try:
        creds_json = os.environ.get('GOOGLE_CREDENTIALS')
        if not creds_json:
            print("❌ GOOGLE_CREDENTIALS no encontrado")
            return None, None
        creds_info = json.loads(creds_json)
        credentials = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        service = build('drive', 'v3', credentials=credentials)
        print("✅ Autenticación de Drive exitosa.")
        return service, creds_info.get('client_email')
    except Exception as e:
        print(f"❌ Error autenticando Drive: {e}")
        return None, None

def find_drive_id(service, q: str, include_all_drives: bool = False, drive_id: str = None):
    """
    Devuelve el primer fileId que cumpla la query q.
    - include_all_drives=True: busca en Mi Unidad + Unidades compartidas (allDrives).
    - drive_id: si apuntas a una Unidad compartida específica, pásala aquí.
    """
    # Asegura que no traes papelera
    q_final = f"({q}) and trashed = false" if "trashed" not in q.lower() else q

    params = {
        "q": q_final,
        "fields": "files(id,name)",
        "spaces": "drive",
        "pageSize": 1,
    }

    # Solo agregamos estos parámetros si realmente se usan
    if include_all_drives or drive_id:
        params["supportsAllDrives"] = True
        params["includeItemsFromAllDrives"] = True
        # Si pasas drive_id usa 'drive', si no, 'allDrives'
        if drive_id:
            params["corpora"] = "drive"
            params["driveId"] = drive_id
        else:
            params["corpora"] = "allDrives"

    resp = service.files().list(**params).execute()
    files = resp.get("files", [])
    return files[0]["id"] if files else None


def download_image_bytes(service, file_id):
    try:
        request_download = service.files().get_media(fileId=file_id)
        file_bytes = io.BytesIO()
        downloader = MediaIoBaseDownload(file_bytes, request_download)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        print(f"✅ Imagen {file_id} descargada.")
        return file_bytes.getvalue()
    except Exception as e:
        # Aquí es donde ocurría el error si 'service' era una tupla
        print(f"❌ Error en download_image_bytes para {file_id}: {e}")
        return None

def generate_ai_description(prompt, image_list):
    try:
        model = genai.GenerativeModel('models/gemini-1.5-pro-latest')
        response = model.generate_content([prompt] + image_list)
        print("✅ Descripción de IA generada.")
        return response.text
    except Exception as e:
        print(f"❌ Error en la API de IA: {e}")
        return f"Error al generar descripción: {e}"

def listar_imagenes_de_carpeta(service, carpeta_id):
    try:
        query = f"'{carpeta_id}' in parents and (mimeType contains 'image/') and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        print(f"📸 Encontradas {len(results.get('files', []))} imágenes en la carpeta.")
        return [{'id': img['id'], 'name': img['name']} for img in results.get('files', [])]
    except Exception as e:
        print(f"❌ Error listando imágenes: {e}")
        return []

def buscar_carpeta_por_nombre(service, nombre_carpeta):
    try:
        query = f"name='{nombre_carpeta}' and mimeType='application/vnd.google-apps.folder' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        folders = results.get('files', [])
        if folders:
            print(f"✅ Carpeta '{nombre_carpeta}' encontrada.")
            return folders[0]['id']
        else:
            print(f"❌ No se encontró la carpeta '{nombre_carpeta}'.")
            return None
    except Exception as e:
        print(f"❌ Error buscando carpeta: {e}")
        return None

# --- ENDPOINTS DE LA API ---
     # En main.py

# En main.py

@app.route('/api/list-images', methods=['POST'])
def list_images():
    try:
        data = request.get_json(force=True) or {}
        print(f"[/api/list-images] payload: {data}")
        info_proyecto = (data.get("info_proyecto") or {})
        folder_name = (info_proyecto.get("folder_name") or data.get("folder_name") or "").strip()
        folder_id = data.get("folder_id")

        # 1) Autenticación SIEMPRE con Service Account
        service, sa_email = authenticate_google_drive()
        if not service:
            return jsonify({"error": "No se pudo autenticar con Drive."}), 500

        # 2) Resolver folder_id si vino sólo el nombre
        if not folder_id:
            if not folder_name:
                return jsonify({"error": "Falta 'folder_name' o 'folder_id'."}), 400

            folder_id = find_drive_id(
                service,
                "name = '{0}' and mimeType = 'application/vnd.google-apps.folder'".format(folder_name),
                include_all_drives=True
            )
            if not folder_id:
                return jsonify({"error": f"No se encontró la carpeta '{folder_name}' (o la SA no tiene permisos)."}), 404

        # 3) Listar imágenes dentro de la carpeta (incluye Shared Drives)
        imgs_q = f"'{folder_id}' in parents and mimeType contains 'image/' and trashed = false"
        resp_imgs = service.files().list(
            q=imgs_q,
            fields="files(id,name,mimeType,webViewLink,thumbnailLink)",
            pageSize=200,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        images = [
            {
                "id": f["id"],
                "name": f["name"],
                "mimeType": f.get("mimeType"),
                "webViewLink": f.get("webViewLink")
            }
            for f in resp_imgs.get("files", [])
        ]

        # 4) Resolver archivos estáticos esperados (dentro de la misma carpeta)
        parent_q = f"'{folder_id}' in parents and trashed = false"

        # Tablas.xlsx
        tablas_id = find_drive_id(service, f"{parent_q} and name = 'Tablas.xlsx'", include_all_drives=True)

        # Logo: intentamos varias extensiones por si cambia
        logo_id = (
            find_drive_id(service, f"{parent_q} and name = 'logo2.jpg'", include_all_drives=True)
            or find_drive_id(service, f"{parent_q} and name = 'logo2.png'", include_all_drives=True)
            or find_drive_id(service, f"{parent_q} and name = 'logo.jpg'", include_all_drives=True)
            or find_drive_id(service, f"{parent_q} and name = 'logo.png'", include_all_drives=True)
        )

        # Ubicación(es)
        img_ubicacion_proyecto_id = (
            find_drive_id(service, f"{parent_q} and name = 'ubicacion.png'", include_all_drives=True)
            or find_drive_id(service, f"{parent_q} and name = 'ubicacion.jpg'", include_all_drives=True)
        )
        img_ubicacion_paradas_id = (
            find_drive_id(service, f"{parent_q} and name = 'ubicacion_paraderos.png'", include_all_drives=True)
            or find_drive_id(service, f"{parent_q} and name = 'ubicacion_paraderos.jpg'", include_all_drives=True)
        )
        print(f"[/api/list-images] OK folder_id={folder_id} tablas_id={tablas_id} imgs={len(images)}")


        # 5) Respuesta (incluye alias 'tablas' para compatibilidad con el front antiguo)
        return jsonify({
            "ok": True,
            "folder_id": folder_id,
            "service_account": sa_email,
            "images": images,
            "drive_file_ids": {
                "tablas_id": tablas_id,
                "logo_id": logo_id,
                "img_ubicacion_proyecto_id": img_ubicacion_proyecto_id,
                "img_ubicacion_paradas_id": img_ubicacion_paradas_id
            },
            "tablas": tablas_id  # <-- alias legacy
        }), 200

    except Exception as e:
        print(f"❌ /api/list-images error: {e}")
        return jsonify({"error": str(e)}), 500
    
@app.route('/api/analyze-image', methods=['POST'], strict_slashes=False)
def handle_analyze_image():
    """Recibe IDs de imagen y datos del paradero, analiza con IA y devuelve una descripción."""
    print("\n--- Petición en /api/analyze-image ---")
    data = request.get_json()
    image_ids = data.get('image_ids', [])
    prompt_type = data.get('prompt_type')
    # ¡NUEVO! Recibimos el código del paradero desde el formulario
    codigo_paradero = data.get('codigo_paradero', 'No especificado')

    if not image_ids or not prompt_type:
        return jsonify({'error': 'Faltan image_ids o prompt_type'}), 400

    # ¡CAMBIO CLAVE! El prompt "general" ahora es una plantilla.
    PROMPTS = {
        'general': (
            "Eres un asistente experto en ingeniería de transporte y vialidad, especializado en la evaluación de paraderos de autobuses. Tu tarea es analizar la imagen proporcionada para el paradero con código {codigo_paradero} y generar una descripción técnica y concisa. En tu descripción, debes identificar claramente la presencia y el estado de los siguientes elementos: refugio, andén, banca, señal informativa, demarcación en el pavimento, y si existe o no huella podo táctil. Finalmente, basándote en todos los elementos observados, determina si el paradero parece cumplir o no con el estándar de diseño del DTPM (Directorio de Transporte Público Metropolitano) y justifica brevemente por qué. Formato: Párrafo único y directo. No uses listas ni puntos."
        ),
        
        'refugio_anden': ("Eres un inspector de infraestructura de transporte. Analiza la(s) imagen(es) de un refugio y andén de paradero. "
        "En tu descripción, evalúa los siguientes puntos clave: "
        "1. Refugio: Estado general de la estructura, materiales y su limpieza (busca rayados o basura). "
        "2. Techumbre: Condición y protección que ofrece contra sol y lluvia. "
        "3. Andén: Estado del pavimento y, muy importante, la presencia o ausencia de baldosas y huellas podo táctiles. "
        "4. Iluminación: Indica si se observa o no iluminación artificial. "
        "Genera un párrafo único y conciso que resuma tus hallazgos."),
        
        'senal': ("Eres un asistente técnico que describe evidencia visual para un informe. Tu única tarea es describir el estado de la "
            "señalización y demarcación de un paradero de bus, basándote exclusivamente en la imagen proporcionada. "
            "1. Sobre la señal (el letrero y su poste): Describe su estado físico. ¿Se ve nuevo, desgastado, dañado o rayado? "
            "2. Sobre la normativa de la señal: Visualmente, ¿el diseño del letrero (colores, tipografía) parece cumplir con los estándares gráficos del DTPM? "
            "3. Sobre la demarcación en el pavimento: Describe lo que ves en el suelo. ¿Hay un 'cajón de detención' pintado para el bus? ¿Está visible o desgastado? "
            "Reglas importantes: No incluyas un título en tu respuesta. No sugieras inspecciones adicionales. Sintetiza todo en un solo párrafo.")
    }

    # Insertamos el código del paradero en el prompt si es de tipo 'general'
    if prompt_type == 'general':
        selected_prompt = PROMPTS['general'].format(codigo_paradero=codigo_paradero)
    else:
        selected_prompt = PROMPTS.get(prompt_type, "Describe la imagen.")

    print(f"Usando prompt para '{prompt_type}': {selected_prompt[:100]}...") # Imprime los primeros 100 caracteres del prompt

    service, _ = authenticate_google_drive()
    if not service:
        return jsonify({'error': 'Fallo en la autenticación con Google Drive'}), 500

    images_for_model = []
    for img_id in image_ids:
        image_bytes = download_image_bytes(service, img_id)
        if image_bytes:
            img = Image.open(io.BytesIO(image_bytes))
            images_for_model.append(img)

    if not images_for_model:
        return jsonify({'error': 'No se pudieron descargar las imágenes seleccionadas'}), 500

    description = generate_ai_description(selected_prompt, images_for_model)

    return jsonify({'description': description})


@app.route('/api/save-description', methods=['POST'], strict_slashes=False)
def save_description():
    try:
        data = request.get_json()
        prompt_type = data.get('prompt_type') or data.get('type')
        description = data.get('description')
        image_ids = data.get('image_ids')

        if not prompt_type or description is None:
            return jsonify({'error': 'Faltan datos (prompt_type o description)'}), 400

        # Creamos una sección de 'analisis' si no existe
        if 'analisis' not in informe_data:
            informe_data['analisis'] = {}

        # Guardamos los datos vinculando la descripción a su tipo e imágenes
        informe_data['analisis'][prompt_type] = {
            'description': description,
            'image_ids': image_ids
        }

        print("✅ Descripción guardada. Estado actual de los datos del informe:")
        # Usamos json.dumps para imprimir el diccionario de forma legible
        print(json.dumps(informe_data, indent=2, ensure_ascii=False))

        return jsonify({'status': 'ok', 'message': f'Descripción para "{prompt_type}" guardada correctamente.'})

    except Exception as e:
        print(f"❌ Error en /api/save-description: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/fill-table', methods=['POST'])
def fill_table_data():
    print("\n--- Petición recibida en /api/fill-table ---")
    if 'analisis' not in informe_data or len(informe_data['analisis']) < 3:
        return jsonify({'error': 'Primero debe generar y guardar las 3 descripciones.'}), 400

    contexto = (
        f"Descripción General: {informe_data['analisis'].get('general', {}).get('description', 'No disponible.')}\n\n"
        f"Descripción de Refugio y Andén: {informe_data['analisis'].get('refugio_anden', {}).get('description', 'No disponible.')}\n\n"
        f"Descripción de Señal y Demarcación: {informe_data['analisis'].get('senal', {}).get('description', 'No disponible.')}"
    )

    # SINCRONIZAMOS LAS CARACTERÍSTICAS Y OPCIONES CON EL FRONTEND
    caracteristicas_map = {
        "Posee refugio": ["Sí", "No"],
        "Estándar del refugio": ["DTPM", "No es DTPM", "N.A."],
        "Estado de conservación del refugio": ["Sin refugio presente", "Deficiente", "Regular", "Bueno"],
        "Posee basurero": ["Sí", "No"],
        "Posee señal de parada": ["Sí", "No"],
        "Señal cumple norma gráfica": ["Sí", "No", "N.A."],
        "Estado de conservación de la señal": ["Sin señal presente", "Deficiente", "Regular", "Bueno"],
        "Iluminación": ["Sin iluminación presente", "Deficiente", "Buena"],
        "Posee andén": ["Sí", "No"],
        "Estado de conservación del andén": ["Sin andén presente", "Deficiente", "Regular", "Bueno"],
        "Posee conexión a la vereda": ["Sí", "No"],
        "Posee huella podo táctil al borde del andén": ["Sí", "No"],
        "Demarcación del cajón de parada": ["Sí posee", "No posee"]
    }

    opciones_texto = ""
    for car, opts in caracteristicas_map.items():
        opciones_texto += f"- Para '{car}', elige una de estas opciones: {opts}\n"

    # EL NUEVO SÚPER PROMPT CON INSTRUCCIONES PARA COMENTARIOS
    prompt_final = (
        "Eres un analista técnico que extrae datos estructurados de informes de inspección. A continuación te entrego el contexto completo "
        "de un paradero de autobús:\n\n--- CONTEXTO ---\n{contexto}\n\n--- FIN DEL CONTEXTO ---\n\n"
        "Tu tarea es leer el contexto y rellenar un objeto JSON. Para cada característica de la siguiente lista, elige la opción que mejor la describa.\n"
        "Lista de características y sus opciones permitidas:\n{opciones_texto}\n"
        "REGLA ESPECIAL: Para la característica 'Estado de conservación del refugio', el valor en el JSON debe ser un objeto con dos claves: "
        "'seleccion' (con la opción elegida) y 'comentario' (con una observación MUY BREVE de máximo 5 palabras, como 'Falta limpieza' o 'Estructura en buen estado').\n"
        "Responde únicamente con un objeto JSON válido, sin explicaciones ni texto adicional."
    ).format(contexto=contexto, opciones_texto=opciones_texto)

    print("Enviando súper prompt final a la IA...")

    try:
        model = genai.GenerativeModel(model_name='gemini-1.5-pro-latest')
        response = model.generate_content(prompt_final)

        json_response_text = response.text.strip().replace('```json', '').replace('```', '')
        table_data = json.loads(json_response_text)

        print("✅ Datos para la tabla generados y parseados exitosamente.")
        return jsonify(table_data)

    except Exception as e:
        print(f"❌ Error generando los datos de la tabla: {e}")
        return jsonify({'error': f'Error al procesar la respuesta de la IA: {e}'}), 500


@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    """
    Llama al generador de informes y devuelve el archivo .docx para su descarga.
    """
    try:
        print("Solicitud para generar informe recibida.")

        datos_completos = request.get_json(force=True) or {}
        if not datos_completos:
            return jsonify({'error': 'No se recibieron datos para generar el informe.'}), 400
        # Supongamos que guardaste las descripciones en informe_data['analisis'] por tipo
        analisis_guardado = (informe_data.get('analisis') if 'informe_data' in globals() else {}) or {}

        # Si quieres que aplique a TODOS los paraderos:
        for p in (datos_completos.get("paraderos") or []):
            base = p.get("analisis") or {}
            # lo guardado pisa lo generado por IA
            p["analisis"] = {**base, **analisis_guardado}

        info_proyecto = (datos_completos.get("info_proyecto") or {})
        folder_name = (info_proyecto.get("folder_name") or "").strip()
        drive_file_ids = (datos_completos.get("drive_file_ids") or {})

        service_drive, _ = authenticate_google_drive()

        # Log mínimo para depurar:
        print(f"[generate-report] folder_name='{folder_name}' | drive_file_ids_keys={list(drive_file_ids.keys())}")


        # 4. Ahora la validación funcionará
        if not folder_name and not drive_file_ids:
            # Si no hay carpeta ni IDs directos, no se puede seguir
            return jsonify({'error': 'Debe indicar "info_proyecto.folder_name" o proveer "drive_file_ids".'}), 400

        # 0) Autenticación y preparación de inputs
        service_drive, _ = authenticate_google_drive()

        info_proyecto = (datos_completos.get("info_proyecto") or {})
        folder_name = (info_proyecto.get("folder_name") or "").strip()
        drive_file_ids_payload = (datos_completos.get("drive_file_ids") or {})

        print(f"[generate-report] folder_name='{folder_name}' | drive_file_ids_keys={list(drive_file_ids_payload.keys())}")

        # 1) Validación flexible: carpeta O ids directos
        if not folder_name and not drive_file_ids_payload:
            return jsonify({'error': 'Debe indicar "info_proyecto.folder_name" o proveer "drive_file_ids".'}), 400

        # 2) Resolver IDs por el ramal correspondiente
        logo_id = tablas_id = img_ubicacion_proyecto_id = img_ubicacion_paradas_id = None

        if folder_name:
            folder_id = find_drive_id(
                service_drive,
                "name = '{0}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false".format(folder_name),
                include_all_drives=True,
            )
            if not folder_id:
                return jsonify({'error': f"No se encontró la carpeta '{folder_name}' en Drive (o no tienes permisos)."}), 404

            parent_q = f"'{folder_id}' in parents"
            tablas_id = find_drive_id(service_drive, f"{parent_q} and name = 'Tablas.xlsx'", include_all_drives=True)
            logo_id = find_drive_id(service_drive, f"{parent_q} and name = 'logo2.jpg'", include_all_drives=True)
            img_ubicacion_proyecto_id = find_drive_id(service_drive, f"{parent_q} and name = 'ubicacion.png'", include_all_drives=True)
            img_ubicacion_paradas_id = find_drive_id(service_drive, f"{parent_q} and name = 'ubicacion_paraderos.png'", include_all_drives=True)

        else:
            # Ramal por IDs directos desde el front (no toques Drive)
            tablas_id = (drive_file_ids_payload or {}).get("tablas_id")
            logo_id = (drive_file_ids_payload or {}).get("logo_id")
            img_ubicacion_proyecto_id = (drive_file_ids_payload or {}).get("img_ubicacion_proyecto_id")
            img_ubicacion_paradas_id = (drive_file_ids_payload or {}).get("img_ubicacion_paradas_id")

        # 4) Llamada al generador: pásale SIEMPRE el paquete de IDs resueltos
        document = report_generator.crear_informe_paraderos(
            datos_informe=datos_completos,
            service_drive=service_drive,
            drive_file_ids={
                "logo_id": logo_id,
                "tablas_id": tablas_id,
                "img_ubicacion_proyecto_id": img_ubicacion_proyecto_id,
                "img_ubicacion_paradas_id": img_ubicacion_paradas_id,
            }
        )

        if document:
            file_stream = io.BytesIO()
            document.save(file_stream)
            file_stream.seek(0)

            info_proyecto = datos_completos.get("info_proyecto", {})
            nombre_archivo = f"Informe_{info_proyecto.get('proyecto', 'Proyecto')}.docx"

            print(f"✅ Enviando el archivo '{nombre_archivo}' para descarga.")

            return send_file(
                file_stream,
                as_attachment=True,
                download_name=nombre_archivo,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            return jsonify({'error': 'No se pudo generar el documento.'}), 500


    except Exception as e:
        print(f"❌ Error en /api/generate-report: {e}")
        return jsonify({'error': str(e)}), 500

@app.route("/api/gem-health")
def gem_health():
    try:
        import google.generativeai as genai
        genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
        return {"ok": True}, 200
    except Exception as e:
        return {"ok": False, "error": str(e)}, 500

# --- INICIO DEL SERVIDOR ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=81, debug=True)
