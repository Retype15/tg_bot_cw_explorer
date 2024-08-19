from datetime import datetime
import pytz
import re
from telegram import Update, InputFile
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook

# Ruta del archivo Excel
EXCEL_PATH = "datos_extraidos.xlsx"
CHATWARS_ID = "265204902"

###########################SECURITY###############################################################

def load_authorized_users(file_path):
    try:
        with open(file_path, 'r') as file:
            users = [int(line.strip()) for line in file.readlines()]
        return users
    except FileNotFoundError:
        print(f"Error: El archivo {file_path} no fue encontrado.")
        return []
    except ValueError:
        print("Error: Asegúrate de que todos los IDs en el archivo sean números enteros.")
        return []

def save_authorized_users(file_path, users):
    try:
        with open(file_path, 'w') as file:
            for user_id in users:
                file.write(f"{user_id}\n")
    except Exception as e:
        print(f"Error al guardar en {file_path}: {e}")

AUTHORIZED_USERS = load_authorized_users("users.txt")

def is_authorized(user_id):
    """Verifica si el usuario está autorizado."""
    return user_id in AUTHORIZED_USERS
    
async def validate(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if is_authorized(update.effective_user.id):
        if context.args:
            try:
                new_user_id = int(context.args[0])
                if new_user_id not in AUTHORIZED_USERS:
                    AUTHORIZED_USERS.append(new_user_id)
                    save_authorized_users("users.txt", AUTHORIZED_USERS)
                    await update.message.reply_text(f"El usuario con ID {new_user_id} ha sido validado exitosamente.")
                else:
                    await update.message.reply_text(f"El usuario con ID {new_user_id} ya está validado.")
            except ValueError:
                await update.message.reply_text("Por favor, proporciona un ID de usuario válido.")
        else:
            await update.message.reply_text("Por favor, proporciona el ID del usuario que deseas validar.")
    else:
        await update.message.reply_text("Lo siento, no tienes permiso para usar este bot.")

##################################################################################################

def cargar_o_crear_excel():
    """Carga el archivo Excel si existe, de lo contrario, lo crea."""
    try:
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Datos Extraídos"
        ws.append(["Ubicación", "🇲🇴", "🇻🇦", "🇮🇲", "🇪🇺", "Texto"])  # Encabezados
        wb.save(EXCEL_PATH)
    return wb, ws

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    #####SEGURIDAD#####
    if is_authorized(update.effective_user.id):
        await update.message.reply_text(f'隆Hola {update.effective_user.first_name}! Bienvenido al bot.')
    else:
        await update.message.reply_text("Lo siento, no tienes permiso para usar este bot.")
    ###################

def extract_location(message: str):
    """Extrae la ubicación del mensaje y la convierte al formato esperado."""
    pattern = r"([RGBY]{1,2})\s*(\d+)(?:#(\d+))?"
    match = re.search(pattern, message)

    if match:
        color_prefix = match.group(1).lower()  # Convertir el color a minúscula
        number1 = match.group(2)  # Primer número
        number2 = match.group(3) if match.group(3) else ""  # Segundo número opcional
        ubicacion = f"{color_prefix}{number1}{number2}"
    else:
        ubicacion = "ubicacion_no_encontrada"

    return ubicacion

def extract_color_counts(message: str):
    """Extrae las cantidades de colores del mensaje y cuenta las banderas si no se encuentra el patrón específico."""
    
    # Inicialización de los conteos de banderas y el patrón de búsqueda
    color_counts = {'🇲🇴': 0, '🇻🇦': 0, '🇮🇲': 0, '🇪🇺': 0}
    color_patterns = {'🇲🇴': 0, '🇻🇦': 0, '🇮🇲': 0, '🇪🇺': 0}

    # Buscar patrones específicos de colores
    for color in color_counts.keys():
        count_pattern = rf"{color}\s*:\s*(\d+)"
        count_match = re.search(count_pattern, message)
        if count_match:
            color_counts[color] = int(count_match.group(1))

    # Si no se encontraron patrones específicos, contar todas las banderas
    if all(value == 0 for value in color_counts.values()):
        for color in color_counts.keys():
            color_patterns[color] = len(re.findall(color, message))
        color_counts = color_patterns

    return color_counts

def find_row_for_location(ws, location):
    """Busca la fila correspondiente a una ubicación en el Excel."""
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == location:
            return row
    return None

async def get_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    #####SEGURIDAD#####
    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Lo siento, no tienes permiso para usar este bot.")
        return[]
    ##################
    
    file_path = EXCEL_PATH
        
    try:
        # Envía el documento al usuario
        with open(file_path, 'rb') as excel_file:
            await update.message.reply_document(document=excel_file)
    except FileNotFoundError:
        await update.message.reply_text("El archivo no se encontró.")
    except Exception as e:
        await update.message.reply_text(f"Ocurrió un error: {e}")

def save_to_excel(ws, location, color_counts, text, user_posted):
    """Guarda los datos extraídos en el archivo Excel."""
    row = find_row_for_location(ws, location)

    if row is None:
        # Si la ubicación no existe, crea una nueva fila
        new_row = [location] + [color_counts.get(color_emoji, 0) for color_emoji in ['🇲🇴', '🇻🇦', '🇮🇲', '🇪🇺']] + [text, user_posted]
        ws.append(new_row)
    else:
        # Si la ubicación ya existe, actualiza la fila existente
        for i, color_emoji in enumerate(['🇲🇴', '🇻🇦', '🇮🇲', '🇪🇺'], start=2):
            ws.cell(row=row, column=i).value = color_counts.get(color_emoji, 0)
        ws.cell(row=row, column=6).value = text  # Actualiza el texto en la columna 6
        ws.cell(row=row, column=7).value = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ws.cell(row=row, column=9).value = user_posted  # Actualiza el nombre del usuario en la columna 9

    # Guardar los cambios en el archivo Excel
    ws.parent.save(EXCEL_PATH)


patterncomp = re.compile(
    r"^You (climbed to the highest point in the|looked to the)\s+"  # Primera línea
    r"(?:\[\w\s\d+#?\d*\]|[Y\s\d+#?\d*])?\s*"  # Ubicación (opcional)
    r"(?:Total:\s*\d+\s*👥\s*(?:🇲🇴|🇻🇦|🇮🇲|🇪🇺\s*:\s*\d+\s*👥,\s*Leader:\s*.+\s*)?)?"  # Total e información de equipo (opcional)
    r"((?:🇲🇴|🇻🇦|🇮🇲|🇪🇺[\w\d\s]+ 🏅\d+ 👣\d+\s*)*)",  # Lista de usuarios (opcional)
    #r"(?:Combat options: /combat)?$",  # Opción de combate (opcional)
    re.DOTALL | re.MULTILINE
)

def es_mensaje_valido(mensaje: str) -> bool:
    # Si hay información de equipo, debe haber total
    #if ("🇲🇴" in mensaje or "🇻🇦" in mensaje or "🇮🇲" in mensaje or "🇪🇺" in mensaje) and "Total:" not in mensaje:
    #    return False
    return bool(patterncomp.match(mensaje))

async def save_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    
    #####SEGURIDAD#####
    if update.message.chat.type in [update.message.chat.SUPERGROUP, update.message.chat.GROUP]:
        return []
    
    if hasattr(update.message, 'forward_from') and update.message.forward_from and update.message.forward_from.id == 265204902:
        await update.message.reply_text("¡Este mensaje fue reenviado desde Chat Wars (@ChatWarsBot)!")
    elif hasattr(update.message, 'forward_origin') and update.message.forward_origin and update.message.forward_origin.sender_user and update.message.forward_origin.sender_user.id == 265204902:
        await update.message.reply_text("Procesando información...")
    else:
        await update.message.reply_text("Este comando solo puede ser usado en mensajes reenviados desde Chat Wars (@ChatWarsBot).")
        return []
    ###################
    
    message = update.message.text
    
    if not es_mensaje_valido(message):
        await update.message.reply_text("Mensaje enviado no valido..!")
        return []
    
    wb, ws = cargar_o_crear_excel()
    
    ubicacion = extract_location(message)
    color_counts = extract_color_counts(message)
    
    user_posted = update.effective_user.username or update.effective_user.full_name

    save_to_excel(ws, ubicacion, color_counts, message, user_posted)

    msg = ""
    for color_emoji, count in color_counts.items():
        msg += f"\n{color_emoji} -> {count}"
    await update.message.reply_text(f"Guardado!\nUbicación: {ubicacion}\nDetalles de colores: {msg}\nPosted By: {user_posted}")


async def send_map(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    #####SEGURIDAD#####
    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Lo siento, no tienes permiso para usar este bot.")
        return[]
    ##################
    
    try:
        await update.message.reply_photo(photo=InputFile("map.jpg"))
    except Exception as e:
        await update.message.reply_text(f"Error al enviar la imagen: {e}")

async def info(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    #####SEGURIDAD#####
    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Lo siento, no tienes permiso para usar este bot.")
        return[]
    ##################
    
    wb, ws = cargar_o_crear_excel()
    
    location = context.args[0] if context.args else None
    if not location:
        await update.message.reply_text("Por favor, proporciona una ubicación en el formato adecuado (por ejemplo, gy2).")
        return

    row = find_row_for_location(ws, location)
    
    if row:
        saved_time = ws.cell(row=row, column=7).value
        time_difference = 0
        
        if saved_time:
            saved_time_dt = datetime.strptime(saved_time, '%Y-%m-%d %H:%M:%S')
            local_tz = pytz.timezone('America/Havana')
            saved_time_local = saved_time_dt.astimezone(local_tz)
            current_time = datetime.now(local_tz)
            time_difference = int((current_time - saved_time_local).total_seconds() / 60)
        else:
            time_difference = -1
        
        text = f"{ws.cell(row=row, column=6).value}\nTiempo transcurrido: {time_difference} minutos\nPosted by: {ws.cell(row=row, column=9).value}"
        await update.message.reply_text(text)
    else:
        await update.message.reply_text(f"No se encontró información para la ubicación {location}.")

async def simple_info(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    #####SEGURIDAD#####
    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Lo siento, no tienes permiso para usar este bot.")
        return[]
    ##################
    wb, ws = cargar_o_crear_excel()
    
    location = context.args[0] if context.args else None
    if not location:
        await update.message.reply_text("Por favor, proporciona una ubicación en el formato adecuado (por ejemplo, gy2).")
        return

    row = find_row_for_location(ws, location)
    color_emoji = ['🇲🇴', '🇻🇦', '🇮🇲', '🇪🇺']
    
    if row:
        saved_time = ws.cell(row=row, column=7).value
        time_difference = 0
        
        if saved_time:
            saved_time_dt = datetime.strptime(saved_time, '%Y-%m-%d %H:%M:%S')
            local_tz = pytz.timezone('America/Havana')
            saved_time_local = saved_time_dt.astimezone(local_tz)
            current_time = datetime.now(local_tz)
            time_difference = int((current_time - saved_time_local).total_seconds() / 60)
        else:
            time_difference = -1
        
        msg = f"Ubicación: {location}"
        color_counts = [ws.cell(row, column=i).value for i in range(2, 6)]
        
        for x, count in enumerate(color_counts):
            msg += f"\n{color_emoji[x]} -> {count}"
        msg += f"\nTiempo transcurrido: {time_difference} minutos\nPosted by: {ws.cell(row=row, column=9).value}"
        await update.message.reply_text(msg)
    else:
        await update.message.reply_text(f"No se encontró información para la ubicación {location}.")


async def help(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    #####SEGURIDAD#####
    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Lo siento, no tienes permiso para usar este bot.")
        return[]
    ##################
    await update.message.reply_text(
        "Usa /i + ubicación (ej: gy2) para obtener la información completa de una casilla específica.\n"
        "Usa /info + ubicación (ej: gy2) para obtener el texto guardado en la quinta casilla de esa ubicación.\n"
        "Usa /get_excel para obtener el Excel con las ubicaciones de la base de datos."
    )

# Configuración del bot
app = ApplicationBuilder().token("7523544789:AAE6u1waeC3kL3LpZK_7-J_CNqNTdPbybG4").build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("help", help))
app.add_handler(CommandHandler("get_excel", get_excel))
app.add_handler(CommandHandler("map", send_map))
app.add_handler(CommandHandler("i", simple_info))  # Comando /i original
app.add_handler(CommandHandler("info", info))   # Nuevo comando /info
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, save_message))

# Inicia el bot
app.run_polling()
