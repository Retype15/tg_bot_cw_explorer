from datetime import datetime
import pytz
import re
from telegram import Update, InputFile
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook

# Ruta del archivo Excel
EXCEL_PATH = "datos_extraidos.xlsx"

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
    await update.message.reply_text(f'¡Hola {update.effective_user.first_name}! Bienvenido al bot.')

async def send(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("Por favor, envía el mensaje que deseas guardar.")

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
    """Extrae las cantidades de colores del mensaje."""
    color_counts = {'🇲🇴': 0, '🇻🇦': 0, '🇮🇲': 0, '🇪🇺': 0}

    for color in color_counts.keys():
        count_pattern = rf"{color}\s*:\s*(\d+)"
        count_match = re.search(count_pattern, message)
        if count_match:
            color_counts[color] = int(count_match.group(1))

    return color_counts

def find_row_for_location(ws, location):
    """Busca la fila correspondiente a una ubicación en el Excel."""
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == location:
            return row
    return None

async def get_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    file_path = EXCEL_PATH
        
    try:
        # Envía el documento al usuario
        with open(file_path, 'rb') as excel_file:
            await update.message.reply_document(document=excel_file)
    except FileNotFoundError:
        await update.message.reply_text("El archivo no se encontró.")
    except Exception as e:
        await update.message.reply_text(f"Ocurrió un error: {e}")

def save_to_excel(ws, location, color_counts, text):
    """Guarda los datos extraídos en el archivo Excel."""
    row = find_row_for_location(ws, location)

    if row is None:
        # Si la ubicación no existe, crea una nueva fila
        new_row = [location] + [color_counts.get(color_emoji, 0) for color_emoji in ['🇲🇴', '🇻🇦', '🇮🇲', '🇪🇺']] + [text]
        ws.append(new_row)
    else:
        # Si la ubicación ya existe, actualiza la fila existente
        for i, color_emoji in enumerate(['🇲🇴', '🇻🇦', '🇮🇲', '🇪🇺'], start=2):
            ws.cell(row=row, column=i).value = color_counts.get(color_emoji, 0)
        ws.cell(row=row, column=6).value = text  # Actualiza el texto en la columna 6
        ws.cell(row=row, column=7).value = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Guardar los cambios en el archivo Excel
    ws.parent.save(EXCEL_PATH)


async def save_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    wb, ws = cargar_o_crear_excel()
    message = update.message.text

    ubicacion = extract_location(message)
    color_counts = extract_color_counts(message)

    save_to_excel(ws, ubicacion, color_counts, message)

    await update.message.reply_text("Tu mensaje ha sido guardado y procesado.")
    await update.message.reply_text(f"Ubicación: {ubicacion}")
    await update.message.reply_text("Detalles de colores:")
    for color_emoji, count in color_counts.items():
        await update.message.reply_text(f"{count} {color_emoji}")

async def send_map(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        await update.message.reply_photo(photo=InputFile("map.jpg"))
    except Exception as e:
        await update.message.reply_text(f"Error al enviar la imagen: {e}")

async def info(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    wb, ws = cargar_o_crear_excel()
    
    location = context.args[0] if context.args else None
    if not location:
        await update.message.reply_text("Por favor, proporciona una ubicación en el formato adecuado (por ejemplo, gy2).")
        return

    row = find_row_for_location(ws, location)
    
    if row:
        # Aquí es donde se devuelve el texto de la quinta casilla
        
        saved_time = ws.cell(row=row, column=7).value
        time_difference = 0
        
        if saved_time:
            # Convertir la hora guardada a un objeto datetime
            saved_time_dt = datetime.strptime(saved_time, '%Y-%m-%d %H:%M:%S')
            
            # Convertir la hora guardada a la zona horaria local
            local_tz = pytz.timezone('America/Havana')
            saved_time_local = saved_time_dt.astimezone(local_tz)
            
            # Obtener la hora actual en la zona horaria local
            current_time = datetime.now(local_tz)
            
            # Calcular la diferencia en minutos
            time_difference = int((current_time - saved_time_local).total_seconds() / 60)
        else:
            time_difference = -1
        
        text = f"{ws.cell(row=row, column=6).value}\nTiempo transcurrido: {time_difference} minutos"
        if text:
            await update.message.reply_text(text)
        else:
            await update.message.reply_text(f"No hay texto guardado para la ubicación {location}.")
    else:
        await update.message.reply_text(f"No se encontró información para la ubicación {location}.")

async def simple_info(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
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
            # Convertir la hora guardada a un objeto datetime
            saved_time_dt = datetime.strptime(saved_time, '%Y-%m-%d %H:%M:%S')
            
            # Convertir la hora guardada a la zona horaria local
            local_tz = pytz.timezone('America/Havana')
            saved_time_local = saved_time_dt.astimezone(local_tz)
            
            # Obtener la hora actual en la zona horaria local
            current_time = datetime.now(local_tz)
            
            # Calcular la diferencia en minutos
            time_difference = int((current_time - saved_time_local).total_seconds() / 60)
        else:
            time_difference = -1
        
        msg = f"Ubicación: {location}"
        color_counts = [ws.cell(row, column=i).value for i in range(2, 6)]  # Extraer valores en una lista
        
        for x, count in enumerate(color_counts):
            msg += f"\n{color_emoji[x]} -> {count}"
        msg += f"\nTiempo transcurrido: {time_difference} minutos"
        await update.message.reply_text(msg)
    else:
        await update.message.reply_text(f"No se encontró información para la ubicación {location}.")


async def help(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Usa /i + ubicación (ej: gy2) para obtener la información completa de una casilla específica.\n"
        "Usa /info + ubicación (ej: gy2) para obtener el texto guardado en la quinta casilla de esa ubicación.\n"
        "Usa /get_excel para obtener el Excel con las ubicaciones de la base de datos."
    )

# Configuración del bot
app = ApplicationBuilder().token("6436295787:AAHQYGQj94g_1iuuzmU5RQa43esNok7Cj3g").build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("help", help))
app.add_handler(CommandHandler("send", send))
app.add_handler(CommandHandler("get_excel", get_excel))
app.add_handler(CommandHandler("map", send_map))
app.add_handler(CommandHandler("i", simple_info))  # Comando /i original
app.add_handler(CommandHandler("info", info))   # Nuevo comando /info
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, save_message))

# Inicia el bot
app.run_polling()
