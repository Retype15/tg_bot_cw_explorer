import re
from telegram import Update, InputFile
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook

# Cargar o crear el archivo Excel
try:
    wb = load_workbook("datos_extraidos.xlsx")
    ws = wb.active
except FileNotFoundError:
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos Extraídos"
    ws.append(["Ubicación", "🇲🇴", "🇻🇦", "🟥", "🇪🇺"])  # Encabezados

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(f'¡Hola {update.effective_user.first_name}! Bienvenido al bot.')

async def send(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("Por favor, envía el mensaje que deseas guardar.")

def extract_location(message: str):
    pattern = r"([RGBY]{1,2}) (\d+#\d+)"
    match = re.search(pattern, message)

    if match:
        color_prefix = match.group(1).lower()  # Convertir el color a minúscula
        number_suffix = match.group(2).replace('#', '')  # Eliminar el #
        ubicacion = f"{color_prefix}{number_suffix}"
    else:
        ubicacion = "ubicacion_no_encontrada"

    return ubicacion

def extract_color_counts(message: str):
    color_counts = {'🇲🇴': 0, '🇻🇦': 0, '🟥': 0, '🇪🇺': 0}

    for color in color_counts.keys():
        count_pattern = rf"{color}\s*:\s*(\d+)"
        count_match = re.search(count_pattern, message)
        if count_match:
            color_counts[color] = int(count_match.group(1))

    return color_counts

def find_row_for_location(location):
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == location:
            return row
    return None

def save_to_excel(location, color_counts):
    row = find_row_for_location(location)

    if row is None:
        # Si la ubicación no existe, crea una nueva fila
        new_row = [location]
        for color_emoji in ['🇲🇴', '🇻🇦', '🟥', '🇪🇺']:
            new_row.append(color_counts.get(color_emoji, 0))
        ws.append(new_row)
    else:
        # Si la ubicación ya existe, actualiza la fila existente
        for i, color_emoji in enumerate(['🇲🇴', '🇻🇦', '🟥', '🇪🇺'], start=2):
            ws.cell(row=row, column=i).value = color_counts.get(color_emoji, 0)

    # Guardar los cambios en el archivo Excel
    wb.save("datos_extraidos.xlsx")

async def save_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.message.text

    ubicacion = extract_location(message)
    color_counts = extract_color_counts(message)

    save_to_excel(ubicacion, color_counts)

    await update.message.reply_text("Tu mensaje ha sido guardado y procesado.")
    await update.message.reply_text(f"Ubicación: {ubicacion}")
    await update.message.reply_text("Detalles de colores:")
    for color_emoji, count in color_counts.items():
        await update.message.reply_text(f"{count} {color_emoji}")

async def send_map(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # Verificar si el archivo existe y es accesible
    try:
        await update.message.reply_photo(photo=InputFile("map.jpg"))
    except Exception as e:
        await update.message.reply_text(f"Error al enviar la imagen: {e}")

async def info(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # Obtener la ubicación solicitada
    location = context.args[0] if context.args else None
    if not location:
        await update.message.reply_text("Por favor, proporciona una ubicación en el formato adecuado (por ejemplo, gy2).")
        return

    # Buscar la fila correspondiente a la ubicación
    row = find_row_for_location(location)
    color_emoji = ['🇲🇴', '🇻🇦', '🟥', '🇪🇺']
    if row:

        await update.message.reply_text(f"Ubicación: {location}")
        color_counts = {ws.cell(row, column=i).value for i in range(2, 8)}
        x = 0
        for count in color_counts:
            await update.message.reply_text(f"{color_emoji[x]} -> {count} ")
            x = x+1
    else:
        await update.message.reply_text(f"No se encontró información para la ubicación {location}.")

app = ApplicationBuilder().token("7523544789:AAE6u1waeC3kL3LpZK_7-J_CNqNTdPbybG4").build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("send", send))
app.add_handler(CommandHandler("map", send_map))
app.add_handler(CommandHandler("info", info))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, save_message))

app.run_polling()