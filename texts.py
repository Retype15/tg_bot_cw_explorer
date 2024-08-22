from telegram import Update, InputFile
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook

TEXTS = {
    'en': {
        'welcome': "Welcome {name} to our exploration center!",
        'no_permission': "Sorry, you don't have permission to use this bot.",
        'provide_location': "Please provide a location in the correct format (e.g., gy2).",
        'simple_info_header': "Location: {location}",
        'color_count': "\n{0} -> {1}",
        'simple_info_footer': "\nElapsed time: {time_difference} minutes\nPosted by: {user}",
        'no_info_found': "No information found for location {location}.",
        'saved_successfully': "Saved!\nLocation: {location}\nColor Details: {msg}\nPosted By: {user_posted}",
        'message_forwarded': "The message must be forwarded from Chat Wars (@ChatWarsBot)!",
        'message_in_group': "Report received, thank you for your help! 🫡",
        'processing_info': "Processing information...",
        'invalid_message': "Invalid message format!",
        'get_excel_error': "File not found.",
        'get_excel_error_exception': "An error occurred: {error}",
        'map_error': "Error sending the image: {error}",
        'choose_language': "Please choose your preferred language:",
        'default': 'ERROR:XXX>Sorry, the requested text is not available.',
        'help_message': (
            "Help:\n",
            "Use /i + location (e.g., gy2) to get complete information about a specific tile.\n"
            "Use /info + location (e.g., y31) to get the saved text in the fifth cell of that location.\n"
            "Use /get_excel to get the Excel file with the database locations.\n"
            "Use /set_language to change the bot's language."
        ),
    },
    'es': {
        'welcome': "Bienvenido/a {name} a nuestro centro de exploracion!",
        'no_permission': "Lo siento, no tienes permiso para usar este bot.",
        'provide_location': "Por favor, proporciona una ubicación en el formato adecuado (por ejemplo, gy2).",
        'simple_info_header': "Ubicación: {location}",
        'color_count': "\n{0} -> {1}",
        'simple_info_footer': "\nTiempo transcurrido: {time_difference} minutos\nPublicado por: {user}",
        'no_info_found': "No se encontró información para la ubicación {location}.",
        'saved_successfully': "¡Guardado!\nUbicación: {location}\nDetalles de colores: {msg}\nPublicado por: {user_posted}",
        'message_forwarded': "¡El mensaje debe ser reenviado desde Chat Wars (@ChatWarsBot)!",
        'message_in_group': "Reporte recibido, ¡gracias por tu ayuda! 🫡",
        'processing_info': "Procesando información...",
        'invalid_message': "¡Mensaje enviado no válido!",
        'get_excel_error': "El archivo no se encontró.",
        'get_excel_error_exception': "Ocurrió un error: {error}",
        'map_error': "Error al enviar la imagen: {error}",
        'choose_language': "Por favor, elige tu idioma preferido:",
        'default': 'ERROR:XXX>Lo siento, el texto solicitado no está disponible.',
        'help_message': (
            "Ayuda:\n",
            "Usa /i + ubicación (ej: gy2) para obtener la información completa de una casilla específica.\n"
            "Usa /info + ubicación (ej: y41) para obtener el texto guardado en la quinta casilla de esa ubicación.\n"
            "Usa /get_excel para obtener el Excel con las ubicaciones de la base de datos.\n"
            "Usa /set_language para cambiar el idioma del bot."
        ),
    },
    'ru': {
        'welcome': "Добро пожаловать, {name}, в наш исследовательский центр!",
        'no_permission': "Извините, у вас нет разрешения на использование этого бота.",
        'provide_location': "Пожалуйста, укажите местоположение в правильном формате (например, gy2).",
        'simple_info_header': "Местоположение: {location}",
        'color_count': "\n{0} -> {1}",
        'simple_info_footer': "\nПрошло времени: {time_difference} минут\nОпубликовано: {user}",
        'no_info_found': "Информация по местоположению {location} не найдена.",
        'saved_successfully': "Сохранено!\nМестоположение: {location}\nДетали цветов: {msg}\nОпубликовано: {user_posted}",
        'message_forwarded': "Сообщение должно быть переадресовано от Chat Wars (@ChatWarsBot)!",
        'message_in_group': "Отчет получен, спасибо за вашу помощь! 🫡",
        'processing_info': "Обработка информации...",
        'invalid_message': "Неверный формат сообщения!",
        'get_excel_error': "Файл не найден.",
        'get_excel_error_exception': "Произошла ошибка: {error}",
        'map_error': "Ошибка при отправке изображения: {error}",
        'choose_language': "Пожалуйста, выберите предпочитаемый язык:",
        'default': 'ERROR:XXX>Извините, запрашиваемый текст недоступен.',
        'help_message': (
            "Помощь:\n",
            "Используйте /i + местоположение (например, gy2), чтобы получить полную информацию о конкретной плитке.\n"
            "Используйте /info + местоположение (например, y41), чтобы получить сохраненный текст в пятой ячейке этого местоположения.\n"
            "Используйте /get_excel, чтобы получить файл Excel с данными о местоположениях.\n"
            "Используйте /set_language, чтобы изменить язык бота."
        ),
    }
}



# Diccionario de idiomas soportados
SUPPORTED_LANGUAGES = ['en', 'es', 'ru']

# Diccionario para almacenar el idioma preferido por cada usuario
USER_LANGUAGES = {}

def detect_language(language_code):
    """Detecta el idioma del usuario basado en el código de idioma proporcionado por Telegram."""
    if language_code == None:
        return 'en'
    elif language_code.startswith('es'):
        return 'es'
    elif language_code.startswith('ru'):
        return 'ru'
    else:
        return 'en'

def get_text(update: Update, key):
    """Obtiene el texto en el idioma preferido del usuario o detecta el idioma automáticamente."""
    lang = ""
    if update.effective_user.id in USER_LANGUAGES:
        lang = USER_LANGUAGES[update.effective_user.id]
    else:
        lang = detect_language(update.effective_user.language_code)
        USER_LANGUAGES[update.effective_user.id] = lang  # Guardar la detección automática en el diccionario

    return TEXTS[lang].get(key, TEXTS[lang]['default'])