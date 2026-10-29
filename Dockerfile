FROM python:3.11-slim

ENV TZ=Europe/Moscow \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Весь проект (код, расписание, меню, кэши, сертификат) монтируется volume'ом
# снаружи при запуске (см. README, разделы 2–3) — это позволяет обновлять
# расписание/меню без пересборки образа.
VOLUME ["/app"]

CMD ["python3", "school_bot.py"]
