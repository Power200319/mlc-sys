FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp

# LibreOffice Writer is required for DOCX-to-PDF conversion on Render. Noto
# fonts provide Khmer, Latin, and CJK glyph fallback for the school templates.
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-writer fontconfig fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN python manage.py collectstatic --noinput

RUN useradd --create-home --uid 10001 gradebook \
    && mkdir -p /app/runtime/jobs \
    && chown -R gradebook:gradebook /app
USER gradebook

CMD gunicorn backend.wsgi:application --bind 0.0.0.0:${PORT:-8000} --workers 1 --timeout 240
