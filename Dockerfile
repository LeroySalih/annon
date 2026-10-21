FROM python:3.12-slim

# LibreOffice is only used to turn DOCX assignments into PDF.
RUN apt-get update \
 && apt-get install -y --no-install-recommends libreoffice-writer fonts-liberation fonts-dejavu-core \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn
COPY annon ./annon

RUN useradd --create-home --uid 10001 app
USER app
ENV HOME=/home/app PYTHONUNBUFFERED=1
EXPOSE 5055

# One worker: the pupil-ID mapping is held in memory between the two downloads.
CMD ["gunicorn", "--workers", "1", "--threads", "4", "--timeout", "300", "--bind", "0.0.0.0:5055", "annon.app:app"]
