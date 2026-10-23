FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8 \
    PORT=7860 \
    JOBS_DIR=/tmp/jobs

WORKDIR /app
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY . /app
RUN mkdir -p /tmp/jobs && chmod -R 777 /tmp/jobs /app

EXPOSE 7860
CMD ["python", "/app/app/server.py"]
