FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app/ ./app/

# Generate self-signed SSL certs for HTTPS
RUN mkdir -p /certs && \
    openssl req -x509 -newkey rsa:2048 -nodes \
      -keyout /certs/localhost.key \
      -out /certs/localhost.crt \
      -days 365 \
      -subj "/CN=localhost" \
      -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"

RUN mkdir -p /app/app/data

EXPOSE 8000

CMD ["python", "-m", "uvicorn", "app.main:app", \
     "--host", "0.0.0.0", "--port", "8000", \
     "--ssl-certfile", "/certs/localhost.crt", \
     "--ssl-keyfile", "/certs/localhost.key"]
