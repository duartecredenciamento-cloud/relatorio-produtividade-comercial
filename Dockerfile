FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# O Railway fornece a porta em PORT; o servidor da aplicação deve usá-la
# e escutar em 0.0.0.0 para aceitar conexões externas.
CMD ["sh", "-c", "python app_server.py"]
