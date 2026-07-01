FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads outputs

ENV FLASK_DEBUG=0
ENV PORT=10000

CMD gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
