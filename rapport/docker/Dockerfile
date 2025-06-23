FROM python:3.11-slim

RUN apt-get update && \
    apt-get install -y build-essential libffi-dev libpango1.0-0 libpangocairo-1.0-0 \
    libcairo2 libcairo2-dev libgdk-pixbuf2.0-0 libgdk-pixbuf2.0-dev libxml2 libxslt1.1 \
    libjpeg62-turbo fonts-liberation && \
    pip install weasyprint && \
    apt-get clean

WORKDIR /app
