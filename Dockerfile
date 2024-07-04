FROM python:3.9
LABEL authors="akira"

WORKDIR /app

COPY Requirements /app/
RUN pip install -r Requirements

COPY SyncFileToCloud /app/
COPY SyncFileToCloud/. /app/SyncFileToCloud/
COPY main.py /app/

ENTRYPOINT python3 main.py