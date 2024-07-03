FROM python:3.9
LABEL authors="akira"

WORKDIR /app

COPY Requirements /app/
RUN pip install -r Requirements

COPY SyncFileTo365.py SyncFileToGcal.py main.py /app/
COPY config.py-EXAMPLE /config/

ENTRYPOINT python3 main.py