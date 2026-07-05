FROM python:3.12-slim
# cups-client = `lp` (submits to host cupsd over IPP); libreoffice = Office->PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
      cups-client libreoffice-nogui fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/* \
    && useradd -m -u 10001 poller
WORKDIR /app
COPY poller.py .
ENV PYTHONUNBUFFERED=1 HOME=/home/poller
USER poller
CMD ["python3", "poller.py"]
