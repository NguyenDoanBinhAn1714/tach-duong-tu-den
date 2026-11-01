FROM python:3.12-slim

WORKDIR /app
COPY requirements.txt requirements-server.txt ./
RUN pip install --no-cache-dir -r requirements-server.txt

COPY core ./core
COPY server ./server
COPY run_server.py ./

# Trong container phải nghe trên 0.0.0.0 thì máy ngoài mới truy cập được.
ENV HOST=0.0.0.0 PORT=10000
EXPOSE 10000
CMD ["python", "run_server.py"]
