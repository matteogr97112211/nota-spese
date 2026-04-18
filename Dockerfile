FROM python:3.10
RUN apt-get update && apt-get install -y tesseract-ocr
WORKDIR /app
COPY . .
RUN pip install -r requirements.txt
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000"]