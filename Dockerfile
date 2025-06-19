FROM python:3.12-slim
RUN apt-get update && apt-get install -y libreoffice curl fonts-dejavu-core && apt-get clean
WORKDIR /app
COPY . .
RUN curl -sL https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css -o static/tailwind.min.css
RUN pip install -r requirements.txt
ENV PPT_TEMPLATE_DIR=/app/templates_ppt
ENV WORKBOOK_XLSX=""
ENV UPLOAD_LIMIT=5
CMD ["python","app.py"]
