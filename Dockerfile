FROM heroku/heroku:22

RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    libreoffice \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY . /app

RUN pip install --upgrade pip && \ 
    pip install -r requirements.txt

EXPOSE 5000

CMD ["sh", "-c", "gunicorn -b 0.0.0.0:$PORT --timeout 240 app:app"]
