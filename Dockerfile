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

CMD ["gunicorn", "-b", "0.0.0.0:5000", "--timeout", "240", "app:app"]
