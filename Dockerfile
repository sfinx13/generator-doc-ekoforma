FROM heroku/heroku:22

RUN apt-get update && apt-get install -y \
    libreoffice \
    && apt-get clean

WORKDIR /app

COPY . /app

RUN pip install --upgrade pip
RUN pip install -r requirements.txt

EXPOSE 5000

CMD ["gunicorn", "-b", "0.0.0.0:5000", "app:app"]
