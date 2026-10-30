FROM python:3.10-slim

# Installation de LibreOffice et des dépendances système indispensables
RUN apt-get update && apt-get install -y \
    libreoffice \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Création du dossier de l'application
WORKDIR /code

# Copie et installation des requirements Python
COPY ./requirements.txt /code/requirements.txt
RUN pip install --no-cache-dir --upgrade -r /code/requirements.txt

# Copie du reste du code
COPY . .

# Commande pour démarrer Flask sur le port 7860 imposé par Hugging Face
CMD ["gunicorn", "-b", "0.0.0.0:7860", "app:app"]
