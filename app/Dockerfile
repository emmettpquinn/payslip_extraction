FROM python:3.11-slim

WORKDIR /app

# Copy poetry dependency files
COPY pyproject.toml poetry.lock* /app/

# Install poetry and dependencies
RUN pip install poetry && \
    poetry config virtualenvs.create false && \
    poetry install --only=main --no-root


# Copy app code and .env file
COPY main.py /app/
COPY credentials.json /app/
COPY .env /app/

# Environment variable for Google credentials JSON path inside container
ENV GOOGLE_APPLICATION_CREDENTIALS=/app/credentials.json

CMD ["python", "main.py"]

