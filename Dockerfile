# Stage 1: Build Frontend
FROM node:22-alpine as frontend-build
WORKDIR /app/frontend
COPY frontend/package*.json ./
RUN npm install
COPY frontend/ ./
RUN npm run build

# Stage 2: Final Image
FROM python:3.11-slim

# Install system dependencies and Azure CLI
RUN apt-get update && apt-get install -y curl gnupg lsb-release \
    && curl -sL https://aka.ms/InstallAzureCLIDeb | bash \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy Backend Code
COPY backend/ ./backend/
COPY script.py .
COPY config.py .
COPY providers_raw.json .
COPY iam_roles_extractor.py .
COPY vm_spot_extractor.py .

# Local Azure icon assets — referenced by backend/architecture_generator.py
# (get_icon_uri) so diagrams embed icons as base64 data URIs instead of
# fetching them at render time from azureicons.com.
COPY Azure-Icons/ ./Azure-Icons/

# Copy Built Frontend from Stage 1
COPY --from=frontend-build /app/frontend/dist ./frontend/dist

# Expose port
EXPOSE 8000

# Run the application
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]
