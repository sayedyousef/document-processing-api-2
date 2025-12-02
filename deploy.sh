#!/bin/bash

# Google Cloud Run Deployment Script
# Usage: ./deploy.sh [PROJECT_ID] [REGION]

PROJECT_ID=${1:-"your-project-id"}
REGION=${2:-"us-central1"}
SERVICE_NAME="document-processing-api"
IMAGE_NAME="gcr.io/${PROJECT_ID}/${SERVICE_NAME}"

echo "=========================================="
echo "Deploying to Google Cloud Run"
echo "=========================================="
echo "Project: ${PROJECT_ID}"
echo "Region: ${REGION}"
echo "Service: ${SERVICE_NAME}"
echo "=========================================="

# Step 1: Build the Docker image
echo ""
echo "[1/3] Building Docker image..."
docker build -t ${IMAGE_NAME} .

if [ $? -ne 0 ]; then
    echo "ERROR: Docker build failed!"
    exit 1
fi

# Step 2: Push to Google Container Registry
echo ""
echo "[2/3] Pushing to Google Container Registry..."
docker push ${IMAGE_NAME}

if [ $? -ne 0 ]; then
    echo "ERROR: Docker push failed!"
    echo "Make sure you've run: gcloud auth configure-docker"
    exit 1
fi

# Step 3: Deploy to Cloud Run
echo ""
echo "[3/3] Deploying to Cloud Run..."
gcloud run deploy ${SERVICE_NAME} \
    --image ${IMAGE_NAME} \
    --platform managed \
    --region ${REGION} \
    --allow-unauthenticated \
    --memory 1Gi \
    --cpu 1 \
    --timeout 300 \
    --max-instances 10 \
    --set-env-vars "PYTHONUNBUFFERED=1"

if [ $? -ne 0 ]; then
    echo "ERROR: Cloud Run deployment failed!"
    exit 1
fi

echo ""
echo "=========================================="
echo "Deployment successful!"
echo "=========================================="
echo ""
echo "Get your service URL with:"
echo "  gcloud run services describe ${SERVICE_NAME} --region ${REGION} --format 'value(status.url)'"
