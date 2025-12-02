@echo off
REM Google Cloud Run Deployment Script for Windows
REM Usage: deploy.bat PROJECT_ID REGION

SET PROJECT_ID=%1
SET REGION=%2

IF "%PROJECT_ID%"=="" SET PROJECT_ID=your-project-id
IF "%REGION%"=="" SET REGION=us-central1

SET SERVICE_NAME=document-processing-api
SET IMAGE_NAME=gcr.io/%PROJECT_ID%/%SERVICE_NAME%

echo ==========================================
echo Deploying to Google Cloud Run
echo ==========================================
echo Project: %PROJECT_ID%
echo Region: %REGION%
echo Service: %SERVICE_NAME%
echo ==========================================

echo.
echo [1/3] Building Docker image...
docker build -t %IMAGE_NAME% .

IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Docker build failed!
    exit /b 1
)

echo.
echo [2/3] Pushing to Google Container Registry...
docker push %IMAGE_NAME%

IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Docker push failed!
    echo Make sure you've run: gcloud auth configure-docker
    exit /b 1
)

echo.
echo [3/3] Deploying to Cloud Run...
gcloud run deploy %SERVICE_NAME% ^
    --image %IMAGE_NAME% ^
    --platform managed ^
    --region %REGION% ^
    --allow-unauthenticated ^
    --memory 1Gi ^
    --cpu 1 ^
    --timeout 300 ^
    --max-instances 10 ^
    --set-env-vars "PYTHONUNBUFFERED=1"

IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Cloud Run deployment failed!
    exit /b 1
)

echo.
echo ==========================================
echo Deployment successful!
echo ==========================================
echo.
echo Get your service URL with:
echo   gcloud run services describe %SERVICE_NAME% --region %REGION% --format "value(status.url)"
