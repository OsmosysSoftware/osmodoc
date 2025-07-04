# Use the official .NET SDK image as the build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# Build configuration variable. Debug for development, Release for production
ENV BUILD_CONFIGURATION=Release
ENV ASPNETCORE_ENVIRONMENT=Production

# Copy the solution and project files
COPY *.sln ./
COPY ./OsmoDoc.API/OsmoDoc.API.csproj OsmoDoc.API/
COPY ./OsmoDoc/OsmoDoc.csproj OsmoDoc/

# Restore dependencies
RUN dotnet restore OsmoDoc.API/OsmoDoc.API.csproj

# Copy required project files
COPY ./OsmoDoc.API/ OsmoDoc.API/
COPY ./OsmoDoc/ OsmoDoc/
COPY .env .

# Build the project and store artifacts in /out folder
RUN dotnet publish OsmoDoc.API/OsmoDoc.API.csproj -c $BUILD_CONFIGURATION -o out

# Use the official ASP.NET runtime image as the base image
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base

WORKDIR /app

ENV ASPNETCORE_URLS=http://+:5000

# Copy the published artifacts from the build stage
COPY --from=build /app/out .

# Install only necessary dependencies for wkhtmltopdf, Node.js and npm
RUN apt-get update && apt-get install -y --no-install-recommends \
    wkhtmltopdf \
    nodejs \
    npm && \
    rm -rf /var/lib/apt/lists/*

# Install wkhtmltopdf and allow execute access
RUN chmod 755 /usr/bin/wkhtmltopdf

# Install EJS globally
RUN npm install -g --only=prod ejs

# Expose the API port
EXPOSE 5000

# Set the entry point for the container
ENTRYPOINT ["dotnet", "OsmoDoc.API.dll"]