# The viewer service in a Linux container. Publish it for linux-x64 into publish/, beside this file, first; the publish
# puts the engine, the two fonts the package carries, the component and the try-it page beside it:
#
#   dotnet publish server -c Release -r linux-x64 --self-contained false -o publish
#   docker build -t comofficekit-viewer .
#   docker run --rm -p 8080:8080 -e COMOFFICEKIT_LICENSE_KEY comofficekit-viewer
#
# Then open http://localhost:8080/. Settings go in as environment variables: -e AllowedOrigins=https://app.example.com.
#
# Microsoft's .NET 8 runtime image for web services is Debian 12, and listens on port 8080; the engine needs the
# distribution's libxml2 besides.
FROM mcr.microsoft.com/dotnet/aspnet:8.0
RUN apt-get update && apt-get install -y --no-install-recommends libxml2 && rm -rf /var/lib/apt/lists/*

# The image has no fonts of its own: documents are drawn in Carlito and Caladea, which the package carries, until
# their own families are installed. Liberation's have the widths of Arial, Times New Roman and Courier New; Noto
# CJK draws Chinese, Japanese and Korean.
# RUN apt-get update && apt-get install -y --no-install-recommends fonts-liberation2 fonts-noto-cjk && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY publish/ .
EXPOSE 8080
ENTRYPOINT ["/app/comofficekit-viewer-service"]
