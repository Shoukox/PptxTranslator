FROM mcr.microsoft.com/dotnet/sdk:10.0-preview AS build
WORKDIR /src

COPY ["PptxTranslator.Web/PptxTranslator.Web.csproj", "PptxTranslator.Web/"]
COPY ["PptxTranslator.ServiceDefaults/PptxTranslator.ServiceDefaults.csproj", "PptxTranslator.ServiceDefaults/"]
RUN dotnet restore "PptxTranslator.Web/PptxTranslator.Web.csproj"

COPY . .
WORKDIR "/src/PptxTranslator.Web"
RUN dotnet publish "PptxTranslator.Web.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM mcr.microsoft.com/dotnet/aspnet:10.0-preview AS final
WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends python3 python3-pip \
    && rm -rf /var/lib/apt/lists/*

COPY PptxTranslator.Web/requirements.txt /tmp/requirements.txt
RUN python3 -m pip install --break-system-packages --no-cache-dir -r /tmp/requirements.txt \
    && rm /tmp/requirements.txt
COPY PptxTranslator.Web/Scripts/pptx_translate_ru.py /app/Scripts/pptx_translate_ru.py

COPY --from=build /app/publish .

ENV ASPNETCORE_URLS=http://+:8080
ENV HTTP_PORTS=8080
ENV PPTX_TRANSLATOR_PYTHON=python3

EXPOSE 8080

ENTRYPOINT ["dotnet", "PptxTranslator.Web.dll"]
