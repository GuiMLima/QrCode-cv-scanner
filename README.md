# Conferência Gueddai / Gueddai Conference

## 🇧🇷 Português

### Descrição
Este projeto é uma ferramenta de visão computacional desenvolvida em Python para validação de envios de pacotes de e-commerce. O sistema utiliza uma webcam para escanear códigos de rastreio em tempo real e validá-los contra uma planilha de pedidos diária.

### Funcionalidades
- **Leitura de Código de Barras/QR:** Detecção e leitura em tempo real via webcam.
- **Validação de Dados:** Comparação instantânea com base de dados em Excel (`.xlsx`).
- **Feedback Visual:**
  - 🟢 **Sucesso:** Pedido encontrado e validado com exibição de NF e Destinatário.
  - 🟡 **Alerta:** Pedido duplicado (já conferido na sessão atual).
  - 🔴 **Erro:** Código de rastreio não encontrado na lista.
- **Gravação Inteligente:** 
  - 🎥 Grava automaticamente um curto vídeo de evidência para cada NF validada. 
  - O vídeo inicia ao detectar a NF e encerra automaticante 3s após a saída do pacote.
- **Registro de Logs:** Geração automática de relatórios de conferência em CSV (incluindo nome do arquivo de vídeo).
- **Auto-Loader:** Detecção automática da planilha de pedidos mais recente na pasta.

### Requisitos
- Python 3.x
- Bibliotecas: `opencv-python`, `pandas`, `numpy`, `pyzbar`, `openpyxl`
- Arquivo de dados: `Export_Order...xlsx` (deve estar na mesma pasta)

---

## 🇺🇸 English

### Description
This project is a computer vision tool developed in Python for validating e-commerce package shipments. The system uses a webcam to scan tracking codes in real-time and verifies them against a daily order spreadsheet.

### Features
- **Barcode/QR Scanning:** Real-time detection and reading via webcam.
- **Data Validation:** Instant comparison against an Excel database (`.xlsx`).
- **Visual Feedback:**
  - 🟢 **Success:** Order found and validated, showing Invoice # and Recipient.
  - 🟡 **Alert:** Duplicate scan (already checked in current session).
  - 🔴 **Error:** Tracking code not found in the list.
- **Smart Recording:**
  - 🎥 Automatically records a short evidence video for each validated Invoice (NF).
  - Recording starts upon detection and stops 3s after the package leaves the frame.
- **Logging:** Automatic generation of conference reports in CSV format (including video filename).
- **Auto-Loader:** Automatically detects the most recent order spreadsheet in the folder.

### Requirements
- Python 3.x
- Libraries: `opencv-python`, `pandas`, `numpy`, `pyzbar`, `openpyxl`
- Data file: `Export_Order...xlsx` (must be in the same folder)

---

## Sobre / About
*Este projeto baseia-se em um leitor de QR Code através de vídeo. O objetivo principal é trabalhar conceitos de engenharia de computação e visão computacional.*
*This project is based on a video QR Code reader. The main goal is to apply computer engineering and computer vision concepts.*
