import cv2
import pandas as pd
import numpy as np
from pyzbar.pyzbar import decode, ZBarSymbol
import datetime
import os
import time

class DataLoader:
    def __init__(self, filepath):
        self.filepath = filepath
        self.df = None
        self.tracking_column = "Nº de Rastreio"
        self.nf_column = "Número da NF-e"
        self.dest_column = "Nome do Destinatário"
        self.load_data()

    def load_data(self):
        if not os.path.exists(self.filepath):
            print(f"ALERTA: Arquivo '{self.filepath}' não encontrado! O programa continuará, mas a validação falhará.")
            self.df = pd.DataFrame(columns=[self.tracking_column, self.nf_column, self.dest_column])
            return

        try:
            # Load only necessary columns, as string to preserve leading zeros
            cols = [self.tracking_column, self.nf_column, self.dest_column]
            
            # Using usecols with a custom function for case-insensitive matching if needed, 
            # but for now we rely on the exact names provided in the requirement.
            # If names might vary in case, we'd need a more robust loading strategy.
            # For this implementation, we assume headers are exact or we verify them.
            
            self.df = pd.read_excel(
                self.filepath, 
                dtype=str,
                usecols=lambda x: x.strip() in cols # basic filter
            )
            
            # Verify if all columns were found; if not, retry loading all and filtering manually to catch case differences
            if len(self.df.columns) < len(cols):
                 print("Recarregando para buscar colunas case-insensitive...")
                 temp_df = pd.read_excel(self.filepath, dtype=str)
                 # Map actual columns to target columns (simple case-insensitive match)
                 col_map = {}
                 for actual in temp_df.columns:
                     for target in cols:
                         if actual.strip().lower() == target.lower():
                             col_map[actual] = target
                 
                 if len(col_map) < len(cols):
                     print("ERRO CRÍTICO: Colunas obrigatórias faltando no Excel!")
                     print(f"Esperado: {cols}")
                     print(f"Encontrado mapeamento: {col_map}")
                 
                 # Rename to standard names
                 temp_df.rename(columns=col_map, inplace=True)
                 # Keep only valid columns
                 self.df = temp_df[cols]

            # Remove rows with empty tracking numbers
            self.df.dropna(subset=[self.tracking_column], inplace=True)
            self.df[self.tracking_column] = self.df[self.tracking_column].astype(str).str.strip()
            
            print(f"Sucesso: {len(self.df)} registros carregados.")
            
        except Exception as e:
            print(f"Erro ao carregar Excel: {e}")
            self.df = pd.DataFrame(columns=[cols])

    def check_tracking(self, tracking_code):
        if self.df is None or self.df.empty:
            return None
        
        # Search for the tracking code
        result = self.df[self.df[self.tracking_column] == tracking_code]
        
        if not result.empty:
            row = result.iloc[0]
            return {
                "nf": row[self.nf_column],
                "destinatario": row[self.dest_column],
                "found": True
            }
        return {"found": False}

class BarcodeScanner:
    def __init__(self, data_loader):
        self.cap = cv2.VideoCapture(0)
        self.cap.set(3, 1280) # Width
        self.cap.set(4, 720)  # Height
        self.data_loader = data_loader
        self.scanned_items = set() # To store scanned items in this session for duplicate check
        self.last_scan_time = {} # For debounce logic
        self.log_file = f"conferencia_log_{datetime.datetime.now().strftime('%Y-%m-%d')}.csv"
        self.font = cv2.FONT_HERSHEY_SIMPLEX
        
        # Ensure log file exists with headers
        if not os.path.exists(self.log_file):
            with open(self.log_file, "w") as f:
                f.write("Timestamp,Rastreio,Status,Mensagem\n")

        self.scan_results_cache = {} # code -> (status_text, header_color, rect_color)

    def log_scan(self, tracking, status, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.log_file, "a") as f:
            f.write(f"{timestamp},{tracking},{status},{message}\n")

    def run(self):
        while True:
            success, img = self.cap.read()
            if not success:
                print("Erro ao acessar a webcam.")
                break

            # Decode QR codes
            decoded_objects = decode(img, symbols=[ZBarSymbol.QRCODE])

            # Header Background
            cv2.rectangle(img, (0, 0), (1280, 80), (0, 0, 0), cv2.FILLED)
            
            # Default Header (Reset if no QR is seen)
            current_header_text = "Aguardando Leitura..."
            current_header_color = (255, 255, 255)

            for obj in decoded_objects:
                code_data = obj.data.decode("utf-8")
                current_time = time.time()
                
                # Check if this code is currently being held/viewed (Debounce/Cache)
                # If we saw it recently (< 3.0s), explicitly reuse the last result logic
                # and extend the timer so it doesn't "expire" while holding.
                if code_data in self.last_scan_time and (current_time - self.last_scan_time[code_data] < 3.0):
                     # Existing active scan session for this code
                     if code_data in self.scan_results_cache:
                         status_text, header_color, rect_color = self.scan_results_cache[code_data]
                         # Keep "alive"
                         self.last_scan_time[code_data] = current_time 
                     else:
                         # Fallback if cache missing (shouldn't happen usually)
                         status_text = "Processando..."
                         header_color = (255, 255, 255)
                         rect_color = (255, 255, 255)
                else:
                     # NEW SCAN event (or re-scan after timeout)
                     self.last_scan_time[code_data] = current_time
                     
                     # Validate
                     result = self.data_loader.check_tracking(code_data)
                     
                     rect_color = (0, 0, 255) # Red
                     status_text = f"ERRO: Rastreio '{code_data}' Nao Consta"
                     header_color = (0, 0, 255)
                     log_status = "ERRO"
                     
                     if result and result["found"]:
                        if code_data in self.scanned_items:
                            # DUPLICATE
                            rect_color = (0, 255, 255) # Yellow
                            status_text = f"ALERTA: Pedido JA Conferido!"
                            header_color = (0, 255, 255)
                            log_status = "DUPLICADO"
                        else:
                            # SUCCESS
                            rect_color = (0, 255, 0) # Green
                            nf = result["nf"]
                            dest = result["destinatario"]
                            status_text = f"OK: NF {nf} - {dest}"
                            header_color = (0, 255, 0)
                            log_status = "SUCESSO"
                            
                            self.scanned_items.add(code_data)
                            self.log_scan(code_data, log_status, f"NF: {nf}")
                     else:
                        self.log_scan(code_data, log_status, "Rastreio nao encontrado na lista")
                     
                     # Cache the result for this session
                     self.scan_results_cache[code_data] = (status_text, header_color, rect_color)

                # Set current frame header to this code's status
                current_header_text = status_text
                current_header_color = header_color

                # Draw Polygon
                pts = np.array([obj.polygon], np.int32)
                pts = pts.reshape((-1, 1, 2))
                cv2.polylines(img, [pts], True, rect_color, 5)

            # Draw Header Text
            cv2.putText(img, current_header_text, (20, 50), self.font, 1, current_header_color, 2)

            cv2.imshow("Conferencia Gueddai", img)
            
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        self.cap.release()
        cv2.destroyAllWindows()

if __name__ == "__main__":
    loader = DataLoader("Export_Order202601061652.xlsx")
    scanner = BarcodeScanner(loader)
    scanner.run()