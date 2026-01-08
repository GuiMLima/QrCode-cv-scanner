import cv2
import pandas as pd
import numpy as np
from pyzbar.pyzbar import decode, ZBarSymbol
import datetime
import os
import time
import shutil

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
            
            self.df = pd.read_excel(
                self.filepath, 
                dtype=str,
                usecols=lambda x: x.strip() in cols # basic filter
            )
            
            # Verify if all columns were found; if not, retry loading all and filtering manually to catch case differences
            if len(self.df.columns) < len(cols):
                 print("Recarregando para buscar colunas case-insensitive...")
                 temp_df = pd.read_excel(self.filepath, dtype=str)
                 col_map = {}
                 for actual in temp_df.columns:
                     for target in cols:
                         if actual.strip().lower() == target.lower():
                             col_map[actual] = target
                 
                 if len(col_map) < len(cols):
                     print("ERRO CRÍTICO: Colunas obrigatórias faltando no Excel!")
                     print(f"Esperado: {cols}")
                     print(f"Encontrado mapeamento: {col_map}")
                 
                 temp_df.rename(columns=col_map, inplace=True)
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
        self.scanned_items = set() # To store scanned items in this session
        self.last_scan_time = {} # For debounce UI logic
        self.log_file = f"conferencia_log_{datetime.datetime.now().strftime('%Y-%m-%d')}.csv"
        self.font = cv2.FONT_HERSHEY_SIMPLEX
        
        # --- LOGGING SETUP ---
        if not os.path.exists(self.log_file):
            with open(self.log_file, "w", encoding='utf-8') as f:
                f.write("Timestamp,Rastreio,Status,Mensagem,Video_Evidence\n")
        else:
             with open(self.log_file, 'r', encoding='utf-8') as f:
                 header = f.readline()
             if "Video_Evidence" not in header:
                 df_log = pd.read_csv(self.log_file)
                 df_log["Video_Evidence"] = ""
                 df_log.to_csv(self.log_file, index=False)

        self.scan_results_cache = {} # code -> (status_text, header_color, rect_color, NF)
        
        # --- RECORDING SETUP (SCAN TRIGGERED) ---
        self.video_dir = "videos_auditoria"
        if not os.path.exists(self.video_dir):
            os.makedirs(self.video_dir)
            
        self.is_recording = False
        self.video_writer = None
        
        # Active Recording State
        self.current_recording_nf = None 
        self.last_nf_seen_time = 0
        self.post_scan_buffer = 3.0 # Segundos para gravar após a NF sair da tela
        self.current_video_filename = None

    def log_scan(self, tracking, status, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.log_file, "a", encoding='utf-8') as f:
            f.write(f"{timestamp},{tracking},{status},{message},\n")

    def _update_log_with_video(self, nf, video_filename):
        """
        Updates the CSV log to add the video filename for the specific NF.
        """
        if not nf:
            return

        try:
            df_log = pd.read_csv(self.log_file)
            
            # Simple heuristic: Look for rows where Mensagem contains "NF: <nf>" 
            # and status is SUCESSO.
            updated = False
            
            # Helper to check if row matches
            def is_match(row):
                msg = str(row['Mensagem'])
                return row['Status'] == 'SUCESSO' and (f"NF: {nf}" in msg or f"NF {nf}" in msg)

            for index, row in df_log.iterrows():
                if is_match(row):
                    # Only update if empty or overwrite? Let's overwrite/append
                    # If duplicate scans exist, we might update all or just latest. 
                    # Simplicity: Update all for this NF.
                    df_log.at[index, 'Video_Evidence'] = video_filename
                    updated = True
            
            if updated:
                df_log.to_csv(self.log_file, index=False)
                # print(f"Log atualizado para NF {nf}: {video_filename}")
                
        except Exception as e:
            print(f"Erro ao atualizar log com vídeo: {e}")

    def start_recording(self, nf):
        if self.is_recording:
            self.stop_recording()

        print(f"Iniciando Gravação para NF: {nf}")
        self.is_recording = True
        self.current_recording_nf = nf
        
        timestamp = datetime.datetime.now().strftime("%H%M%S")
        self.current_video_filename = f"NF_{nf}_{timestamp}.mp4"
        filepath = os.path.join(self.video_dir, self.current_video_filename)
        
        fourcc = cv2.VideoWriter_fourcc(*'mp4v') 
        self.video_writer = cv2.VideoWriter(filepath, fourcc, 20.0, (1280, 720))

    def stop_recording(self):
        if self.is_recording:
            print(f"Parando Gravação de {self.current_recording_nf}...")
            if self.video_writer:
                self.video_writer.release()
                self.video_writer = None
            
            # Update Log immediately
            if self.current_recording_nf:
                 self._update_log_with_video(self.current_recording_nf, self.current_video_filename)

            self.is_recording = False
            self.current_recording_nf = None
            self.current_video_filename = None

    def run(self):
        while True:
            success, img = self.cap.read()
            if not success:
                print("Erro ao acessar a webcam.")
                break

            current_time = time.time()
            decoded_objects = decode(img, symbols=[ZBarSymbol.QRCODE])
            
            # Header Layout
            cv2.rectangle(img, (0, 0), (1280, 80), (0, 0, 0), cv2.FILLED)
            current_header_text = "Aguardando Nota..."
            current_header_color = (255, 255, 255)

            # --- PROCESS DETECTED CODES ---
            valid_nf_in_frame = None # To track what we see NOW
            is_valid_nf_duplicate = False

            # MULTIPLE NFs CHECK
            # Check for UNIQUE codes. If we have multiple QRs but they are identical, it's fine.
            unique_codes_in_frame = set()
            for obj in decoded_objects:
                unique_codes_in_frame.add(obj.data.decode("utf-8"))

            if len(unique_codes_in_frame) > 1:
                current_header_text = "ERRO: Multiplas Notas Distintas! Deixe apenas uma."
                current_header_color = (0, 0, 255)
                
                # Draw Red Boxes on all
                for obj in decoded_objects:
                    pts = np.array([obj.polygon], np.int32).reshape((-1, 1, 2))
                    cv2.polylines(img, [pts], True, (0, 0, 255), 5)
                
                # SKIP PROCESSING
                
            else:
                # SINGLE OR NO OBJECT PROCESSING
                for obj in decoded_objects:
                    code_data = obj.data.decode("utf-8")
                    
                    # Default Visuals
                    status_text = "Processando..."
                    rect_color = (255, 255, 255)
                    header_color = (255, 255, 255)
                    found_nf = None
                    is_duplicate = False

                    # 1. Processing / Validation
                    # Check cache first to avoid re-querying dataframe every frame
                    if code_data in self.scan_results_cache:
                         status_text, header_color, rect_color, found_nf = self.scan_results_cache[code_data]
                         # Update Access Time
                         self.last_scan_time[code_data] = current_time
                         # Re-check duplicate status in real-time because scanned_items grows
                         if code_data in self.scanned_items:
                             is_duplicate = True
                             # Force update visual if it was cached as success but now is duplicate
                             if "OK:" in status_text:
                                 status_text = f"ALERTA: Pedido JA Conferido!"
                                 rect_color = (0, 255, 255)
                                 header_color = (0, 255, 255)
                                 # Update cache to reflect duplicate status
                                 self.scan_results_cache[code_data] = (status_text, header_color, rect_color, found_nf)

                    else:
                         # New Code Processing
                         self.last_scan_time[code_data] = current_time
                         result = self.data_loader.check_tracking(code_data)
                         
                         if result and result["found"]:
                             # FOUND
                             found_nf = result["nf"]
                             dest = result["destinatario"]
                             
                             if code_data in self.scanned_items:
                                 is_duplicate = True
                                 status_text = f"ALERTA: Pedido JA Conferido!"
                                 rect_color = (0, 255, 255) # Yellow
                                 header_color = (0, 255, 255)
                                 self.log_scan(code_data, "DUPLICADO", f"NF: {found_nf}")
                             else:
                                 is_duplicate = False
                                 status_text = f"OK: NF {found_nf} - {dest}"
                                 rect_color = (0, 255, 0) # Green
                                 header_color = (0, 255, 0)
                                 self.scanned_items.add(code_data)
                                 self.log_scan(code_data, "SUCESSO", f"NF: {found_nf}")
                         else:
                             # NOT FOUND
                             status_text = f"ERRO: Rastreio '{code_data}' Nao Consta"
                             rect_color = (0, 0, 255) # Red
                             header_color = (0, 0, 255)
                             found_nf = None
                             self.log_scan(code_data, "ERRO", "Rastreio nao encontrado")

                         # Save to Cache
                         self.scan_results_cache[code_data] = (status_text, header_color, rect_color, found_nf)

                    # 2. Visuals per Code
                    # OVERRIDE: If this is the NF we are currently recording, keep it GREEN!
                    if self.is_recording and found_nf == self.current_recording_nf:
                         status_text = f"NF {found_nf} - Gravando"
                         rect_color = (0, 255, 0)
                         header_color = (0, 255, 0)

                    # Draw Polygon
                    pts = np.array([obj.polygon], np.int32).reshape((-1, 1, 2))
                    cv2.polylines(img, [pts], True, rect_color, 5)
                    
                    # Update Header (Last code processed takes precedence on header text)
                    current_header_text = status_text
                    current_header_color = header_color
                    
                    # 3. Identify if this is a valid NF for Recording purposes
                    if found_nf:
                        valid_nf_in_frame = found_nf
                        # Important: If we are effectively treating it as "Green" because we are recording it,
                        # we should behave as if it's not a duplicate for the purpose of maintaining the session.
                        if self.is_recording and found_nf == self.current_recording_nf:
                            is_valid_nf_duplicate = False
                        else:
                            is_valid_nf_duplicate = is_duplicate

            # --- RECORDING STATE MACHINE ---
            
            if valid_nf_in_frame:
                # We see a valid NF right now
                self.last_nf_seen_time = current_time
                
                if not self.is_recording:
                    # START NEW RECORDING ONLY IF NOT DUPLICATE
                    if not is_valid_nf_duplicate:
                        self.start_recording(valid_nf_in_frame)
                    else:
                        # Duplicate and NOT recording: Ensure UI shows alert (handled above) but DO NOT record.
                        pass
                
                elif self.is_recording:
                    # ALREADY RECORDING
                    # Check if it is the SAME NF
                    if self.current_recording_nf == valid_nf_in_frame:
                        # Continue recording this session, even if it is now technically a duplicate in the set
                        pass
                    else:
                        # SWITCHING NF (A -> B)
                        print(f"Troca detectada: {self.current_recording_nf} -> {valid_nf_in_frame}")
                        self.stop_recording()
                        
                        # Only start B if it is NOT a duplicate
                        if not is_valid_nf_duplicate:
                            self.start_recording(valid_nf_in_frame)
            
            else:
                # No valid NF in this frame
                if self.is_recording:
                    # Check Buffer
                    if (current_time - self.last_nf_seen_time) > self.post_scan_buffer:
                        self.stop_recording()
                    else:
                        # Continue recording (Buffer Phase)
                        pass

            # Update Frame content (Header)
            cv2.putText(img, current_header_text, (20, 50), self.font, 1, current_header_color, 2)
            
            # Update Frame content (REC Indicator)
            if self.is_recording:
                 # Blinking Red Dot
                if int(current_time * 2) % 2 == 0:
                    cv2.circle(img, (1250, 50), 20, (0, 0, 255), cv2.FILLED)
                    cv2.putText(img, "REC", (1160, 60), self.font, 1, (0, 0, 255), 2)
                    # Show which NF is recording
                    cv2.putText(img, f"NF: {self.current_recording_nf}", (1100, 100), self.font, 0.7, (0, 0, 255), 2)

            # Write Frame if recording
            if self.is_recording and self.video_writer:
                self.video_writer.write(img)

            # Show
            cv2.imshow("Conferencia Gueddai", img)
            
            if cv2.waitKey(1) & 0xFF == ord('q'):
                if self.is_recording:
                    self.stop_recording()
                break

        self.cap.release()
        cv2.destroyAllWindows()

def get_latest_export_file(directory=".", prefix="Export_Order", extension=".xlsx"):
    try:
        files = [f for f in os.listdir(directory) if f.startswith(prefix) and f.endswith(extension)]
        if not files:
            return None
        files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)), reverse=True)
        return os.path.join(directory, files[0])
    except Exception as e:
        print(f"Erro ao buscar arquivos: {e}")
        return None

if __name__ == "__main__":
    latest_file = get_latest_export_file()
    
    if latest_file:
        print(f"Carregando arquivo mais recente encontrado: {latest_file}")
        loader = DataLoader(latest_file)
    else:
        print("Nenhum arquivo 'Export_Order*.xlsx' encontrado. Tentando padrao...")
        # Fallback file (make sure this file exists or logic handles missing file)
        loader = DataLoader("Export_Order_Fallback.xlsx")

    scanner = BarcodeScanner(loader)
    print("Iniciando Leitor com Gravacao por Sessao (Scan Trigger)...")
    print("Pressione 'q' para sair.")
    scanner.run()