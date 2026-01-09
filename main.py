import tkinter as tk
from tkinter import filedialog, messagebox, Tk, ttk
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    print("Aviso: tkinterdnd2 nao instalado. Drag and Drop desativado.")
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
        
        """
        Descrição: Inicializa o carregador de dados com o caminho do arquivo.
        Description: Initializes the data loader with the file path.
        """

    def load_data(self):
        """
        Descrição: Carrega e processa os dados da planilha Excel.
        Description: Loads and processes data from the Excel spreadsheet.
        """
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
        """
        Descrição: Verifica se o código de rastreio existe na base de dados.
        Description: Checks if the tracking code exists in the database.
        """
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

class ScannerButton:
    def __init__(self, text, x, y, w, h, bg_color, text_color):
        self.text = text
        self.rect = (x, y, w, h) 
        self.bg_color = bg_color
        self.text_color = text_color
    
    def draw(self, img):
        x, y, w, h = self.rect
        cv2.rectangle(img, (x, y), (x+w, y+h), self.bg_color, cv2.FILLED)
        cv2.putText(img, self.text, (x + 10, y + 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, self.text_color, 2)

    def is_clicked(self, cx, cy):
        bx, by, bw, bh = self.rect
        return bx <= cx <= bx+bw and by <= cy <= by+bh

class BarcodeScanner:
    def __init__(self, data_loader, video_path="videos_auditoria", report_path="."):
        self.cap = cv2.VideoCapture(0)
        self.cap.set(3, 1280) # Width
        self.cap.set(4, 720)  # Height
        self.data_loader = data_loader
        self.scanned_items = set() # To store scanned items in this session
        self.last_scan_time = {} # For debounce UI logic
        
        # --- PATH CONFIGURATION ---
        self.report_dir = report_path
        if not os.path.exists(self.report_dir):
            os.makedirs(self.report_dir)
            
        self.video_dir = video_path
        if not os.path.exists(self.video_dir):
            os.makedirs(self.video_dir)
            
        self.log_file = os.path.join(self.report_dir, f"conferencia_log_{datetime.datetime.now().strftime('%Y-%m-%d')}.csv")
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
        
        self.load_scanned_items()

        self.scan_results_cache = {} # code -> (status_text, header_color, rect_color, NF)
        
        # --- NAVIGATION STATE ---
        self.nav_action = None # 'home', 'gallery', or None (quit)
        self.buttons = []

        # --- RECORDING SETUP ---
        self.is_recording = False
        self.video_writer = None
        
        # Active Recording State
        self.current_recording_nf = None 
        self.last_nf_seen_time = 0
        self.post_scan_buffer = 3.0 # Segundos para gravar após a NF sair da tela
        self.current_video_filename = None
        
        # Stability Check State
        self.current_candidate_nf = None
        self.candidate_start_time = 0
        self.stability_duration = 2.0 # Seconds to hold before recording starts

    def load_scanned_items(self):
        """
        Descrição: Carrega itens já conferidos (Status=SUCESSO) do log de hoje para evitar duplicatas.
        Description: Loads already checked items (Status=SUCESSO) from today's log to prevent duplicates.
        """
        if not os.path.exists(self.log_file):
            return

        try:
            df = pd.read_csv(self.log_file)
            # Filter for SUCESSO
            if not df.empty and "Status" in df.columns and "Rastreio" in df.columns:
                 success_rows = df[df["Status"] == "SUCESSO"]
                 for code in success_rows["Rastreio"].astype(str):
                     self.scanned_items.add(code.strip())
            print(f"Log carregado. {len(self.scanned_items)} itens já conferidos.")
        except Exception as e:
            print(f"Erro ao carregar log de duplicatas: {e}")

    def log_scan(self, tracking, status, message):
        """
        Descrição: Registra uma operação de escaneamento no arquivo de log CSV.
        Description: Logs a scan operation to the CSV log file.
        """
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.log_file, "a", encoding='utf-8') as f:
            f.write(f"{timestamp},{tracking},{status},{message},\n")

    def _update_log_with_video(self, nf, video_filename):
        """
        Descrição: Atualiza o log CSV para adicionar o nome do arquivo de vídeo para uma NF específica.
        Description: Updates the CSV log to add the video filename for a specific NF.
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
        """
        Descrição: Inicia a gravação de vídeo para uma Nota Fiscal.
        Description: Starts video recording for a Invoice (NF).
        """
        if self.is_recording:
            self.stop_recording()

        print(f"Iniciando Gravação para NF: {nf}")
        self.is_recording = True
        self.current_recording_nf = nf
        
        # Format: NF{NUMBER}.mp4 as requested
        self.current_video_filename = f"NF{nf}.mp4"
        filepath = os.path.join(self.video_dir, self.current_video_filename)
        
        fourcc = cv2.VideoWriter_fourcc(*'mp4v') 
        self.video_writer = cv2.VideoWriter(filepath, fourcc, 20.0, (1280, 720))

    def stop_recording(self):
        """
        Descrição: Para a gravação atual e finaliza o arquivo de vídeo.
        Description: Stops the current recording and finalizes the video file.
        """
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

    def _mouse_callback(self, event, x, y, flags, param):
        if event == cv2.EVENT_LBUTTONDOWN:
            for btn in self.buttons:
                if btn.is_clicked(x, y):
                    if btn.text == "HOME":
                         self.nav_action = "home"
                    elif btn.text == "VIDEOS":
                         self.nav_action = "gallery"

    def run(self):
        """
        Descrição: Loop principal de captura de vídeo e processamento de QR Codes.
        Description: Main loop for video capture and QR Code processing.
        """
        # Define Buttons
        # Bottom Left for Navigation
        self.buttons = [
            ScannerButton("HOME", 20, 650, 100, 50, (50, 50, 50), (255, 255, 255)),
            ScannerButton("VIDEOS", 140, 650, 120, 50, (50, 50, 50), (255, 255, 255))
        ]
        
        window_name = "Conferencia Gueddai"
        cv2.namedWindow(window_name)
        cv2.setMouseCallback(window_name, self._mouse_callback)

        while True:
            # Check external navigation request
            if self.nav_action:
                if self.is_recording:
                    self.stop_recording()
                break

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
            valid_tracking_code_in_frame = None 
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
                             if "Segure" in status_text or "OK:" in status_text:
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
                                 # WAIT FOR HOLD - Do NOT commit yet
                                 status_text = f"Identificado: {found_nf} - Segure..."
                                 rect_color = (0, 255, 255) # Yellow (Wait)
                                 header_color = (0, 255, 255)
                                 # Do NOT add to scanned_items yet
                                 # Do NOT log yet
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
                        valid_tracking_code_in_frame = code_data
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
                
                # --- STABILITY CHECK ---
                if self.current_candidate_nf != valid_nf_in_frame:
                    self.current_candidate_nf = valid_nf_in_frame
                    self.candidate_start_time = current_time
                
                # Calculate how long we have been staring at this specific NF
                elapsed_hold = current_time - self.candidate_start_time
                
                # Visual Feedback for Hold
                if elapsed_hold < self.stability_duration and not self.is_recording and not is_valid_nf_duplicate:
                    # Show progress
                    pct = int((elapsed_hold / self.stability_duration) * 100)
                    cv2.putText(img, f"Segure... {pct}%", (20, 100), self.font, 0.7, (0, 255, 255), 2)
                
                # Trigger Condition
                should_start = (elapsed_hold >= self.stability_duration)
                
                if not self.is_recording:
                    # START NEW RECORDING ONLY IF NOT DUPLICATE AND STABLE
                    if not is_valid_nf_duplicate:
                        if should_start:
                            # COMMIT SCAN HERE
                            if valid_tracking_code_in_frame and valid_tracking_code_in_frame not in self.scanned_items:
                                self.scanned_items.add(valid_tracking_code_in_frame)
                                self.log_scan(valid_tracking_code_in_frame, "SUCESSO", f"NF: {valid_nf_in_frame}")
                                
                            self.start_recording(valid_nf_in_frame)
                    else:
                        pass # Duplicate handling
                
                elif self.is_recording:
                    # ALREADY RECORDING
                    # Check if it is the SAME NF
                    if self.current_recording_nf == valid_nf_in_frame:
                         pass
                    else:
                        # SWITCHING NF (A -> B)
                        if should_start:
                             print(f"Troca detectada: {self.current_recording_nf} -> {valid_nf_in_frame}")
                             self.stop_recording()
                             if not is_valid_nf_duplicate:
                                 # COMMIT SCAN HERE (SWITCH CASE)
                                 if valid_tracking_code_in_frame and valid_tracking_code_in_frame not in self.scanned_items:
                                     self.scanned_items.add(valid_tracking_code_in_frame)
                                     self.log_scan(valid_tracking_code_in_frame, "SUCESSO", f"NF: {valid_nf_in_frame}")
                                 
                                 self.start_recording(valid_nf_in_frame)
            
            else:
                # No valid NF in this frame
                self.current_candidate_nf = None # Reset candidate
                
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

            # Draw Navigation Buttons
            for btn in self.buttons:
                btn.draw(img)

            # Write Frame if recording
            if self.is_recording and self.video_writer:
                self.video_writer.write(img)

            # Show
            cv2.imshow(window_name, img)
            
            if cv2.waitKey(1) & 0xFF == ord('q'):
                if self.is_recording:
                    self.stop_recording()
                break

        self.cap.release()
        cv2.destroyAllWindows()
        return self.nav_action



        self.cap.release()
        cv2.destroyAllWindows()


def rounded_rect(canvas, x, y, w, h, c, bg_color):
    """
    Descrição: Desenha um retângulo com bordas arredondadas em um canvas Tkinter.
    Description: Draws a rounded rectangle on a Tkinter canvas.
    """
    canvas.create_oval(x, y, x + c * 2, y + c * 2, fill=bg_color, outline=bg_color)
    canvas.create_oval(x + w - c * 2, y, x + w, y + c * 2, fill=bg_color, outline=bg_color)
    canvas.create_oval(x, y + h - c * 2, x + c * 2, y + h, fill=bg_color, outline=bg_color)
    canvas.create_oval(x + w - c * 2, y + h - c * 2, x + w, y + h, fill=bg_color, outline=bg_color)
    canvas.create_rectangle(x + c, y, x + w - c, y + h, fill=bg_color, outline=bg_color)
    canvas.create_rectangle(x, y + c, x + w, y + h - c, fill=bg_color, outline=bg_color)

class RoundedEntry(tk.Canvas):
    """
    Descrição: Widget customizado de entrada de texto com bordas arredondadas e placeholder.
    Description: Custom text entry widget with rounded corners and placeholder support.
    """
    def __init__(self, parent, width, height, corner_radius, padding=10, placeholder_text="", **kwargs):
        tk.Canvas.__init__(self, parent, width=width, height=height, borderwidth=0, highlightthickness=0, **kwargs)
        self.width = width
        self.height = height
        self.padding = padding
        self.placeholder_text = placeholder_text
        self.placeholder_color = "#555555" # Darker gray as requested
        self.text_color = "#000000"
        
        # Draw background
        rounded_rect(self, 0, 0, width, height, corner_radius, "white")
        
        # Embedded Entry
        self.entry = tk.Entry(self, bg="white", bd=0, font=("Segoe UI", 11), highlightthickness=0, fg=self.placeholder_color)
        self.entry_window = self.create_window(padding, height//2, window=self.entry, anchor="w", width=width-padding*2)
        
        if self.placeholder_text:
            self.entry.insert(0, self.placeholder_text)
            
        self.entry.bind("<FocusIn>", self._on_focus_in)
        self.entry.bind("<FocusOut>", self._on_focus_out)

    def _on_focus_in(self, event):
        if self.entry.get() == self.placeholder_text:
            self.entry.delete(0, tk.END)
            self.entry.config(fg=self.text_color)

    def _on_focus_out(self, event):
        if not self.entry.get():
            self.entry.insert(0, self.placeholder_text)
            self.entry.config(fg=self.placeholder_color)

    def get(self):
        val = self.entry.get()
        if val == self.placeholder_text:
            return ""
        return val

    def set(self, text):
        self.entry.delete(0, tk.END)
        self.entry.insert(0, text)
        self.entry.config(fg=self.text_color) # Valid text is always dark

    def config_entry(self, **kwargs):
        self.entry.config(**kwargs)


class App(TkinterDnD.Tk if HAS_DND else tk.Tk):
    """
    Descrição: Classe principal da aplicação que gerencia a janela e a navegação entre páginas.
    Description: Main application class managing the window and page navigation.
    """
    def __init__(self):
        super().__init__()
        self.title("Conferência Gueddai - Launcher")
        
        # Modern Dimensions & Center Window
        w, h = 800, 600
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.resizable(False, False)
        
        # Styling Constants
        self.BG_COLOR = "#F5F6FA" # Light Neutral
        self.SIDEBAR_COLOR = "#2D3436"
        self.TEXT_COLOR = "#2D3436"
        self.ACCENT_COLOR = "#0984e3"
        self.BTN_COLOR = "#dfe6e9" 

        self.configure(bg=self.BG_COLOR)
        
        # Main Layout: Sidebar + Content
        self.sidebar = tk.Frame(self, bg=self.SIDEBAR_COLOR, width=200)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        
        self.content_area = tk.Frame(self, bg=self.BG_COLOR)
        self.content_area.pack(side="right", fill="both", expand=True)
        
        # Navigation Buttons
        self.add_nav_button("🏠 Início", "LauncherPage")
        self.add_nav_button("🎥 Galeria de Vídeos", "VideoGalleryPage")
        
        # Version
        lbl_ver = tk.Label(self.sidebar, text="v1.4", bg=self.SIDEBAR_COLOR, fg="#636e72", font=("Segoe UI", 8))
        lbl_ver.pack(side="bottom", pady=10)

        # Pages
        self.frames = {}
        for F in (LauncherPage, VideoGalleryPage):
            page_name = F.__name__
            frame = F(parent=self.content_area, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        self.show_frame("LauncherPage")
        
    def add_nav_button(self, text, page_name):
        """
        Descrição: Adiciona um botão de navegação à barra lateral.
        Description: Adds a navigation button to the sidebar.
        """
        btn = tk.Button(
            self.sidebar,
            text=text,
            font=("Segoe UI", 11),
            bg=self.SIDEBAR_COLOR,
            fg="white",
            bd=0,
            activebackground="#636e72",
            activeforeground="white",
            cursor="hand2",
            anchor="w",
            padx=20,
            command=lambda: self.show_frame(page_name)
        )
        btn.pack(fill="x", pady=5)

    def show_frame(self, page_name):
        """
        Descrição: Exibe o frame da página solicitada e executa sua função on_show se existir.
        Description: Displays the requested page frame and executes its on_show function if it exists.
        """
        frame = self.frames[page_name]
        frame.tkraise()
        if hasattr(frame, "on_show"):
            frame.on_show()

class LauncherPage(tk.Frame):
    """
    Descrição: Página inicial para seleção do arquivo Excel e início da conferência.
    Description: Home page for selecting the Excel file and starting the conference.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg="#F5F6FA")
        self.controller = controller
        self.full_file_path = None
        
        self.BG_COLOR = "#F5F6FA"
        self.TEXT_COLOR = "#2D3436"
        self.ACCENT_COLOR = "#0984e3"
        self.SUCCESS_COLOR = "#00b894"
        self.BTN_COLOR = "#dfe6e9"
        
        self.setup_ui()
        self.setup_events()
        
    def clear_focus(self, event):
        self.focus_set()
        
    def setup_ui(self):
        """
        Descrição: Configura os elementos visuais da página (botões, labels, inputs).
        Description: Configures the page's visual elements (buttons, labels, inputs).
        """
        # Header
        lbl_title = tk.Label(self, text="Nova Conferência", font=("Segoe UI", 24, "bold"), bg=self.BG_COLOR, fg=self.TEXT_COLOR)
        lbl_title.pack(anchor="w", padx=40, pady=(40, 5))
        
        lbl_subtitle = tk.Label(self, text="Importe a planilha de pedidos para iniciar", font=("Segoe UI", 11), bg=self.BG_COLOR, fg="#636e72")
        lbl_subtitle.pack(anchor="w", padx=40, pady=(0, 40))
        
        # File Input
        input_frame = tk.Frame(self, bg=self.BG_COLOR)
        input_frame.pack(fill=tk.X, padx=40)
        
        self.rounded_entry = RoundedEntry(input_frame, width=400, height=50, corner_radius=15, bg=self.BG_COLOR, placeholder_text="Selecione o arquivo Excel...")
        self.rounded_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.btn_browse = tk.Button(input_frame, text="📂", font=("Segoe UI", 14), bg=self.BG_COLOR, fg=self.ACCENT_COLOR, bd=0, cursor="hand2", command=self.browse_file)
        self.btn_browse.pack(side=tk.RIGHT, padx=(10, 0))
        
        extra_text = "Arraste o arquivo aqui ou cole (Ctrl+V)" if HAS_DND else "Selecione o arquivo manualmente"
        lbl_hint = tk.Label(self, text=extra_text, font=("Segoe UI", 9), bg=self.BG_COLOR, fg="#b2bec3")
        lbl_hint.pack(anchor="w", padx=40, pady=(5, 30))
        
        # Start Button
        self.btn_start = tk.Button(
            self, 
            text="INICIAR SISTEMA", 
            font=("Segoe UI", 12, "bold"),
            bg=self.BTN_COLOR, 
            fg="#b2bec3", 
            state=tk.DISABLED,
            command=self.start_system,
            height=2,
            bd=0,
            relief=tk.FLAT,
            cursor="arrow"
        )
        self.btn_start.pack(fill=tk.X, padx=40, pady=20)
        
    def setup_events(self):
        """
        Descrição: Configura eventos como Drag-and-Drop e Colar da área de transferência.
        Description: Configures events like Drag-and-Drop and Paste from clipboard.
        """
        self.bind("<Button-1>", self.clear_focus)
        # Bind paste to parent controller root, but check if this frame is visible? 
        # Simpler to bind to root and handle checks, or just update if valid path found.
        self.controller.bind('<Control-v>', self.paste_from_clipboard)
        
        if HAS_DND:
            # We must bind drop to the controller root
            self.controller.drop_target_register(DND_FILES)
            self.controller.dnd_bind('<<Drop>>', self.drop_event)

    def update_file_selection(self, filepath):
        """
        Descrição: Atualiza a interface com o arquivo selecionado e valida se é um Excel.
        Description: Updates the interface with the selected file and validates if it is an Excel file.
        """
        if not filepath: return
        filepath = filepath.strip().strip('"').strip("'")
        
        if os.path.exists(filepath) and filepath.lower().endswith('.xlsx'):
            self.full_file_path = filepath
            filename = os.path.basename(filepath)
            self.rounded_entry.set(filename)
            self.rounded_entry.entry.config(fg="#2d3436")
            
            self.btn_start.config(state=tk.NORMAL, bg=self.SUCCESS_COLOR, fg="white", cursor="hand2")
        else:
            messagebox.showwarning("Arquivo Inválido", "Por favor selecione um arquivo Excel (.xlsx) válido.")

    def browse_file(self):
        filename = filedialog.askopenfilename(title="Selecione a Planilha", filetypes=[("Excel", "*.xlsx")])
        if filename: self.update_file_selection(filename)
            
    def drop_event(self, event):
        data = event.data
        if data.startswith('{') and data.endswith('}'): data = data[1:-1]
        self.update_file_selection(data)
        
    def paste_from_clipboard(self, event=None):
        try:
            data = self.controller.clipboard_get()
            if data and '.xlsx' in data.lower():
                self.update_file_selection(data.split('\n')[0])
        except: pass

    def start_system(self):
        """
        Descrição: Valida as entradas e inicia o processo de escaneamento.
        Description: Validates inputs and starts the scanning process.
        """
        if not self.full_file_path or not os.path.exists(self.full_file_path):
            messagebox.showerror("Erro", "Arquivo inválido!")
            return
            
        self.controller.withdraw() # Hide Launcher
        try:
            loader = DataLoader(self.full_file_path)
            # Default paths
            v_path = "videos_auditoria"
            r_path = "."
            
            scanner = BarcodeScanner(loader, video_path=v_path, report_path=r_path)
            exit_code = scanner.run()
            
            # On return (q pressed or button clicked)
            self.controller.deiconify()
            
            if exit_code == "gallery":
                self.controller.show_frame("VideoGalleryPage")
            else:
                self.controller.show_frame("LauncherPage")
            
        except Exception as e:
            self.controller.deiconify()
            messagebox.showerror("Erro Fatal", f"Ocorreu um erro:\n{e}")

class VideoGalleryPage(tk.Frame):
    """
    Descrição: Página de galeria para visualizar e buscar vídeos gravados.
    Description: Gallery page to view and search recorded videos.
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg="#F5F6FA")
        self.controller = controller
        self.video_dir = "videos_auditoria"
        
        self.setup_ui()
        
    def setup_ui(self):
        self.bind("<Button-1>", self.clear_focus)
        # Header
        header_frame = tk.Frame(self, bg="#F5F6FA")
        header_frame.pack(fill="x", padx=40, pady=(40, 20))
        
        lbl_title = tk.Label(header_frame, text="Galeria de Vídeos", font=("Segoe UI", 24, "bold"), bg="#F5F6FA", fg="#2D3436")
        lbl_title.pack(side="left")
        
        # Search Bar
        search_frame = tk.Frame(self, bg="#F5F6FA")
        search_frame.pack(fill="x", padx=40, pady=(0, 20))
        
        self.search_entry = RoundedEntry(search_frame, width=400, height=40, corner_radius=15, bg="#F5F6FA", placeholder_text="Buscar por NF...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.search_entry.entry.bind("<KeyRelease>", self.filter_videos)
        
        btn_refresh = tk.Button(search_frame, text="🔄 Atualizar", font=("Segoe UI", 10), command=self.load_videos, bg="white", bd=0, cursor="hand2")
        btn_refresh.pack(side="right")
        
        # Video List (Treeview)
        list_frame = tk.Frame(self, bg="white")
        list_frame.pack(fill="both", expand=True, padx=40, pady=(0, 40))
        
        columns = ("nf", "filename", "date")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.tree.heading("nf", text="NF")
        self.tree.heading("filename", text="Arquivo")
        self.tree.heading("date", text="Data")
        
        self.tree.column("nf", width=100)
        self.tree.column("filename", width=300)
        self.tree.column("date", width=150)
        
        self.tree.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.bind("<Double-1>", self.on_double_click)
        
    def on_show(self):
        self.load_videos()
        
    def load_videos(self):
        """
        Descrição: Carrega a lista de vídeos da pasta e exibe na tabela.
        Description: Loads the list of videos from the folder and displays them in the table.
        """
        # Clear
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not os.path.exists(self.video_dir):
            os.makedirs(self.video_dir)
            
        self.all_videos = []
        
        for f in os.listdir(self.video_dir):
            if f.lower().endswith(('.mp4', '.avi')):
                # Extract NF from filename if possible "NF1234.mp4"
                nf = "Desconhecido"
                if f.startswith("NF"):
                    # Try to extract numbers
                    try:
                        base = os.path.splitext(f)[0]
                        nf = base.replace("NF", "").split("_")[0]
                    except: pass
                
                # Date
                full_path = os.path.join(self.video_dir, f)
                mod_time = datetime.datetime.fromtimestamp(os.path.getmtime(full_path)).strftime('%Y-%m-%d %H:%M')
                
                self.all_videos.append((nf, f, mod_time))
                self.tree.insert("", "end", values=(nf, f, mod_time))
                
    def clear_focus(self, event):
        self.focus_set()

    def filter_videos(self, event):
        """
        Descrição: Filtra a lista de vídeos com base no texto digitado (busca por NF).
        Description: Filters the video list based on typed text (search by NF).
        """
        query = self.search_entry.get().lower()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for vid in self.all_videos:
            nf, f, date = vid
            if query in nf.lower() or query in f.lower():
                self.tree.insert("", "end", values=vid)

    def on_double_click(self, event):
        """
        Descrição: Abre o arquivo de vídeo selecionado no player padrão do sistema.
        Description: Opens the selected video file in the system's default player.
        """
        item = self.tree.selection()
        if not item: return
        
        vals = self.tree.item(item[0])["values"]
        filename = vals[1]
        filepath = os.path.join(self.video_dir, filename)
        
        if os.path.exists(filepath):
            os.startfile(filepath)

if __name__ == "__main__":
    from tkinter import ttk # Import ttk here
    app = App()
    app.mainloop()