import customtkinter as ctk
import cv2
import os
import csv
import psutil
from datetime import datetime
from PIL import Image
from tkinter import filedialog
from ultralytics import YOLO
import gc

# 1. Set the professional theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# 2. Setup Directories and Logging (FIXED: Using Absolute Paths)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DIR = os.path.join(BASE_DIR, "raw_data")
RESULTS_DIR = os.path.join(BASE_DIR, "results")
LOG_FILE = os.path.join(BASE_DIR, "detection_log.csv")

for folder in [RAW_DIR, RESULTS_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder)

if not os.path.exists(LOG_FILE):
    with open(LOG_FILE, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Timestamp", "File Name", "Grade A", "Grade B", "Grade C", "Inference Time (ms)", "CPU Load (%)", "RAM Usage (GB)"])

class TunaGraderApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("🐟 Advanced Tuna Quality Inspector")
        self.geometry("1100x700")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        print("Loading YOLO model...")
        self.model = YOLO("best.pt") # Update path if needed
        self.cap = None
        self.is_running = False
        self.current_frame = None 
        
        # --- FIXED: Safe defaults to prevent silent startup crashes ---
        self.current_annotated_frame = None
        self.current_counts = {0: 0, 1: 0, 2: 0}
        self.current_inf_time = 0.0
        self.current_cpu = 0.0
        self.current_ram = 0.0

        self.setup_gui()
        self.after(100, self.update_telemetry)

    def update_telemetry(self):
        # 1. Measure the current CPU and App RAM
        self.current_cpu = psutil.cpu_percent()
        current_process = psutil.Process(os.getpid())
        self.current_ram = round(current_process.memory_info().rss / (1024 ** 3), 2)
        
        # 2. FIXED: Safety Check so it doesn't crash before UI loads
        if hasattr(self, 'lbl_cpu') and hasattr(self, 'lbl_ram'):
            self.lbl_cpu.configure(text=f"CPU Load: {self.current_cpu}%")
            self.lbl_ram.configure(text=f"App RAM: {self.current_ram} GB")
        
        # 3. Tell the app to run this function again in 1000 milliseconds (1 second)
        self.after(1000, self.update_telemetry)

    def setup_gui(self):
        # --- Left Side: Tabs for Modes ---
        self.tabview = ctk.CTkTabview(self, width=750, height=650)
        self.tabview.pack(side="left", padx=20, pady=20, fill="both", expand=True)

        self.tab_camera = self.tabview.add("Live Camera")
        self.tab_image = self.tabview.add("Image Mode")

        # Camera Tab Setup
        self.video_label = ctk.CTkLabel(self.tab_camera, text="Camera Offline", font=("Arial", 24))
        self.video_label.pack(expand=True, pady=10)
        
        self.cam_btn_frame = ctk.CTkFrame(self.tab_camera, fg_color="transparent")
        self.cam_btn_frame.pack(fill="x", pady=10)
        
        self.btn_start = ctk.CTkButton(self.cam_btn_frame, text="▶ Start Camera", fg_color="green", hover_color="darkgreen", command=self.start_camera)
        self.btn_start.pack(side="left", padx=10, expand=True)
        self.btn_stop = ctk.CTkButton(self.cam_btn_frame, text="⏹ Stop Camera", fg_color="red", hover_color="darkred", state="disabled", command=self.stop_camera)
        self.btn_stop.pack(side="left", padx=10, expand=True)
        self.btn_capture = ctk.CTkButton(self.cam_btn_frame, text="📸 Capture & Log", state="disabled", command=self.capture_and_log)
        self.btn_capture.pack(side="left", padx=10, expand=True)

        # Image Tab Setup
        self.image_label = ctk.CTkLabel(self.tab_image, text="No Image Selected", font=("Arial", 24))
        self.image_label.pack(expand=True, pady=10)
        
        self.img_btn_frame = ctk.CTkFrame(self.tab_image, fg_color="transparent")
        self.img_btn_frame.pack(fill="x", pady=10)
        
        self.btn_upload = ctk.CTkButton(self.img_btn_frame, text="📂 Select Image", command=self.upload_image)
        self.btn_upload.pack(side="left", padx=10, expand=True)
        self.btn_process = ctk.CTkButton(self.img_btn_frame, text="⚡ Process & Log", state="disabled", command=self.process_static_image)
        self.btn_process.pack(side="left", padx=10, expand=True)

        # --- Right Side: Controls & Telemetry ---
        self.control_frame = ctk.CTkFrame(self, width=250)
        self.control_frame.pack(side="right", padx=20, pady=20, fill="y")

        ctk.CTkLabel(self.control_frame, text="⚙️ Configuration", font=("Arial", 18, "bold")).pack(pady=10)
        ctk.CTkLabel(self.control_frame, text="Confidence Threshold").pack(pady=(10, 0))
        self.conf_slider = ctk.CTkSlider(self.control_frame, from_=0.1, to=1.0, number_of_steps=90)
        self.conf_slider.set(0.5)
        self.conf_slider.pack(pady=5, padx=20)

        ctk.CTkLabel(self.control_frame, text="📊 Current Detection", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.lbl_a = ctk.CTkLabel(self.control_frame, text="Grade A: 0", font=("Arial", 16))
        self.lbl_a.pack(pady=2)
        self.lbl_b = ctk.CTkLabel(self.control_frame, text="Grade B: 0", font=("Arial", 16))
        self.lbl_b.pack(pady=2)
        self.lbl_c = ctk.CTkLabel(self.control_frame, text="Grade C: 0", font=("Arial", 16))
        self.lbl_c.pack(pady=2)

        ctk.CTkLabel(self.control_frame, text="💻 System Telemetry", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.lbl_inf = ctk.CTkLabel(self.control_frame, text="Inference: 0.0 ms", font=("Arial", 14))
        self.lbl_inf.pack(pady=2)
        self.lbl_cpu = ctk.CTkLabel(self.control_frame, text="CPU Load: 0%", font=("Arial", 14))
        self.lbl_cpu.pack(pady=2)
        self.lbl_ram = ctk.CTkLabel(self.control_frame, text="App RAM: 0.0 GB", font=("Arial", 14))
        self.lbl_ram.pack(pady=2)

    # --- Live Camera Logic ---
    def start_camera(self):
        if not self.is_running:
            self.cap = cv2.VideoCapture(0)
            self.is_running = True
            self.btn_start.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            self.btn_capture.configure(state="normal")
            self.video_label.configure(text="")
            self.update_frame()

    def stop_camera(self):
        if self.is_running:
            self.is_running = False
            if self.cap:
                self.cap.release()
            self.btn_start.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            self.btn_capture.configure(state="disabled")
            self.video_label.configure(image="", text="Camera Offline")
            
            # --- FORCE MEMORY CLEANUP ---
            self.current_frame = None
            if hasattr(self, 'current_annotated_frame'):
                self.current_annotated_frame = None
                
            gc.collect() 
            self.reset_stats()

    def update_frame(self):
        if self.is_running and self.cap.isOpened() and self.tabview.get() == "Live Camera":
            ret, frame = self.cap.read()
            if ret:
                self.current_frame = frame.copy() # Save raw frame for capture
                self.run_inference_and_update_ui(frame, self.video_label)
            self.after(15, self.update_frame)
        elif self.is_running and self.tabview.get() != "Live Camera":
            self.after(100, self.update_frame)

    def capture_and_log(self):
        if self.current_frame is not None:
            self.save_and_log_data(self.current_frame, "camera_capture")

    # --- Static Image Logic ---
    def upload_image(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.jpg *.jpeg *.png")])
        if file_path:
            self.static_img_path = file_path
            self.static_img_cv = cv2.imread(file_path)
            self.run_inference_and_update_ui(self.static_img_cv, self.image_label)
            self.btn_process.configure(state="normal")

    def process_static_image(self):
        if hasattr(self, 'static_img_cv'):
            self.save_and_log_data(self.static_img_cv, "static_upload")

    # --- Core Processing & Logging ---
    def run_inference_and_update_ui(self, frame_bgr, target_label):
        conf_thresh = self.conf_slider.get()
        results = self.model.predict(frame_bgr, conf=conf_thresh, verbose=False)[0]

        # Process counts
        self.current_counts = {0: 0, 1: 0, 2: 0}
        for box in results.boxes:
            self.current_counts[int(box.cls[0])] += 1
        
        # Telemetry (Only Inference time belongs here now)
        self.current_inf_time = round(results.speed['inference'], 1)

        # Update UI Labels
        self.lbl_a.configure(text=f"Grade A: {self.current_counts[0]}")
        self.lbl_b.configure(text=f"Grade B: {self.current_counts[1]}")
        self.lbl_c.configure(text=f"Grade C: {self.current_counts[2]}")
        self.lbl_inf.configure(text=f"Inference: {self.current_inf_time} ms")

        # Render Image
        self.current_annotated_frame = results.plot()
        color_frame = cv2.cvtColor(self.current_annotated_frame, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(color_frame).resize((700, 500))
        ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(700, 500))
        target_label.configure(image=ctk_img, text="")
        target_label.image = ctk_img

    def save_and_log_data(self, raw_frame, prefix):
        # FIXED: Try/Except block with Absolute Paths to prevent silent crashing
        try:
            if self.current_annotated_frame is None:
                print("⚠️ Warning: Model hasn't processed an image yet. Wait a second!")
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{prefix}_{timestamp}.jpg"

            # Save Images
            success_raw = cv2.imwrite(os.path.join(RAW_DIR, filename), raw_frame)
            success_res = cv2.imwrite(os.path.join(RESULTS_DIR, filename), self.current_annotated_frame)

            if not success_raw or not success_res:
                print("❌ ERROR: OpenCV failed to save the image. Check folder permissions.")
                return

            # Log to CSV
            with open(LOG_FILE, mode='a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    filename,
                    self.current_counts.get(0, 0), 
                    self.current_counts.get(1, 0), 
                    self.current_counts.get(2, 0), 
                    self.current_inf_time,
                    self.current_cpu,
                    self.current_ram
                ])
            print(f"✅ SUCCESS: Saved '{filename}' and updated CSV!")

        except Exception as e:
            print(f"❌ CRITICAL ERROR saving data: {e}")

    def reset_stats(self):
        for lbl in [self.lbl_a, self.lbl_b, self.lbl_c]: 
            lbl.configure(text=lbl.cget("text").split(":")[0] + ": 0")
        self.lbl_inf.configure(text="Inference: 0.0 ms")

    def on_closing(self):
        self.stop_camera()
        self.destroy()

if __name__ == "__main__":
    app = TunaGraderApp()
    app.mainloop()