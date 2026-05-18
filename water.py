import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkfont
import time
import json
import datetime
import os

# --- Constants ---
INITIAL_MINUTES = 30
INITIAL_SECONDS = 0
WINDOW_WIDTH = 400
WINDOW_HEIGHT = 580
MAIN_BG_COLOR = "#F0F8FF"  # AliceBlue
BUTTON_BLUE_COLOR = "#1E90FF" # DodgerBlue
BUTTON_GREEN_COLOR = "#32CD32" # LimeGreen
BUTTON_TEXT_COLOR = "white"
LABEL_TEXT_COLOR = "#333333" # Dark grey for text
TIMER_FONT = ("Arial", 48, "bold")
TITLE_FONT_FAMILY = "SimHei"
TITLE_FONT_FALLBACK = "Arial"
TITLE_FONT_SIZE = 20
TIPS_FONT_FAMILY = "SimHei"
TIPS_FONT_FALLBACK = "Arial"
NORMAL_FONT_FAMILY = "Arial"
NORMAL_FONT_SIZE = 10
TIPS_TEXT_FONT_SIZE = 9
BUTTON_FONT_SIZE = 11

# 修改数据文件路径到用户目录
DATA_FILE_NAME = os.path.join(os.path.expanduser("~"), "water_reminder_data.json")

class WaterReminderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("喝水提醒") # Drink Water Reminder
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.configure(bg=MAIN_BG_COLOR)
        self.root.resizable(False, False)

        # --- State Variables ---
        self.time_remaining = INITIAL_MINUTES * 60 + INITIAL_SECONDS
        self.timer_running = False
        self.water_count = 0  # This will be updated from loaded data
        self.timer_id = None

        # --- Data Handling ---
        self.app_data = self._load_data()
        # Stores the date string for which self.water_count is currently valid
        self.current_date_str_for_data = "" 
        self._ensure_current_day_data() # Initialize/load today's water count

        # --- Font Setup ---
        try:
            tkfont.Font(family=TITLE_FONT_FAMILY, size=TITLE_FONT_SIZE) # Test font
            self.title_font_spec = (TITLE_FONT_FAMILY, TITLE_FONT_SIZE, "bold")
        except tk.TclError:
            self.title_font_spec = (TITLE_FONT_FALLBACK, TITLE_FONT_SIZE, "bold")

        try:
            tkfont.Font(family=TIPS_FONT_FAMILY, size=TIPS_TEXT_FONT_SIZE) # Test font
            self.tips_font_spec = (TIPS_FONT_FAMILY, TIPS_TEXT_FONT_SIZE)
            self.tips_title_font_spec = (TIPS_FONT_FAMILY, TIPS_TEXT_FONT_SIZE + 2, "bold")
        except tk.TclError:
            self.tips_font_spec = (TIPS_FONT_FALLBACK, TIPS_TEXT_FONT_SIZE)
            self.tips_title_font_spec = (TIPS_FONT_FALLBACK, TIPS_TEXT_FONT_SIZE + 2, "bold")

        self.normal_font_spec = (NORMAL_FONT_FAMILY, NORMAL_FONT_SIZE)
        self.button_font_spec = (NORMAL_FONT_FAMILY, BUTTON_FONT_SIZE, "bold")

        # --- UI Elements ---
        self.create_widgets()
        self.update_timer_display()
        self.update_water_count_display() # Display initial count

        # --- Handle window close ---
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _load_data(self):
        """Loads data from the JSON file."""
        if os.path.exists(DATA_FILE_NAME):
            try:
                with open(DATA_FILE_NAME, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # Basic validation for expected structure
                    if "log" not in data or not isinstance(data["log"], dict):
                        print(f"Warning: {DATA_FILE_NAME} has unexpected structure. Re-initializing.")
                        return {"log": {}}
                    return data
            except json.JSONDecodeError:
                print(f"Error: Could not decode JSON from {DATA_FILE_NAME}. Starting with empty data.")
                return {"log": {}} # Return default structure if file is corrupt
            except Exception as e:
                print(f"Error loading data: {e}. Starting with empty data.")
                return {"log": {}}
        return {"log": {}} # Return default if file doesn't exist

    def _save_data(self):
        """Saves the current app_data to the JSON file."""
        try:
            with open(DATA_FILE_NAME, 'w', encoding='utf-8') as f:
                json.dump(self.app_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error saving data: {e}")
            messagebox.showerror("保存错误", f"无法保存数据到 {DATA_FILE_NAME}:\n{e}")


    def _ensure_current_day_data(self):
        """
        Checks if the date has changed. If so, updates self.current_date_str_for_data,
        and resets self.water_count for the new day (or loads it if an entry already exists).
        Ensures the current day has an entry in self.app_data['log'].
        """
        today_iso_date = datetime.date.today().isoformat()

        if today_iso_date != self.current_date_str_for_data:
            # Date has changed or it's the first run setting this up.
            # The count for the previous self.current_date_str_for_data (if any) is already in self.app_data["log"]
            # and would have been saved by a previous _save_data() call.
            
            self.current_date_str_for_data = today_iso_date

            if self.current_date_str_for_data not in self.app_data["log"]:
                # New day, or first time running on this day, and no entry exists yet.
                self.app_data["log"][self.current_date_str_for_data] = 0
                # No immediate save needed here, will be saved when water count actually changes or on close.
            
            # Load the water count for the (new) current day.
            self.water_count = self.app_data["log"][self.current_date_str_for_data]
            
            # It's important to update the display if the day changed and water_count was reset/loaded.
            if hasattr(self, 'water_count_label'): # Check if UI is initialized
                 self.update_water_count_display()


    def create_widgets(self):
        # Main Title
        title_label = tk.Label(self.root, text="健康喝水提醒", font=self.title_font_spec, bg=MAIN_BG_COLOR, fg=BUTTON_BLUE_COLOR)
        title_label.pack(pady=20)

        # Water Drop Icon
        icon_label = tk.Label(self.root, text="💧", font=("Arial", 60), bg=MAIN_BG_COLOR, fg=BUTTON_BLUE_COLOR)
        icon_label.pack()

        # Timer Display
        self.timer_display_label = tk.Label(self.root, text="30:00", font=TIMER_FONT, bg=MAIN_BG_COLOR, fg=LABEL_TEXT_COLOR)
        self.timer_display_label.pack(pady=10)

        # Separator
        separator_frame = tk.Frame(self.root, height=5, width=WINDOW_WIDTH - 80, bg=BUTTON_BLUE_COLOR)
        separator_frame.pack(pady=5)

        # Water Count Display
        self.water_count_label = tk.Label(self.root, text=f"今日已喝水: {self.water_count} 次", font=self.normal_font_spec, bg=MAIN_BG_COLOR, fg=LABEL_TEXT_COLOR)
        self.water_count_label.pack(pady=10)

        # Health Tips Frame
        tips_frame = tk.Frame(self.root, bg="#E0FFFF", relief=tk.SOLID, borderwidth=1)
        tips_frame.pack(pady=15, padx=20, fill="x")

        tips_title_label = tk.Label(tips_frame, text="💡 健康小贴士", font=self.tips_title_font_spec, bg="#E0FFFF", fg=BUTTON_BLUE_COLOR)
        tips_title_label.pack(anchor="w", padx=10, pady=(5,2))

        tips_text = [
            "• 每30分钟喝一杯水 (200ml)",
            "• 每天建议喝水2000ml",
            "• 早晨起床后先喝一杯水",
            "• 饭前半小时喝水有助消化"
        ]
        for tip in tips_text:
            tip_label = tk.Label(tips_frame, text=tip, font=self.tips_font_spec, bg="#E0FFFF", fg=LABEL_TEXT_COLOR, justify=tk.LEFT)
            tip_label.pack(anchor="w", padx=10, pady=1)

        # Buttons Frame
        button_frame = tk.Frame(self.root, bg=MAIN_BG_COLOR)
        button_frame.pack(pady=20, fill="x", side=tk.BOTTOM, padx=20, ipady=10)

        button_ipadx = 20
        button_ipady = 8
        button_relief = tk.GROOVE
        button_borderwidth = 2

        self.start_button = tk.Button(button_frame, text="开始提醒", font=self.button_font_spec,
                                      bg=BUTTON_BLUE_COLOR, fg=BUTTON_TEXT_COLOR,
                                      command=self.handle_start_reminder,
                                      relief=button_relief, borderwidth=button_borderwidth,
                                      padx=button_ipadx, pady=button_ipady)
        self.start_button.pack(side=tk.LEFT, expand=True, padx=10)

        self.drank_button = tk.Button(button_frame, text="已喝水", font=self.button_font_spec,
                                     bg="grey", fg=BUTTON_TEXT_COLOR, # Initially grey
                                     command=self.handle_drank_water,
                                     relief=button_relief, borderwidth=button_borderwidth,
                                     padx=button_ipadx, pady=button_ipady)
        self.drank_button.pack(side=tk.RIGHT, expand=True, padx=10)

    def update_timer_display(self):
        mins, secs = divmod(self.time_remaining, 60)
        self.timer_display_label.config(text=f"{mins:02d}:{secs:02d}")

    def update_water_count_display(self):
        self.water_count_label.config(text=f"今日已喝水: {self.water_count} 次")

    def countdown(self):
        if self.timer_running and self.time_remaining > 0:
            self.time_remaining -= 1
            self.update_timer_display()
            self.timer_id = self.root.after(1000, self.countdown)
        elif self.timer_running and self.time_remaining == 0:
            self.timer_running = False
            self.update_timer_display()
            self.shake_window()
            messagebox.showinfo("时间到!", "该喝水啦！点击'已喝水'开始下一次提醒。")
            self.drank_button.config(bg=BUTTON_GREEN_COLOR) # Make "Drank Water" prominent
            self.start_button.config(state=tk.NORMAL) # Allow starting again

    def handle_start_reminder(self):
        self._ensure_current_day_data() # Ensure date context is current before starting timer actions

        if not self.timer_running:
            self.timer_running = True
            if self.time_remaining == 0: # Reset if timer was at 00:00
                self.time_remaining = INITIAL_MINUTES * 60 + INITIAL_SECONDS
            self.update_timer_display()
            self.countdown()
            self.drank_button.config(bg=BUTTON_GREEN_COLOR) # Change "已喝水" to green
            self.start_button.config(state=tk.DISABLED)
        # This elif condition might need review based on exact desired pause/resume logic
        # For now, starting reminder implies a fresh countdown sequence
        # elif self.timer_id: 
        #     self.root.after_cancel(self.timer_id)
        #     self.timer_running = True
        #     self.update_timer_display()
        #     self.countdown()
        #     self.drank_button.config(bg=BUTTON_GREEN_COLOR)
        #     self.start_button.config(state=tk.DISABLED)


    def handle_drank_water(self):
        self._ensure_current_day_data() # Crucial: ensures water_count and date are for today

        self.water_count += 1
        self.app_data["log"][self.current_date_str_for_data] = self.water_count
        self._save_data() # Save data after incrementing
        self.update_water_count_display()

        if self.timer_id:
            self.root.after_cancel(self.timer_id)

        self.time_remaining = INITIAL_MINUTES * 60 + INITIAL_SECONDS
        self.timer_running = True
        self.update_timer_display()
        self.countdown()
        self.drank_button.config(bg=BUTTON_GREEN_COLOR)
        self.start_button.config(state=tk.DISABLED) # Timer is now running, so disable start

    def shake_window(self):
        original_x = self.root.winfo_x()
        original_y = self.root.winfo_y()
        shake_intensity = 5
        shake_duration = 50 # ms
        num_shakes = 6

        for i in range(num_shakes):
            if i % 2 == 0:
                self.root.geometry(f"+{original_x + shake_intensity}+{original_y}")
            else:
                self.root.geometry(f"+{original_x - shake_intensity}+{original_y}")
            self.root.update_idletasks()
            time.sleep(shake_duration / 1000)
        self.root.geometry(f"+{original_x}+{original_y}")

    def on_closing(self):
        """Handles window close event."""
        # Data is saved on each "drank water" action, but an explicit save here can be a fallback.
        # However, if _ensure_current_day_data made changes that weren't followed by a count update,
        # those might not be saved. For robustness, ensure all data is current.
        if self.app_data: # Check if app_data exists
             self._ensure_current_day_data() # Make sure current day's 0 is logged if no drinks today
             if self.current_date_str_for_data and self.current_date_str_for_data not in self.app_data["log"]:
                 # This case should ideally be covered by _ensure_current_day_data
                 self.app_data["log"][self.current_date_str_for_data] = 0
             elif self.current_date_str_for_data and self.app_data["log"][self.current_date_str_for_data] != self.water_count:
                 # Sync if there's a mismatch, though unlikely with current flow
                 self.app_data["log"][self.current_date_str_for_data] = self.water_count

             self._save_data()
        self.root.destroy()

# --- Main ---
if __name__ == "__main__":
    root = tk.Tk()
    app = WaterReminderApp(root)
    root.mainloop()
