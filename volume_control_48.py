import tkinter as tk
import threading
import keyboard
from comtypes import CoInitialize
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


class VolumeControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("System Volume Control")
        self.root.geometry("350x200")  # Adjust the window size here

        # Create labels and entry boxes for low and high volumes
        self.low_label = tk.Label(root, text="Low Volume (0-100):")
        self.low_label.grid(row=0, column=0, padx=10, pady=5)
        self.low_entry = tk.Entry(root)
        self.low_entry.grid(row=0, column=1, padx=10, pady=5)

        self.high_label = tk.Label(root, text="High Volume (0-100):")
        self.high_label.grid(row=1, column=0, padx=10, pady=5)
        self.high_entry = tk.Entry(root)
        self.high_entry.grid(row=1, column=1, padx=10, pady=5)

        self.key_label = tk.Label(root, text="Key to toggle volume:")
        self.key_label.grid(row=2, column=0, padx=10, pady=5)
        self.key_entry = tk.Entry(root)
        self.key_entry.grid(row=2, column=1, padx=10, pady=5)

        # Create button to submit volumes
        self.submit_button = tk.Button(root, text="Enter", command=self.submit_volumes)
        self.submit_button.grid(row=3, column=0, columnspan=2, padx=10, pady=5)

        # Create start and stop buttons
        self.start_button = tk.Button(root, text="Start", command=self.start_toggle)
        self.start_button.grid(row=4, column=0, padx=10, pady=5)
        self.stop_button = tk.Button(root, text="Stop", command=self.stop_toggle)
        self.stop_button.grid(row=4, column=1, padx=10, pady=5)

        # Initialize variables
        self.low_volume = 20  # Set default low volume to 20
        self.high_volume = 50  # Set default high volume to 50
        self.current_volume = self.low_volume
        self.toggle_key = "space"
        self.is_running = False

    def submit_volumes(self):
        # Retrieve and set volumes from entry boxes
        try:
            low_volume = int(self.low_entry.get())
            high_volume = int(self.high_entry.get())
            self.low_volume = max(
                0, min(low_volume, 100)
            )  # Clamp values between 0 and 100
            self.high_volume = max(
                0, min(high_volume, 100)
            )  # Clamp values between 0 and 100
            self.toggle_key = self.key_entry.get()
            print("Volumes and toggle key updated successfully.")
        except ValueError:
            print("Invalid volume settings. Please enter values between 0 and 100.")

    def start_toggle(self):
        # Start toggling volumes when start button is clicked
        if not self.is_running:
            self.is_running = True
            self.volume_toggle_thread = threading.Thread(
                target=self.toggle_volume_thread
            )
            self.volume_toggle_thread.daemon = True
            self.volume_toggle_thread.start()
            print("Volume toggling started.")

    def stop_toggle(self):
        # Stop toggling volumes when stop button is clicked
        if self.is_running:
            self.is_running = False
            print("Volume toggling stopped.")

    def toggle_volume_thread(self):
        # Initialize COM objects
        CoInitialize()

        # Get default audio endpoint
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)

        if not interface:
            print("Failed to find audio interface.")
            return

        # Cast the interface to IAudioEndpointVolume
        volume_interface = interface.QueryInterface(IAudioEndpointVolume)

        if not volume_interface:
            print("Failed to cast to IAudioEndpointVolume interface.")
            return

        # Toggle between low and high volume in a loop
        while self.is_running:
            if keyboard.is_pressed(self.toggle_key):
                if self.current_volume == self.low_volume:
                    volume_interface.SetMasterVolumeLevelScalar(
                        self.high_volume / 100, None
                    )  # Set volume to high volume
                    self.current_volume = self.high_volume
                    print(f"Volume set to {self.high_volume}%")
                else:
                    volume_interface.SetMasterVolumeLevelScalar(
                        self.low_volume / 100, None
                    )  # Set volume to low volume
                    self.current_volume = self.low_volume
                    print(f"Volume set to {self.low_volume}%")
            # Avoid busy-waiting and allow other threads to execute
            threading.Event().wait(0.1)


if __name__ == "__main__":
    root = tk.Tk()
    app = VolumeControlApp(root)
    root.mainloop()
