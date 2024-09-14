"""
    USE:
        This module defines the ShutdownApp class which handles the shutdown process
        based on a user-defined time input.

    INSTALLATION:
        - Standard modules: time, os, tkinter, datetime, subprocess
"""

##################
# IMPORT SECTION
##################
# STANDARD libraries
import time  # For time checking and delays
import os  # For file operations and shutdown execution
from datetime import datetime, timedelta  # To manage time calculations

# THIRD-PARTY libraries
from tkinter import Tk, Entry, Label, Button, Frame, END  # GUI Components
import tkinter.messagebox as mb  # Message boxes for user warnings
import subprocess  # For executing shutdown command



##################
# GLOBAL CONSTANTS
##################
# Path to the shutdown shortcut
SHUTDOWN_BAT_PATH = "D:/_Cloud/MEGA/EtudierCAD/_PartageSeb/Info-Tech/Appli/Scripts/PC-SGPro/Store apps and Shutdown.bat"  # Constant => pylint: disable=C0103, C0301
# Fallback shutdown command
SHUTDOWN_COMMAND = "rundll32.exe powrprof.dll,SetSuspendState 0,1,0"  # Constant => pylint: disable=C0103

# Window settings
WIN_WIDTH = 100  # Constant => pylint: disable=C0103
WIN_HEIGHT = 100  # Constant => pylint: disable=C0103



##################
# CLASS DEFINITION
##################

class ShutdownApp:
    """
    This class represents the shutdown application, which allows users to input a time
    for the system to shut down. The class handles validation, user interface, and timing.
    """

    ############
    # ATTRIBUTES
    ############

    ###########################
    # INITIALIZATION FUNCTION
    ###########################

    def __init__(self, i_delay_min=None):
        """
        Initializes the ShutdownApp instance, setting up the user interface and initial settings.

        Parameters
        ----------
        i_delay_min : int, optional
            Time delay in minutes for the default shutdown input (default is 10 minutes).

        Returns
        -------
        None
        """

        # Create the root window
        self.root = Tk()
        self.root.title("Extinction")
        self.root.geometry(f"+{int((1920 - WIN_WIDTH) / 2)}+{int((1080 - WIN_HEIGHT) / 2)}")
        self.root.attributes('-topmost', True)
        self.root.focus_force()

        # Create a label to instruct the user
        Label(self.root, text="Extinction à (HH.MM) :").pack()

        # Create the entry widget for user input
        self.time_entry = Entry(self.root, justify="center")
        self.time_entry.pack()

        # Set default time if not provided
        if not i_delay_min:
            i_delay_min = 10  # Default delay is 10 minutes
        l_text = (datetime.now() + timedelta(minutes=i_delay_min)).strftime("%H.%M")
        self.time_entry.insert(0, l_text)
        self.time_entry.focus_set()
        self.time_entry.select_range(0, END)

        # Create buttons for "OK" and "Cancel"
        button_frame = Frame(self.root)
        button_frame.pack()

        self.start_button = Button(button_frame, text="OK", command=self.start_shutdown)
        self.start_button.pack(side="left")

        cancel_button = Button(button_frame, text="Annuler", command=self.root.destroy)
        cancel_button.pack(side="right")

        # Initialize shutdown time attribute and warning flag
        self.shutdown_time = ""
        self.warning_shown = False

        # Start the main loop
        self.root.mainloop()

        return
    # end function



    ###########################
    # PRIVATE FUNCTIONS
    ###########################


    def is_valid_time(self, i_time):
        """
        Validates if the given time is in the correct format (HH.MM) and within valid ranges.

        Parameters
        ----------
        i_time : str
            Time string to validate in the format HH.MM.

        Returns
        -------
        bool
            True if the time is valid, False otherwise.
        """

        # Split the time by the dot separator
        l_time_parts = i_time.split(".")

        # Ensure the input has exactly two parts and both are numeric
        if len(l_time_parts) != 2 or not (l_time_parts[0].isdigit() and l_time_parts[1].isdigit()):
            return False

        # Parse hour and minute
        l_hour = int(l_time_parts[0])
        l_minute = int(l_time_parts[1])

        # Check if hour and minute are within valid ranges
        return 0 <= l_hour <= 23 and 0 <= l_minute <= 59
    # end function



    ###########################
    # PUBLIC FUNCTIONS
    ###########################


    def start_shutdown(self):
        """
        Initiates the shutdown process by validating the input time and starting the timer loop.

        Parameters
        ----------
        None

        Returns
        -------
        None
        """

        # Get the user input
        l_user_input = self.time_entry.get()

        # Validate the time format
        if self.is_valid_time(l_user_input):
            # Set shutdown time if valid
            self.shutdown_time = l_user_input

            # Disable input and buttons
            self.time_entry.config(state="disabled")
            self.start_button.config(state="disabled")

            # Hide the main window
            self.root.withdraw()

            # Start checking the time
            self.check_time()
        else:
            # Show error message if format is invalid
            mb.showerror("Erreur de format",
                         "Format invalide, indiquer l'heure sous la forme HH.MM.")

        return
    # end function


    def check_time(self):
        """
        Periodically checks the current time and compares it to the shutdown time.
        Initiates a shutdown when the time matches or warns the user when two minutes are left.

        Parameters
        ----------
        None

        Returns
        -------
        None
        """

        # Get the current time in HH.MM format
        l_current_time = time.strftime("%H.%M")

        # Check if it's time to shut down
        if l_current_time >= self.shutdown_time:
            self.shutdown()
        elif l_current_time == self.get_two_minutes_before(self.shutdown_time) and \
             not self.warning_shown:
            # Warn the user two minutes before shutdown
            mb.showwarning("Warning", "Extinction dans deux minutes")
            self.warning_shown = True

        # Re-check after 1 second
        self.root.after(1000, self.check_time)

        return
    # end function


    def get_two_minutes_before(self, i_time):
        """
        Returns the time two minutes before the provided time.

        Parameters
        ----------
        i_time : str
            Time string in HH.MM format.

        Returns
        -------
        str
            Time two minutes before the given time in HH.MM format.
        """

        # Split time and calculate two minutes before
        l_time_parts = i_time.split(".")
        l_hour = int(l_time_parts[0])
        l_minute = int(l_time_parts[1]) - 2

        # Adjust hour and minute for overflow
        if l_minute < 0:
            l_minute += 60
            l_hour -= 1
        if l_hour < 0:
            l_hour = 23

        # Return formatted time
        return f"{l_hour:02d}.{l_minute:02d}"
    # end function


    def shutdown(self):
        """
        Executes the system shutdown or falls back to the system command if the shortcut fails.

        Parameters
        ----------
        None

        Returns
        -------
        None
        """

        # Verify if the shortcut exists
        if os.path.exists(SHUTDOWN_BAT_PATH):
            try:
                # Execute the .bat file using subprocess.run()
                result = subprocess.run([SHUTDOWN_BAT_PATH], shell=True, check=True)

                # Optional: Print the result, which includes exit code and output
                print(f"Batch file executed successfully with return code: {result.returncode}")
            except subprocess.CalledProcessError as e:
                print(f"Error executing the .bat file: {e}.\nExecute the shutdown command.")
                # Fallback to system shutdown command
                subprocess.run(["shutdown", "/s", "/t", "0"], check=True)
        else:
            print("Batch file not found.\nExecute the shutdown command.")
            subprocess.run(["shutdown", "/s", "/t", "0"], check=True)

        return
    # end function

# End of file
