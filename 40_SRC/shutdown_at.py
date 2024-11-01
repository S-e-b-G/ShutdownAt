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
import sys                                      # To close the program
import time                                     # For time checking and delays
import os                                       # For file operations and shutdown execution
from datetime import datetime, timedelta        # To manage time calculations

# THIRD-PARTY libraries
    # GUI Components:
from tkinter import Tk, Entry, Label, Button, Frame, END, Toplevel, BooleanVar, Checkbutton
import tkinter.messagebox as mb                 # Message boxes for user warnings
import subprocess                               # For executing shutdown command
import winsound                                 # To make a sound



##################
# GLOBAL CONSTANTS
##################
# Path to the shutdown shortcut
SHUTDOWN_BAT_PATH = "D:/_Cloud/MEGA/EtudierCAD/_PartageSeb/Info-Tech/Appli/Scripts/PC-SGPro/Store apps and Shutdown.bat"  # Constant => pylint: disable=C0103, C0301
# Hibernation command
HIBERNATE_COMMAND   = "rundll32.exe powrprof.dll,SetSuspendState 0,1,0" # Constant => pylint: disable=C0103

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

    #############
    # ATTRIBUTES
    #############
    # Default delay before shutdown
    SHUTDOWN_DEF_DELAY  = 80

    TMR_WIN_BACKGD_COLOR            = 'black'
    TMR_WIN_BACKGD_COLOR_TIME_1     = '#512010'
    TMR_WIN_BACKGD_COLOR_TIME_2     = '#442E17'
    TMR_WIN_BACKGD_COLOR_TIME_3     = '#3F3813'
    TMR_WIN_TEXT_TMR_COLOR          = 'white'
    TMR_WIN_TEXT_MSG_COLOR          = '#D8D8D8'
    TMR_WIN_BACKGD_COLOR_INF_2_MIN  = '#512010'
    TMR_WIN_BACKGD_COLOR_INF_5_MIN  = '#442E17'
    TMR_WIN_BACKGD_COLOR_INF_10_MIN = '#3F3813'

    TMR_WIN_TIME_1                  = 2
    TMR_WIN_TIME_2                  = 5
    TMR_WIN_TIME_3                  = 10



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
        self.root.title("Extinction/Hibernation")
        self.root.geometry(f"+{int((1920 - WIN_WIDTH) / 2)}+{int((1080 - WIN_HEIGHT) / 2)}")
        self.root.attributes('-topmost', True)
        self.root.focus_force()


        # Create a label to instruct the user
        Label(self.root, text="Extinction/Hibernation à (HH.MM) :").pack()


        # Create the entry widget for user input
        self.time_entry = Entry(self.root, justify="center")
        self.time_entry.pack()


        # Set default time if not provided
        if not i_delay_min:
            i_delay_min = 10  # Default delay is 10 minutes
        #endif
        l_text = (datetime.now() + timedelta(minutes=i_delay_min)).strftime("%H.%M")
        self.time_entry.insert(0, l_text)
        self.time_entry.focus_set()
        self.time_entry.select_range(0, END)


        self.shutdown_option = BooleanVar(value=True)  # Default is checked for shutdown
        self.option_checkbox = Checkbutton(
            self.root,
            text="Hibernation (vide) /\n Extinction (coché)",
            variable=self.shutdown_option
        )
        self.option_checkbox.pack()


        # Create buttons for "OK" and "Cancel"
        button_frame = Frame(self.root)
        button_frame.pack()

        self.start_button = Button(button_frame, text="OK", command=self.start_shutdown)
        self.start_button.pack(side="left")

        cancel_button = Button(button_frame, text="Annuler", command=self.root.destroy)
        cancel_button.pack(side="right")


        # Timer related
        self.timer_window           = None
        self.timer_is_topmost       = False
        self._timer_drag_start_x    = 0
        self._timer_drag_start_y    = 0


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

    # MAIN APP/WINDOW RELATED
    #########################

    def show_initial_window(self, event):
        """
        Show the initial window configuration.
        """
        if not self.root or not self.root.winfo_exists():
            self.root = Tk()
            ShutdownApp(self.SHUTDOWN_DEF_DELAY)
        else:
            if self.root.winfo_exists():
                self.root.deiconify()  # Show the initial window
                self.root.attributes("-topmost", True)  # Bring it to the front
            else:
                self.root = Tk()
                ShutdownApp(self.SHUTDOWN_DEF_DELAY)

        event.widget.master.destroy()  # Destroy the timer window

        return
    # end function



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

            # Create the timer window
            self.create_timer_window("Extinction dans")

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



    # TIMER WINDOW RELATED
    ######################

    def create_timer_window(self, message):
        """
        Create a new window to display the timer.

        Args:
            time_difference (timedelta): The time difference between target and current time.
            message (str): The message to display when the timer reaches zero.
        """
        # Create window
        self.timer_window = Toplevel(self.root)
        self.timer_window.title("Minuteur")
        self.timer_window.config(bg=self.TMR_WIN_BACKGD_COLOR)
        # Set top most
        self.timer_is_topmost = True
        self.timer_window.attributes("-topmost", self.timer_is_topmost)
        self.timer_window.overrideredirect(True) # Do not display title bar

        # Bind events
        self.timer_bind_shortcuts(self.timer_window)

        # Create labels to display the timer
        message_label = Label(self.timer_window, text=message, font=("Calibri", 10, "italic"))
        message_label.pack(padx=5)
        message_label.config(bg=self.TMR_WIN_BACKGD_COLOR)
        message_label.config(fg=self.TMR_WIN_TEXT_MSG_COLOR)

        timer_label = Label(self.timer_window, text="00:00", font=("DSEG7 Classic", 10))
        timer_label.pack(padx=5)
        timer_label.config(bg=self.TMR_WIN_BACKGD_COLOR)
        timer_label.config(fg=self.TMR_WIN_TEXT_TMR_COLOR)

        # Make the window draggable
        self._timer_drag_start_x = 0
        self._timer_drag_start_y = 0
        self.timer_make_draggable()

        self.timer_countdown(self.timer_window,
                             message_label,
                             timer_label,
                             message)

        return
    # end function



    def timer_make_draggable(self):
        """
        Binds the button so that the window can be draggable
        """
        self.timer_window.bind("<Button-1>", self.timer_start_move)
        self.timer_window.bind("<B1-Motion>", self.timer_do_move)

        return


    def timer_start_move(self, event):
        """
        Starts the window to be moved.
        """
        self._timer_drag_start_x = event.x
        self._timer_drag_start_y = event.y

        return


    def timer_do_move(self, event):
        """
        Continuously moves the window.
        """
        x = self.timer_window.winfo_x() - self._timer_drag_start_x + event.x
        y = self.timer_window.winfo_y() - self._timer_drag_start_y + event.y
        self.timer_window.geometry(f"+{x}+{y}")

        return



    def timer_countdown(self, window, message_label, timer_label, message):
        """
        Perform countdown and display timer.

        Args:
            time_difference (timedelta): The time difference between target and current time.
            message (str): The message to display when the timer reaches zero.
            timer_label (Label): The label to display the timer.
        """
        timer_time = datetime.strptime(self.shutdown_time, "%H.%M").time()
        time_difference = timedelta(hours=timer_time.hour, minutes=timer_time.minute) - \
            timedelta(hours=datetime.now().time().hour, minutes=datetime.now().time().minute)

        minutes, seconds = divmod(time_difference.seconds, 60)
        hours, minutes = divmod(minutes, 60)

        if not (hours == minutes == seconds == 0):
            minutes, seconds = divmod(time_difference.seconds, 60)
            hours, minutes = divmod(minutes, 60)

            if hours == 0 and minutes <= self.TMR_WIN_TIME_1:
                l_color = self.TMR_WIN_BACKGD_COLOR_TIME_1
                window.config(bg=l_color)
                message_label.config(bg=l_color)
                timer_label.config(bg=l_color)
            elif hours == 0 and minutes <= self.TMR_WIN_TIME_2:
                l_color = self.TMR_WIN_BACKGD_COLOR_TIME_2
                window.config(bg=l_color)
                message_label.config(bg=l_color)
                timer_label.config(bg=l_color)
            elif hours == 0 and minutes <= self.TMR_WIN_TIME_3:
                l_color = self.TMR_WIN_BACKGD_COLOR_TIME_3
                window.config(bg=l_color)
                message_label.config(bg=l_color)
                timer_label.config(bg=l_color)
            # else : no color change

            timer_label.config(text=f'{hours:02d}:{(minutes):02d}')

            timer_label.after(1000, self.timer_countdown, window,
                              message_label, timer_label, message)
        # about to shutdown

        return



    def custom_messagebox(self, parent, title, message):
        """
        Display.
        """
        parent.destroy()

        dialog = Toplevel()
        dialog.title(title)
        dialog.attributes('-topmost', 'true')
        dialog.config(bg = 'red')

        label = Label(dialog, text=str(self.time_entry.get()), font=("Calibri", 10, "italic"))
        label.pack(padx=180, pady=0)
        label.config(bg = 'red', fg='#D0D0D0')

        label = Label(dialog, text=message, font=("Calibri", 16, "bold"))
        label.pack(padx=10, pady=3)
        label.config(bg = 'red', fg='#E8E8E8')

        ok_button = Button(dialog, text='OK', command=sys.exit)
        ok_button.pack(pady=10)

        dialog.focus_force()  # Set focus to this window
        dialog.grab_set()     # Grab focus

        # Play a sound
        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)

        return
    # End function



    def toggle_topmost_timer(self, event, i_win): # pylint: disable=W0613 # "Unused argument 'event'"
        """
        Toggle topmost status of the window.
        """
        self.timer_is_topmost = not self.timer_is_topmost
        i_win.attributes("-topmost", self.timer_is_topmost)

        if self.timer_is_topmost:
            i_win.overrideredirect(True)
        else:
            i_win.overrideredirect(False)

        return



    def timer_bind_shortcuts(self, l_win: Toplevel):
        """
        Bind keyboard shortcuts and events (timer window).
        """
        l_win.bind('t', lambda event, win=l_win: self.toggle_topmost_timer(event, win))

        return
    # end function



    ###########################
    # PUBLIC FUNCTIONS
    ###########################

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

        if self.shutdown_option.get():  # Checked: shutdown
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
        else: # Unchecked: hibernate
            # subprocess.run(["shutdown", "/h"], check=True)
            os.system(HIBERNATE_COMMAND)
        # endif


        return
    # end function

# End of file
