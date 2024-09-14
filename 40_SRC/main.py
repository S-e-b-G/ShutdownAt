"""
    USE:
        Main module to initiate the ShutdownApp class.

    INSTALLATION:
        - Requires `shutdown_at.py` for the `ShutdownApp` class.
        - For win32gui and win32con, install via: pip pywin32
"""

##################
# IMPORT SECTION
##################
# APPLICATION modules
from shutdown_at import ShutdownApp  # Import ShutdownApp class


##################
# GLOBAL CONSTANTS
##################
# None


##################
# GLOBAL VARIABLES
##################
# None


##################
# MAIN FUNCTION
##################

def main():
    """
    The main function of the program.
    It creates an instance of the ShutdownApp class to start the shutdown process.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """

    # Create an instance of ShutdownApp with 80 minutes default delay
    ShutdownApp(80)

    return
# end function


# Call the main function if the script is executed directly
if __name__ == "__main__":
    main()
# end if

# End of file
