import winreg as reg
import msvcrt  # For detecting any key press
import os


class ContextmenuTool:
    """
    A class for managing Windows registry operations, including creating keys and updating values.
    """

    BASE_PATH = r"Directory\Background\shell"  # Base path for creating keys

    def __init__(self):
        self.key_name = ""  # User-provided registry key name
        self.command_path = ""  # User-provided file path to execute

    def set_terminal_color(self):
        """
        Sets the terminal color to cyan (turquoise).
        """
        os.system("color 3")  # '3' sets the terminal color to cyan (turquoise)

    def check_admin_rights(self):
        """
        Checks if the program is running with administrator privileges.
        Raises:
            PermissionError: If the program lacks admin rights.
        """
        try:
            with reg.OpenKey(reg.HKEY_CLASSES_ROOT, self.BASE_PATH, 0, reg.KEY_READ):
                pass
        except PermissionError:
            raise PermissionError("This program requires administrator privileges to run.")

    def validate_input(self, key_name, command_path):
        """
        Validates the user-provided key name and file path.
        Args:
            key_name (str): The name of the registry key.
            command_path (str): The path of the file to execute.
        Raises:
            ValueError: If the key name contains invalid characters or the file path is invalid.
        """
        invalid_chars = r'\/:*?"<>|'
        if any(char in key_name for char in invalid_chars):
            raise ValueError(f"The key name contains invalid characters: {invalid_chars}")

        if not os.path.exists(command_path.strip('"')):
            raise ValueError(f"The file path does not exist: {command_path}")

    def get_user_input(self):
        """
        Collects the registry key name and file path from the user.
        """
        # Get the registry key name
        self.key_name = input("Enter the name of the registry key to create: ").strip()
        if not self.key_name:
            raise ValueError("The key name cannot be empty.")

        # Get the file path
        self.command_path = input("Enter the file path to execute: ").strip()
        if not self.command_path:
            raise ValueError("The file path cannot be empty.")

        # Add quotes if not present
        if not (self.command_path.startswith('"') and self.command_path.endswith('"')):
            self.command_path = f'"{self.command_path}"'

        # Validate the inputs
        self.validate_input(self.key_name, self.command_path)

    def create_registry_key(self):
        """
        Creates the registry key and updates its `(Default)` value.
        """
        reg_path = rf"{self.BASE_PATH}\{self.key_name}\command"

        # Create the registry key
        with reg.CreateKey(reg.HKEY_CLASSES_ROOT, rf"{self.BASE_PATH}\{self.key_name}") as key:
            print(f"'{self.key_name}' key was successfully created.")

        # Create the 'command' subkey and set its default value
        with reg.CreateKey(reg.HKEY_CLASSES_ROOT, reg_path) as command_key:
            reg.SetValue(command_key, "", reg.REG_SZ, self.command_path)
            print(f"The '(Default)' value of the 'command' subkey was set to '{self.command_path}'.")

    def run(self):
        """
        Runs the tool: validates inputs, creates keys, and updates values.
        """
        try:
            self.set_terminal_color()  # Set terminal color to cyan
            self.check_admin_rights()  # Ensure admin privileges
            self.get_user_input()  # Collect user inputs
            self.create_registry_key()  # Create the registry key
            print("All operations were successfully completed.")
        except PermissionError as e:
            print(f"Permission error: {e}")
        except ValueError as e:
            print(f"Input error: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    tool = ContextmenuTool()
    tool.run()

    # Wait for any key press before exiting
    print("Press any key to exit...")
    msvcrt.getch()
