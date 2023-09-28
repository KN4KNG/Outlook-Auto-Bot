
# OutlookAutoBot

OutlookAutoBot is a Python script that automates various tasks in Microsoft Outlook, such as sending emails, responding to emails, and other common email management actions. This script is designed to save time and increase productivity by reducing repetitive tasks in Outlook.

## Features

-   **Open Outlook**: Automatically open the Microsoft Outlook application.
-   **Send Email**: Compose and send emails with custom recipients, subjects, and messages.
-   **Respond to Email**: Quickly respond to incoming emails with customizable responses.
-   **Customizable Responses**: Easily tailor response messages to suit your needs.

## Requirements

Before running OutlookAutoBot, make sure you have the following:

-   Python installed on your system.
-   The `pyautogui` library installed. You can install it using `pip install pyautogui`.

## Usage

1.  Clone this repository to your local machine:
    
    `git clone https://github.com/KN4KNG/Outlook-Auto-Bot.git` 
    
2.  Navigate to the project directory:
    
    `cd OutlookAutoBot` 
    
3.  Open the `outlook_automation.py` file and customize it according to your needs. You may need to adjust the path to your Outlook executable and configure your email templates.
    
4.  Run the script:
    
    `python outlook_automation.py` 
    
## Customization

### Outlook Path

In the `outlook_automation.py` file, replace the `outlook_path` variable with the correct path to your Outlook executable.

### Email Templates

You can customize email templates by modifying the `send_email` and `respond_to_email` functions in the script. Add your desired email subject and message.

## Contributing

If you'd like to contribute to OutlookAutoBot or have ideas for additional features, please feel free to submit issues or pull requests. Your contributions are welcome!

## License

This project is licensed under the MIT License - see the LICENSE.
