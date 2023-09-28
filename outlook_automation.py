import time
import pyautogui

# Function to open Outlook
def open_outlook():
    # Replace with the path to your Outlook shortcut or executable
    outlook_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
    pyautogui.hotkey("win", "r")
    pyautogui.write(outlook_path)
    pyautogui.press("enter")
    time.sleep(5)  # Wait for Outlook to open

# Function to send an email
def send_email(to, subject, message):
    # Open a new email composition window
    pyautogui.hotkey("ctrl", "n")
    time.sleep(2)  # Wait for the new email window to open

    # Fill in recipient, subject, and message
    pyautogui.write(to)
    pyautogui.press("tab")
    pyautogui.write(subject)
    pyautogui.press("tab")
    pyautogui.write(message)

    # Send the email
    pyautogui.hotkey("ctrl", "enter")

# Function to respond to an email with a customizable message
def respond_to_email(sender, response_subject, response_message):
    # Search for the sender's email in the inbox
    pyautogui.hotkey("ctrl", "e")
    time.sleep(2)
    pyautogui.write(sender)
    time.sleep(2)
    pyautogui.press("enter")

    # Reply to the email
    pyautogui.hotkey("ctrl", "r")
    time.sleep(2)
    pyautogui.write(response_subject)
    pyautogui.press("tab")
    pyautogui.write(response_message)

    # Send the response
    pyautogui.hotkey("ctrl", "enter")

if __name__ == "__main__":
    open_outlook()

    # Example usage with customized responses
    send_email("recipient@example.com", "Hello", "This is an automated email.")
    
    sender = "sender@example.com"
    response_subject = "Re: Your Email"
    response_message = (
        "Thank you for your email. "
        "This is an automated response with a customized message."
    )
    
    respond_to_email(sender, response_subject, response_message)
