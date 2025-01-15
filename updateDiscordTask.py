import os
import subprocess
from xml.etree import ElementTree as ET
import logging
from datetime import datetime

# Configuration
base_path = r"C:\Users\Paul\AppData\Local\Discord"  # Base directory where app folders are located
task_name = "Discord-Admin"  # Replace with the actual Task Scheduler task name

# Namespace for Task Scheduler XML
NAMESPACE = {"task": "http://schemas.microsoft.com/windows/2004/02/mit/task"}

# Set up logging
script_dir = os.path.dirname(os.path.abspath(__file__))  # Directory of the Python script
logs_dir = os.path.join(script_dir, "logs")
os.makedirs(logs_dir, exist_ok=True)  # Create logs folder if it doesn't exist

log_file = os.path.join(logs_dir, f"update_task_{datetime.now().strftime('%Y-%m-%d')}.log")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def find_latest_app_path(base_path):
    # Find the latest Discord app folder.
    folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f)) and f.startswith("app-1.")]
    if not folders:
        logging.error("No 'app-1.X.XXXX' folder found!")
        raise FileNotFoundError("No 'app-1.X.XXXX' folder found!")
    # Sort folders by version number (latest first)
    folders.sort(reverse=True)
    latest_path = os.path.join(base_path, folders[0])
    logging.info(f"Latest app path found: {latest_path}")
    return latest_path

def get_task_xml(task_name):
    # Export the task's XML configuration.
    result = subprocess.run(["schtasks", "/query", "/tn", task_name, "/xml"], capture_output=True, text=True)
    if result.returncode != 0:
        logging.error(f"Failed to query task: {result.stderr}")
        raise RuntimeError(f"Failed to query task: {result.stderr}")
    return result.stdout

def get_current_task_path(task_name):
    # Retrieve the current executable path from the task's XML.
    xml_content = get_task_xml(task_name)
    root = ET.fromstring(xml_content)
    # Find the Command element within the Actions element
    command_element = root.find(".//task:Command", NAMESPACE)
    if command_element is None:
        logging.error("Command element not found in task XML!")
        raise RuntimeError("Command element not found in task XML!")
    return command_element.text

def update_task_xml(task_name, new_path):
    # Update the task's executable path.
    xml_content = get_task_xml(task_name)
    root = ET.fromstring(xml_content)
    
    # Update the Command element with the new path
    command_element = root.find(".//task:Command", NAMESPACE)
    if command_element is None:
        logging.error("Command element not found in task XML!")
        raise RuntimeError("Command element not found in task XML!")
    command_element.text = os.path.join(new_path, "Discord.exe")
    
    # Save updated XML to a temp file
    temp_xml_path = "temp_task.xml"
    with open(temp_xml_path, "w") as temp_file:
        temp_file.write(ET.tostring(root, encoding="unicode"))
    
    # Update the existing task with the modified XML
    subprocess.run(["schtasks", "/create", "/tn", task_name, "/xml", temp_xml_path, "/f"], check=True)
    logging.info("Task updated successfully.")
    
    # Cleanup
    os.remove(temp_xml_path)

def main():
    try:
        latest_path = find_latest_app_path(base_path)
        current_task_path = get_current_task_path(task_name)
        new_executable_path = os.path.join(latest_path, "Discord.exe")
        
        print(f"Current task path: {current_task_path}")
        print(f"New executable path: {new_executable_path}")
        
        if current_task_path == new_executable_path:
            logging.info("No changes detected. Task update not needed.")
            return
        
        logging.info("Path has changed. Updating the task...")
        update_task_xml(task_name, latest_path)
    except Exception as e:
        logging.error(f"Error: {e}")
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
