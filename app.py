import os
import subprocess
import webbrowser
from flask import Flask, render_template, request, redirect, url_for
import psutil

# Flask app
app = Flask(__name__)

# Global variable to store the chosen IP address
chosen_ip = None

# Functions for the requested actions
def connect_to_winbox(address):
    """Attempt to connect to Winbox with the provided address."""
    try:
        winbox_path = r"C:\Users\Yoni\Desktop\winbox64.exe"  # Path to Winbox
        subprocess.Popen([winbox_path, address, "---", "---"])
        print(f"Trying to connect to Winbox with address: {address}...")
    except Exception as e:
        print(f"Failed to connect to Winbox with address {address}. Error: {e}")

def update_vpn_address(vpn_name, address):
    """Update the VPN address in the rasphone.pbk file."""
    phonebook_path = r"C:\Users\Yoni\AppData\Roaming\Microsoft\Network\Connections\Pbk\rasphone.pbk"
    if not os.path.exists(phonebook_path):
        print(f"Phonebook file not found at {phonebook_path}")
        return
    try:
        with open(phonebook_path, 'r') as file:
            lines = file.readlines()

        with open(phonebook_path, 'w') as file:
            in_section = False
            for line in lines:
                if line.strip().startswith(f"[{vpn_name}]"):
                    in_section = True
                elif in_section and line.strip().startswith("[") and not line.strip().startswith(f"[{vpn_name}]"):
                    in_section = False
                if in_section and line.strip().startswith("PhoneNumber="):
                    print(f"Updating PhoneNumber to {address}")
                    line = f"PhoneNumber={address}\n"
                file.write(line)

        print(f"VPN address updated successfully to {address} for '{vpn_name}'.")
    except Exception as e:
        print(f"Failed to update VPN address. Error: {e}")

def connect_to_vpn(vpn_name, username, password, domain):
    """Initiate a VPN connection using rasdial with credentials."""
    try:
        print(f"Connecting to VPN '{vpn_name}' with domain '{domain}'...")
        subprocess.run(["rasdial", vpn_name, username, password, "/domain:" + domain], check=True)
        print("VPN connection successfully initiated.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to connect to VPN '{vpn_name}'. Error: {e}")

def open_browser():
    """Open incognito browsers to access managed devices."""
    urls = ["http://------", "http://-------"]
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    for url in urls:
        print(f"Opening incognito browser to access {url}...")
        try:
            subprocess.Popen([chrome_path, "--incognito", url])
        except FileNotFoundError:
            print(f"Chrome not found. Opening default browser for {url}.")
            webbrowser.open(url)


def close_components():
    """Close all opened components (Winbox, VPN, incognito browsers)."""
    try:
        # Disconnect VPN
        print("Disconnecting VPN...")
        vpn_disconnect_result = subprocess.run(["rasdial", "/disconnect"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if "No connections" in vpn_disconnect_result.stdout or "No connections" in vpn_disconnect_result.stderr:
            print("No active VPN connection to disconnect.")
        else:
            print("VPN disconnected.")

        # Kill Winbox
        print("Closing Winbox...")
        winbox_result = subprocess.run(["taskkill", "/IM", "winbox64.exe", "/F"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if "SUCCESS" in winbox_result.stdout.upper():
            print("Winbox successfully closed.")
        else:
            print("Winbox process not found or already closed.")

        # Close only incognito Chrome windows using WMIC
        print("Closing incognito browser windows...")
        cmd = 'wmic process where "name=\'chrome.exe\' and CommandLine like \'%--incognito%\'" get ProcessId'
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=True)

        if result.stdout:
            process_ids = [line.strip() for line in result.stdout.splitlines() if line.strip().isdigit()]
            for pid in process_ids:
                print(f"Terminating Chrome incognito process (PID: {pid})...")
                subprocess.run(["taskkill", "/PID", pid, "/F"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        print("Incognito Chrome windows closed.")

    except subprocess.CalledProcessError as e:
        print(f"Error closing components. {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")



@app.route('/', methods=['GET', 'POST'])
def home():
    global chosen_ip
    if request.method == 'POST':
        ip_address = request.form.get('ip_address')
        if ip_address:
            chosen_ip = ip_address  # Store chosen IP
            # Step 1: Connect to Winbox
            connect_to_winbox(ip_address)
            
            # Step 2: Update VPN address
            vpn_name = "VPN"  # Replace with your VPN entry name
            update_vpn_address(vpn_name, ip_address)

            # Step 3: Connect to VPN
            username = "----"
            password = "----"
            connect_to_vpn(vpn_name, username, password, ip_address)

            # Step 4: Open browsers
            open_browser()

            return redirect(url_for('success'))  # Redirect to success page
    return render_template('home.html')

@app.route('/success', methods=['GET', 'POST'])
def success():
    global chosen_ip
    if request.method == 'POST':
        action = request.form.get('action')  # Retrieve the action from the form
        if action == "close_components":
            close_components()  # Close all components
            return redirect(url_for('home'))  # Redirect to the home page
    return render_template('success.html', ip=chosen_ip)


if __name__ == '__main__':
    app.run(debug=True)
