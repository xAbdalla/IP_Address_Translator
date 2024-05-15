<img src="res/img/IAT_icon.png" align="right" style="height: 64px" alt="icon"/>

# IP Address Translator (IAT)

![Static Badge](https://img.shields.io/badge/language-python-yellow)
![Static Badge](https://img.shields.io/badge/version-1.0-blue)
![Static Badge](https://img.shields.io/badge/license-MIT-green)
[![Static Badge](https://img.shields.io/badge/author-xAbdalla-red)](https://github.com/xAbdalla)

IP Address Translator is a Python script that maps IP addresses to descriptive object names or domain names. With a GUI interface and a simple design, the script is easy to use by anyone.

<p align="center">
  <img src="res/img/screenshot.png">
</p>

## Features

### Various Input Options
- Accepts Excel / CSV / Text files.
- Direct input of single IP address, subnet, range, or list of them (Comma Separated).

  ####   Files Specifications
    - Input Excel / CSV files could have multiple sheets, each must contain a column named '**Subnet**'.
    - Reference Excel / CSV files require three columns: '**Tenant**', '**Address object**', and '**Subnet**'.
    - Check the provided [examples](res/examples) for the correct format.

### Various Searching Methods
- **Reference File**: Searches for matches in a user-provided reference file.
- **Palo Alto**: Connects via SSH to Panorama to fetch address objects.
- **Fortinet**: Connects via REST API to FortiGate to retrieve address objects.
- **Cisco ACI**: Connects via SSH to APIC to retrieve address objects based on a specified Class.
- **DNS**: Resolves IPs to domain names using the system dns servers or a user-provided DNS servers.

  ####   Palo Alto Panorama/FW Specifications
    - Ensure Panorama/Firewall is reachable and has CLI access.
    - Leave "VSYS" field empty if you want to retrieve address objects from all virtual systems.
    - The program may fail to import the addresses due to slow response, just try again until it works.
  
  ####   Fortinet FortiGate Specifications
    - Ensure FortiGate is reachable and you have REST API enabled.
    - Leave "VDOM" field empty if you want to retrieve address objects from all virtual domains.
    
  ####   Cisco ACI Specifications
    - Ensure APIC is reachable and has CLI access.
    - Specify the Class of the address objects to be searched.
    - The program searches the "dn" attribute exclusively.
    
  ####   DNS Resolver
    - Resolves IPs to domain names using the system dns servers or a user-provided DNS servers.
    - You can specify up to four DNS servers.

### Encrypted Credentials Storage
- An option to save your credentials for future use and avoid re-entering them.
- Credentials are stored locally in the application directory.
- Stored information are encrypted for security purposes.
- The encryption key is unique to each user and machine.

### Saving the Output
- Results are exported to a new Excel (.xlsx) file for ease of access and analysis.
- The generated file contains the original data along with the mapped names.
- The user can specify the output file name and location to prevent overwriting.

### Logging
- Detailed logs are generated for each operation.
- Logs are saved in a separate file for future reference.
- Logs are also displayed in the GUI for immediate feedback.
- Logs are color-coded for better readability.

### Error Handling
- Detailed error messages are displayed in the GUI.
- Logs are generated for each error for future reference.
- Errors are color-coded for better readability.

## Requirements
- Install [Python 3](https://www.python.org/downloads/) (v3.12 recommended).
- Install the required packages using the following command:
  ```commandline
  pip install -r requirements.txt
  ```
   ### Build Executable
    - You can find the prebuild executable files in the [releases page](https://github.com/xAbdalla/IP_Address_Translator/releases).
    - To build the executable file, you need to install the following packages:
    ```commandline
     pip install pyinstaller
     ```
    - Run the following command to build the executable file:
    ```commandline
       pyinstaller --name "IP_Address_Translator_v1.0" --onefile --windowed --icon "res/img/IAT_icon.ico" --splash "res/img/IAT_splash.png" --disable-windowed-traceback "main.py"
     ```
    - If you want to compress the executable file, you can use [UPX](https://upx.github.io/) by adding the following flag:
    ```text
       --upx-dir "path/to/upx/folder" --upx-exclude "python3.dll" --upx-exclude "_uuid.pyd"
     ```
    - The executable file will be generated in the 'dist' folder.

## Usage
- Run the executable file or run the script using the following command:
  ```commandline
    python "main.py"
  ```
- Fill in the required fields and select the desired search method.
- Ensure the chosen search methods are accessible and correctly configured.
- Provide necessary credentials and remember to save them if needed.
- For Cisco ACI, you must specify the correct Class for targeted searches.
- Review the generated Excel file for mapped IPs based on the selected search methods.
- Review the logs for detailed information about the operation.
- For any issues or inquiries, please contact the author.

## Contributing
Feel free to contribute to this project by forking it and submitting a pull request.