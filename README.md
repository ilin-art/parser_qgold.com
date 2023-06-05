# Jewelry Data Parser

This script is a data parser for jewelry items from web pages and saves them to Excel files. It utilizes the `requests` module for making HTTP requests and the `openpyxl` module for working with Excel files.

## Installation

To use the script, you need to install the dependencies specified in the `requirements.txt` file. You can install them using `pip` by running the following command:
pip install -r requirements.txt

## Usage

The script performs data parsing from two different jewelry item URLs, each executed in a separate thread. The item data is then saved to Excel files. The script can be customized to set the interval between parsing using the `--interval` command-line argument.

Example usage of the script with a 900-second interval (default):
python your_script.py

Example usage of the script with a 600-second interval:
python your_script.py --interval 600

## Files

- `your_script.py`: The main script file containing the core code for parsing data and saving it to Excel files.
- `requirements.txt`: The file containing the list of dependencies to be installed using `pip`.

## Dependencies

The script uses the following dependencies:

- `requests`: A module for making HTTP requests.
- `openpyxl`: A module for working with Excel files.

You can install the dependencies by running the following command:
pip install -r requirements.txt

## Contributing

If you find any issues or want to contribute improvements, please create an issue or pull request in this repository.

## License

This project is licensed under the terms of the [MIT License](LICENSE).
