# MailConverter

MailConverter is a tool to convert Outlook `.msg` files to `.mht` files. It processes emails, scales images, and adds print styles to ensure the content fits within a DIN A4 page size.

## Features

- Convert Outlook `.msg` files to `.mht` format.
- Scale images to fit within a DIN A4 page size.
- Add print styles for A4 size.
- Replace German tags with English tags in the `.mht` files.
- Open `.msg` and `.mht` files on different monitors.

## Requirements

- Python 3.8 or higher
- Windows OS (due to dependencies on `pywin32` and `pywinauto`)

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/your-username/mailconverter.git
    cd mailconverter
    ```

2. Install dependencies using Poetry:
    ```sh
    poetry install
    ```

## Usage

### Command Line Interface

To convert `.msg` files in a specified directory, run:
```sh
poetry run main <directory>
```
Replace `<directory>` with the path to the directory containing `.msg` files.

### Debugging with VSCode

1. Open the project in VSCode.
2. Open the Command Palette (`Ctrl+Shift+P`), type `Debug: Select and Start Debugging`, and choose the desired configuration:
    - **Python: MailConverter** - Runs the `mailconverter.py` script with the `data` directory as an argument.
    - **Python: Test MailConverter** - Runs the `test_main.py` script for testing purposes.

## Project Structure

```
MailConverter/
├── .vscode/
│   └── launch.json
├── mailconverter/
│   └── mailconverter.py
├── tests/
│   └── test_main.py
├── .gitignore
├── pyproject.toml
└── README.md
```

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
````