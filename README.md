# Bingo Card Generator

This Python script generates Bingo cards and saves them to an Excel file. The Bingo cards are created with random numbers within the range of 1 to 25. The generated cards are formatted and saved to an Excel file named `bingo_cards.xlsx`. The script also provides options to either generate new cards or keep existing ones.

## Prerequisites

Make sure you have the following libraries installed before running the script:

- `pandas`
- `openpyxl`

You can install them using the following command:

```bash
pip install pandas openpyxl
```

## How to Use

1. Run the script using a Python interpreter.
2. Choose whether to generate new cards or keep existing ones.
3. Enter the number of cards you want to generate.
4. The script will generate and save Bingo cards to the `bingo_cards.xlsx` file.
5. The Excel file will be formatted with proper styling and saved in landscape mode.

## File Structure

- `bingo_cards.xlsx`: Excel file containing the generated Bingo cards.

## Note

- The script automatically formats the Excel file with appropriate font sizes, alignments, and borders.
- Each Bingo card is represented as a separate row in the Excel file.

Feel free to customize the script according to your needs and preferences. If you encounter any issues or have suggestions for improvement, please feel free to create an issue or contribute to the development.
