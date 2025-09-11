readme_content = """
# Discogs-ROI

A Python script that scrapes your personal Discogs collection and fetches marketplace stats to calculate the lowest resale price for each release. Data is saved as an Excel table with structured formatting for easy filtering and analysis. Ideal for collectors tracking potential ROI on vinyl records.

## Features

- Scrapes your Discogs collection using the Discogs API
- Retrieves marketplace statistics for each release
- Calculates the lowest resale price
- Saves data as a structured Excel table
- Includes a search function for filtering by artist, title, or release ID
- Placeholder input handled via VBA in Excel

## Installation

1. Clone the repository:

git clone https://github.com/soupbadger/discogs-roi.git
cd discogs-roi

2. Install dependencies:

pip install -r requirements.txt

3. Obtain a Discogs API token: https://www.discogs.com/developers/#page:authentication

4. Create a `.env` file in the project root and add your API token:

DISCOGS_API_TOKEN=your_api_token_here

## Usage

Run the Python script to fetch collection data and save it as an Excel file:

python discogs_roi.py

The output will be saved as `collection_roi_table.xlsx`.

## Excel Search Formula

Filter and sort your collection data dynamically in Excel:

=IF(OR($J$2="", $J$2="Search"), "", IFERROR(SORT(FILTER(CollectionTable[[release_id]:[lowest_price]], (ISNUMBER(SEARCH($J$2, CollectionTable[artist]))) + (ISNUMBER(SEARCH($J$2, CollectionTable[title]))) + (ISNUMBER(SEARCH($J$2, TEXT(CollectionTable[release_id], "0"))))), 4, -1), "Not found"))

## VBA Placeholder Code

Handles clearing and restoring the search placeholder in Excel:

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim inputCell As Range
    Set inputCell = Me.Range("J2") ' change to your input cell

    ' Clear placeholder when clicked
    If Not Intersect(Target, inputCell) Is Nothing Then
        If inputCell.Value = "Search" Then
            inputCell.Value = ""
            inputCell.Font.Italic = False
            inputCell.Font.Color = vbBlack
        End If
    Else
        ' Restore placeholder if cell left empty
        If inputCell.Value = "" Then
            inputCell.Value = "Search"
            inputCell.Font.Italic = True
            inputCell.Font.Color = RGB(150, 150, 150)
        End If
    End If
End Sub

## Learning Experience

This project was a learning exercise where I tested ChatGPT, Gemini, and Claude to see which AI provided the best results for coding and documentation.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgements

- Discogs API: https://www.discogs.com/developers/
- openpyxl: https://openpyxl.readthedocs.io/en/stable/
- requests: https://docs.python-requests.org/en/latest/
"""

with open("README.md", "w", encoding="utf-8") as f:
    f.write(readme_content)

print("README.md created successfully.")
