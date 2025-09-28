# E-commerce Inventory Analysis for a Vinyl Record Reseller

A **Python** and **Excel** project that automates the process of evaluating a vinyl record inventory for an online reseller. This tool scrapes data from the Discogs marketplace to provide actionable intelligence for pricing, inventory management, and sales strategy.

---

## The Business Problem

A small online record seller needs an automated way to assess the current market value of their inventory. Manually researching each record is time-consuming and inefficient. This project solves that problem by creating a script that programmatically fetches marketplace data to identify high-value items and prioritize which records to list for sale.

---

## Results & Value Proposition

This automated tool provides a clear return on investment by translating raw collection data into a strategic asset.

* **Drastically Reduced Research Time:** The script reduces manual research time by over 95%, allowing the user to focus on sales rather than data entry.
* **Data-Driven Pricing:** By pulling the lowest current market price, it provides a baseline value for every item, allowing the seller to price their inventory competitively and confidently.
* **Identified High-Priority Items:** In a test collection, the analysis flagged dozens of records with a potential profit margin of over 300%, identifying them as high-priority items to list immediately.

---

## Live Demo & Output

The final output is a structured Excel table with dynamic search and filtering capabilities.

- **Screenshot of Excel Table:**
 

- **Demo GIF of Filtering & Sorting:**
  ![Demo GIF](demo.gif)

---

## Technical Features

- Scrapes a user's Discogs collection using the Discogs API
- Retrieves marketplace statistics for each release
- Calculates the lowest resale price based on current listings
- Saves data as a structured Excel table
- Autosaves results periodically to avoid data loss during long runs
- Includes an advanced formula for dynamic search and filtering in Excel
- Optional VBA for a user-friendly placeholder in the search cell

---

## Setup & Usage

1.  **Obtain Discogs API key:** https://www.discogs.com/settings/developers

2.  **Edit the config section** of `discogs_pull.py`:
    ```python
    # -- config ----
    API_TOKEN = "YOUR_API_KEY" # enter your Discogs API key
    USERNAME = "YOUR_USERNAME" # enter your Discogs username
    OUTPUT = "collection_roi_table.xlsx" # change output name
    REQ_SLEEP = 1.1 # No more requests than 60/min
    PER_PAGE = 100
    AUTOSAVE_INTERVAL = 10  # autosave after every N releases processed
    # --------------
    ```
3.  **Run the Python script** to fetch collection data:
    ```bash
    python discogs_pull.py
    ```
    The output will be saved as `collection_roi_table.xlsx`.

---

## Advanced Implementation Details

### Dynamic Excel Search Formula

This formula enables dynamic filtering and sorting of the collection table. The results (in cell `L4`) spill downward automatically based on the search term in cell `J2`.

* **Formula (for cell `L4`):**
    ```excel
    =IF(OR($J$2="", $J$2="Search"), "", IFERROR(SORT(FILTER(CollectionTable[[release_id]:[lowest_price]], (ISNUMBER(SEARCH($J$2, CollectionTable[artist]))) + (ISNUMBER(SEARCH($J$2, CollectionTable[title]))) + (ISNUMBER(SEARCH($J$2, TEXT(CollectionTable[release_id], "0"))))), 4, -1), "Not found"))
    ```

* **Readability Version:**
    ```excel
    =IF(OR($J$2="", $J$2="Search"), "", 
        IFERROR(
            SORT(   
                FILTER(
                    CollectionTable[[release_id]:[lowest_price]], 
                    (ISNUMBER(SEARCH($J$2, CollectionTable[artist]))) 
                    + (ISNUMBER(SEARCH($J$2, CollectionTable[title]))) 
                    + (ISNUMBER(SEARCH($J$2, TEXT(CollectionTable[release_id], "0"))))),
            4, -1), 
        "Not found")
    )
    ```

### VBA Placeholder Code (Optional)

This VBA script provides a user-friendly placeholder in the Excel search box, which clears when clicked.

```VBA
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
```
## Acknowledgements

Assistance with coding, Excel formulas, and Python scripting was obtained from AI tools including:

- **ChatGPT (GPT-5-mini)**
- **Gemini (2.5 Flash)**
- **Claude (Sonnet 4)** — provided the most accurate and complete guidance
 

All code was actively developed and refined with human oversight — issues were identified and corrected during the process.