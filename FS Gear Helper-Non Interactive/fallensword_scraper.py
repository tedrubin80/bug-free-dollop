#!/usr/bin/env python3
import requests
from bs4 import BeautifulSoup
import csv
import json
import time
from typing import List, Dict
import sys

def scrape_items_page(page_num: int = 0) -> List[Dict]:
    """Scrape a single page of items from Fallen Sword guide"""
    url = f"https://guide.fallensword.com/index.php?cmd=items&index={page_num}"

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')

        # Find the main items table
        items_data = []

        # Look for table rows containing item data
        tables = soup.find_all('table')

        for table in tables:
            rows = table.find_all('tr')

            for row in rows:
                cells = row.find_all(['td', 'th'])
                if len(cells) > 0:
                    # Extract text from each cell
                    row_data = [cell.get_text(strip=True) for cell in cells]

                    # Filter out empty rows and header rows
                    if row_data and any(row_data) and not all(x in ['Name', 'Level', 'Type', 'Rarity', 'Attack', 'Defense', 'Armor', 'Damage', 'HP', 'Stamina'] for x in row_data if x):
                        # Try to parse as item row
                        if len(row_data) >= 2:
                            items_data.append(row_data)

        return items_data

    except requests.RequestException as e:
        print(f"Error fetching page {page_num}: {e}")
        return []

def scrape_all_items(max_pages: int = 61) -> List[List]:
    """Scrape all pages of items"""
    all_items = []

    for page in range(max_pages):
        print(f"Scraping page {page + 1}/{max_pages}...")

        items = scrape_items_page(page * 25)  # Pages seem to increment by 25
        all_items.extend(items)

        # Be respectful with rate limiting
        time.sleep(1)

    return all_items

def save_to_csv(items: List[List], filename: str = "fallensword_items.csv"):
    """Save items to CSV file"""
    with open(filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)

        # Write header
        writer.writerow(['Column1', 'Column2', 'Column3', 'Column4', 'Column5',
                        'Column6', 'Column7', 'Column8', 'Column9', 'Column10'])

        # Write data
        for item in items:
            # Pad row to ensure consistent columns
            padded_row = item + [''] * (10 - len(item))
            writer.writerow(padded_row[:10])

    print(f"Saved {len(items)} items to {filename}")

def create_html_table(items: List[List], filename: str = "fallensword_items.html"):
    """Create an HTML file with sortable/searchable table"""
    html = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Fallen Sword Items Database</title>
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/jquery.dataTables.min.css">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            color: #333;
            text-align: center;
        }
        #itemsTable {
            background-color: white;
        }
        #itemsTable_wrapper {
            margin-top: 20px;
        }
    </style>
</head>
<body>
    <h1>Fallen Sword Items Database</h1>
    <table id="itemsTable" class="display" style="width:100%">
        <thead>
            <tr>
"""

    # Add header columns
    for i in range(10):
        html += f"                <th>Column {i+1}</th>\n"

    html += """            </tr>
        </thead>
        <tbody>
"""

    # Add data rows
    for item in items:
        html += "            <tr>\n"
        padded_row = item + [''] * (10 - len(item))
        for cell in padded_row[:10]:
            # Escape HTML characters
            cell_text = str(cell).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            html += f"                <td>{cell_text}</td>\n"
        html += "            </tr>\n"

    html += """        </tbody>
    </table>

    <script>
        $(document).ready(function() {
            $('#itemsTable').DataTable({
                "pageLength": 25,
                "lengthMenu": [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
                "scrollX": true,
                "order": [[0, "asc"]]
            });
        });
    </script>
</body>
</html>"""

    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"Created HTML file: {filename}")

def main():
    print("Fallen Sword Items Scraper")
    print("=" * 40)

    # Test with first page
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        print("Testing with first page only...")
        items = scrape_items_page(0)
        print(f"Found {len(items)} items on first page")
        if items:
            print("Sample items:")
            for item in items[:5]:
                print(f"  {item}")
        return

    # Scrape all pages
    print("Starting full scrape of all 61 pages...")
    all_items = scrape_all_items(61)

    print(f"\nTotal items scraped: {len(all_items)}")

    # Save to CSV
    save_to_csv(all_items, "fallensword_items.csv")

    # Create HTML file
    create_html_table(all_items, "fallensword_items.html")

    print("\nScraping complete!")
    print("Files created:")
    print("  - fallensword_items.csv (for Google Sheets import)")
    print("  - fallensword_items.html (searchable/sortable web page)")

if __name__ == "__main__":
    main()