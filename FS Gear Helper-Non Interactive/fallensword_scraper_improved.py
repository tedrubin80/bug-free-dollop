#!/usr/bin/env python3
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import sys
import re

def scrape_items_page(page_index: int = 0) -> pd.DataFrame:
    """Scrape a single page of items from Fallen Sword guide"""
    url = f"https://guide.fallensword.com/index.php?cmd=items&index={page_index}"

    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
    }

    try:
        print(f"  Fetching {url}...")
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')

        # Find the main table with items
        # Look for table that contains item data
        items = []

        # Find all table rows
        for table in soup.find_all('table'):
            rows = table.find_all('tr')

            for row in rows:
                cells = row.find_all('td')

                # Skip rows that don't have enough cells
                if len(cells) < 8:
                    continue

                # Extract text from cells
                row_text = [cell.get_text(strip=True) for cell in cells]

                # Check if this looks like an item row (has level number)
                if len(row_text) >= 2:
                    # Try to parse second column as level
                    try:
                        level = int(row_text[1])
                        # If we can parse level, this is likely an item row
                        items.append({
                            'Name': row_text[0] if len(row_text) > 0 else '',
                            'Level': row_text[1] if len(row_text) > 1 else '',
                            'Type': row_text[2] if len(row_text) > 2 else '',
                            'Rarity': row_text[3] if len(row_text) > 3 else '',
                            'Attack': row_text[4] if len(row_text) > 4 else '',
                            'Defense': row_text[5] if len(row_text) > 5 else '',
                            'Armor': row_text[6] if len(row_text) > 6 else '',
                            'Damage': row_text[7] if len(row_text) > 7 else '',
                            'HP': row_text[8] if len(row_text) > 8 else '',
                            'Stamina': row_text[9] if len(row_text) > 9 else '',
                            'Enhancements': row_text[10] if len(row_text) > 10 else ''
                        })
                    except (ValueError, IndexError):
                        # Not a valid item row
                        continue

        return pd.DataFrame(items)

    except Exception as e:
        print(f"  Error fetching page at index {page_index}: {e}")
        return pd.DataFrame()

def scrape_all_items(total_pages: int = 61) -> pd.DataFrame:
    """Scrape all pages of items"""
    all_items = []

    print(f"Starting scrape of {total_pages} pages...")
    print("=" * 50)

    for page in range(total_pages):
        print(f"Page {page + 1}/{total_pages}:")

        # Pages increment by 25
        page_index = page * 25
        df = scrape_items_page(page_index)

        if not df.empty:
            all_items.append(df)
            print(f"  Found {len(df)} items")
        else:
            print(f"  No items found")

        # Rate limiting
        if page < total_pages - 1:
            time.sleep(1.5)

    if all_items:
        return pd.concat(all_items, ignore_index=True)
    return pd.DataFrame()

def save_to_csv(df: pd.DataFrame, filename: str = "fallensword_items.csv"):
    """Save DataFrame to CSV"""
    df.to_csv(filename, index=False, encoding='utf-8')
    print(f"\nSaved {len(df)} items to {filename}")

def create_html_table(df: pd.DataFrame, filename: str = "fallensword_items.html"):
    """Create an interactive HTML file with DataTables"""

    # Convert DataFrame to HTML table rows
    table_rows = ""
    for _, row in df.iterrows():
        table_rows += "            <tr>\n"
        for col in df.columns:
            value = str(row[col]).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            table_rows += f"                <td>{value}</td>\n"
        table_rows += "            </tr>\n"

    # Create column headers
    headers = ""
    for col in df.columns:
        headers += f"                <th>{col}</th>\n"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Fallen Sword Items Database</title>

    <!-- DataTables CSS -->
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.7/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.2/css/buttons.dataTables.min.css">

    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>

    <!-- DataTables JS -->
    <script src="https://cdn.datatables.net/1.13.7/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.2/js/dataTables.buttons.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.2/js/buttons.html5.min.js"></script>

    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }}

        .container {{
            background: white;
            border-radius: 10px;
            padding: 30px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        }}

        h1 {{
            color: #333;
            text-align: center;
            margin-bottom: 30px;
            font-size: 2.5em;
        }}

        .stats {{
            text-align: center;
            margin-bottom: 20px;
            color: #666;
        }}

        #itemsTable {{
            width: 100% !important;
        }}

        .dataTables_wrapper {{
            margin-top: 20px;
        }}

        .dataTables_filter input {{
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 5px 10px;
            margin-left: 10px;
        }}

        table.dataTable thead th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: bold;
        }}

        table.dataTable tbody tr:hover {{
            background-color: #f5f5ff !important;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>⚔️ Fallen Sword Items Database ⚔️</h1>
        <div class="stats">
            <p>Total Items: <strong>{len(df)}</strong> | Last Updated: <strong>{pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}</strong></p>
        </div>

        <table id="itemsTable" class="display" style="width:100%">
            <thead>
                <tr>
{headers}                </tr>
            </thead>
            <tbody>
{table_rows}            </tbody>
        </table>
    </div>

    <script>
        $(document).ready(function() {{
            $('#itemsTable').DataTable({{
                "pageLength": 50,
                "lengthMenu": [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
                "scrollX": true,
                "order": [[1, "desc"]], // Sort by level descending
                "dom": 'Bfrtip',
                "buttons": [
                    'copy', 'csv', 'excel'
                ],
                "columnDefs": [
                    {{ "targets": [1, 4, 5, 6, 7, 8, 9], "type": "num" }} // Numeric columns
                ],
                "language": {{
                    "search": "Search items:",
                    "lengthMenu": "Show _MENU_ items per page",
                    "info": "Showing _START_ to _END_ of _TOTAL_ items",
                    "paginate": {{
                        "first": "First",
                        "last": "Last",
                        "next": "Next",
                        "previous": "Previous"
                    }}
                }}
            }});
        }});
    </script>
</body>
</html>"""

    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"Created HTML file: {filename}")

def main():
    print("\n🗡️  Fallen Sword Items Scraper 🗡️")
    print("=" * 50)

    # Check for test mode
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        print("\n📋 TEST MODE - Scraping first 3 pages only...")
        df = scrape_all_items(3)
    else:
        print("\n📋 FULL SCRAPE MODE - This will take several minutes...")
        print("Scraping all 61 pages with rate limiting...")
        df = scrape_all_items(61)

    if df.empty:
        print("\n❌ No items were scraped. Please check the website structure.")
        return

    print(f"\n✅ Total items scraped: {len(df)}")

    # Display sample data
    print("\n📊 Sample data (first 5 items):")
    print(df.head().to_string())

    # Save to CSV
    csv_file = "fallensword_items.csv"
    save_to_csv(df, csv_file)

    # Create HTML
    html_file = "fallensword_items.html"
    create_html_table(df, html_file)

    print("\n✨ Scraping complete! Files created:")
    print(f"  📄 {csv_file} - Import this to Google Sheets")
    print(f"  🌐 {html_file} - Open in browser for searchable/sortable table")
    print("\n💡 Tips:")
    print("  - The CSV file can be imported directly into Google Sheets")
    print("  - The HTML file works offline and has export buttons")
    print("  - You can host the HTML file on any web server or GitHub Pages")

if __name__ == "__main__":
    main()