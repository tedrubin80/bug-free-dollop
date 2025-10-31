# Fallen Sword Gear Calculator

A comprehensive gear calculator and item database for Fallen Sword players. This tool scrapes all item data from the official Fallen Sword guide and provides an interactive calculator for optimizing character builds.

## Features

- **Complete Item Database**: 3,963+ items scraped from the official guide
- **Interactive Gear Calculator**: Build and optimize your character loadout
- **Item Comparison Tool**: Compare items side-by-side
- **Auto-calculation**: Instant stat totals when selecting equipment
- **Google Sheets Compatible**: Full integration with dropdown menus and formulas
- **Level-based Filtering**: Find items appropriate for your character level

## Files Included

| File | Description |
|------|-------------|
| `create_interactive_gear_calculator.py` | Generates the interactive calculator spreadsheets |
| `fallensword_google_sheets_calculator.xlsx` | Ready-to-use Google Sheets calculator |
| `fallensword_items.csv` | Raw CSV database of all items |
| `fallensword_items.html` | Searchable web-based item browser |

## Quick Start

### Using the Pre-built Calculator

1. Download `fallensword_google_sheets_calculator.xlsx`
2. Open [Google Sheets](https://sheets.google.com)
3. Create a new spreadsheet
4. File → Import → Upload the downloaded file
5. Choose "Replace spreadsheet"
6. Set up item dropdowns:
   - Select cells B4:B13 in Gear Calculator sheet
   - Data → Data validation
   - Set criteria to "Dropdown from Items Database!A:A"
7. Start selecting items and building your gear set!

### Updating Item Data

To fetch the latest items from the game:

```bash
# Install dependencies
pip install requests beautifulsoup4 pandas openpyxl

# Scrape latest item data (takes ~2 minutes)
python fallensword_scraper_improved.py

# Generate calculator spreadsheets
python create_interactive_gear_calculator.py
```

## Calculator Features

### Gear Calculator Sheet
- **10 Equipment Slots**: Helmet, Armor, Gloves, Boots, Weapon, Shield, 2 Rings, Amulet, Rune
- **Auto-lookup**: Type any item name to see its stats instantly
- **Total Stats**: Automatic calculation of combined stats
- **Smart Formulas**: VLOOKUP formulas pull data from the Items Database

### Item Comparison Sheet
- Compare any two items side-by-side
- See stat differences highlighted
- Calculate total stat advantages

### Items Database
- All 3,963+ game items with complete stats
- Sortable and filterable
- Includes: Name, Level, Type, Rarity, Attack, Defense, Armor, Damage, HP, Stamina

## How It Works

1. **Data Collection**: The scraper fetches all 61 pages from `guide.fallensword.com/index.php?cmd=items`
2. **Data Processing**: Items are parsed and cleaned into structured format
3. **Calculator Generation**: Excel formulas link item selection to stats
4. **Google Sheets Integration**: Data validation enables searchable dropdowns

## Example Usage

1. **Building a Tank Set**:
   - Select defensive items in each slot
   - Calculator shows total Defense and Armor
   - Compare different shield options

2. **Comparing Upgrades**:
   - Enter current weapon in Item 1
   - Enter potential upgrade in Item 2
   - See exact stat differences

3. **Level-Appropriate Gear**:
   - Filter Items Database by your level range
   - Find best items you can equip

## Requirements

- Python 3.7+
- Libraries: `requests`, `beautifulsoup4`, `pandas`, `openpyxl`
- Google Sheets account (for online calculator)

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/fallen-sword-gear-calculator.git
cd fallen-sword-gear-calculator

# Install Python dependencies
pip install -r requirements.txt

# Run the scraper to get latest data
python fallensword_scraper_improved.py

# Generate calculator spreadsheets
python create_interactive_gear_calculator.py
```

## Data Structure

Items are stored with the following attributes:
- **Name**: Item name
- **Level**: Required level to equip
- **Type**: Equipment slot type
- **Rarity**: Common, Rare, Unique, Legendary, Super Elite, Crystalline, Epic
- **Attack**: Attack stat bonus
- **Defense**: Defense stat bonus
- **Armor**: Armor stat bonus
- **Damage**: Damage stat bonus
- **HP**: Hit points bonus
- **Stamina**: Stamina bonus

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for:
- Bug fixes
- New features
- Calculator improvements
- Documentation updates

## License

This project is for educational purposes. Fallen Sword is owned by Hunted Cow Studios Ltd.

## Acknowledgments

- Data sourced from the official [Fallen Sword Guide](https://guide.fallensword.com)
- Built for the Fallen Sword community

## Support

If you find this tool helpful, please star the repository and share with other Fallen Sword players!

---

*Last updated: October 2024*
*Total items in database: 3,963*
