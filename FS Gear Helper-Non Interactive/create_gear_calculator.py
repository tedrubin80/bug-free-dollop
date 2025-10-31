#!/usr/bin/env python3
import pandas as pd
import numpy as np

def create_gear_calculator_sheets():
    """Create a comprehensive gear calculator with multiple sheets for Google Sheets"""

    # Read the scraped items data
    try:
        df = pd.read_csv('fallensword_items.csv')
    except FileNotFoundError:
        print("Error: fallensword_items.csv not found. Please run the scraper first.")
        return

    # Clean numeric columns
    numeric_cols = ['Level', 'Attack', 'Defense', 'Armor', 'Damage', 'HP', 'Stamina']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    # Create Excel writer with multiple sheets
    with pd.ExcelWriter('fallensword_gear_calculator.xlsx', engine='openpyxl') as writer:

        # Sheet 1: Items Database
        df.to_excel(writer, sheet_name='Items Database', index=False)

        # Sheet 2: Gear Calculator
        calculator_data = {
            'Equipment Slot': ['Helmet', 'Armor', 'Gloves', 'Boots', 'Weapon', 'Shield',
                              'Ring 1', 'Ring 2', 'Amulet', 'Rune'],
            'Equipped Item': [''] * 10,
            'Level': ['=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,2,FALSE),0)'] * 10,
            'Attack': ['=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,5,FALSE),0)'] * 10,
            'Defense': ['=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,6,FALSE),0)'] * 10,
            'Armor': ['=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,7,FALSE),0)'] * 10,
            'Damage': ['=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,8,FALSE),0)'] * 10,
            'HP': ['=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,9,FALSE),0)'] * 10,
            'Stamina': ['=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,10,FALSE),0)'] * 10
        }

        # Update formulas for each row
        for i in range(1, 10):
            row_num = i + 2
            calculator_data['Level'][i] = f'=IFERROR(VLOOKUP(B{row_num},\'Items Database\'!A:K,2,FALSE),0)'
            calculator_data['Attack'][i] = f'=IFERROR(VLOOKUP(B{row_num},\'Items Database\'!A:K,5,FALSE),0)'
            calculator_data['Defense'][i] = f'=IFERROR(VLOOKUP(B{row_num},\'Items Database\'!A:K,6,FALSE),0)'
            calculator_data['Armor'][i] = f'=IFERROR(VLOOKUP(B{row_num},\'Items Database\'!A:K,7,FALSE),0)'
            calculator_data['Damage'][i] = f'=IFERROR(VLOOKUP(B{row_num},\'Items Database\'!A:K,8,FALSE),0)'
            calculator_data['HP'][i] = f'=IFERROR(VLOOKUP(B{row_num},\'Items Database\'!A:K,9,FALSE),0)'
            calculator_data['Stamina'][i] = f'=IFERROR(VLOOKUP(B{row_num},\'Items Database\'!A:K,10,FALSE),0)'

        calc_df = pd.DataFrame(calculator_data)

        # Add total row
        totals_row = pd.DataFrame({
            'Equipment Slot': ['TOTAL STATS'],
            'Equipped Item': [''],
            'Level': ['=MAX(C2:C11)'],
            'Attack': ['=SUM(D2:D11)'],
            'Defense': ['=SUM(E2:E11)'],
            'Armor': ['=SUM(F2:F11)'],
            'Damage': ['=SUM(G2:G11)'],
            'HP': ['=SUM(H2:H11)'],
            'Stamina': ['=SUM(I2:I11)']
        })

        calc_df = pd.concat([calc_df, totals_row], ignore_index=True)
        calc_df.to_excel(writer, sheet_name='Gear Calculator', index=False)

        # Sheet 3: Item Comparison
        comparison_data = {
            'Comparison': ['Item 1', 'Item 2', 'Difference'],
            'Item Name': ['', '', ''],
            'Level': [
                '=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,2,FALSE),0)',
                '=IFERROR(VLOOKUP(B3,\'Items Database\'!A:K,2,FALSE),0)',
                '=C3-C2'
            ],
            'Attack': [
                '=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,5,FALSE),0)',
                '=IFERROR(VLOOKUP(B3,\'Items Database\'!A:K,5,FALSE),0)',
                '=D3-D2'
            ],
            'Defense': [
                '=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,6,FALSE),0)',
                '=IFERROR(VLOOKUP(B3,\'Items Database\'!A:K,6,FALSE),0)',
                '=E3-E2'
            ],
            'Armor': [
                '=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,7,FALSE),0)',
                '=IFERROR(VLOOKUP(B3,\'Items Database\'!A:K,7,FALSE),0)',
                '=F3-F2'
            ],
            'Damage': [
                '=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,8,FALSE),0)',
                '=IFERROR(VLOOKUP(B3,\'Items Database\'!A:K,8,FALSE),0)',
                '=G3-G2'
            ],
            'HP': [
                '=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,9,FALSE),0)',
                '=IFERROR(VLOOKUP(B3,\'Items Database\'!A:K,9,FALSE),0)',
                '=H3-H2'
            ],
            'Stamina': [
                '=IFERROR(VLOOKUP(B2,\'Items Database\'!A:K,10,FALSE),0)',
                '=IFERROR(VLOOKUP(B3,\'Items Database\'!A:K,10,FALSE),0)',
                '=I3-I2'
            ],
            'Total Stats': [
                '=D2+E2+F2+G2+H2+I2',
                '=D3+E3+F3+G3+H3+I3',
                '=J3-J2'
            ]
        }

        comparison_df = pd.DataFrame(comparison_data)
        comparison_df.to_excel(writer, sheet_name='Item Comparison', index=False)

        # Sheet 4: Best Items by Type
        best_items = []

        for item_type in df['Type'].unique():
            if pd.notna(item_type) and item_type:
                type_items = df[df['Type'] == item_type].copy()
                if not type_items.empty:
                    # Calculate total stats
                    type_items['Total Stats'] = (
                        type_items['Attack'] +
                        type_items['Defense'] +
                        type_items['Armor'] +
                        type_items['Damage'] +
                        type_items['HP'] +
                        type_items['Stamina']
                    )
                    # Get top 5 items by total stats
                    top_items = type_items.nlargest(5, 'Total Stats')[
                        ['Name', 'Level', 'Type', 'Rarity', 'Attack', 'Defense',
                         'Armor', 'Damage', 'HP', 'Stamina', 'Total Stats']
                    ]
                    best_items.append(top_items)

        if best_items:
            best_items_df = pd.concat(best_items, ignore_index=True)
            best_items_df.to_excel(writer, sheet_name='Best Items by Type', index=False)

        # Sheet 5: Level Range Filter
        level_ranges = []
        ranges = [(1, 500), (501, 1000), (1001, 2000), (2001, 3000), (3001, 4000), (4001, 5000), (5001, 6000)]

        for min_lvl, max_lvl in ranges:
            range_items = df[(df['Level'] >= min_lvl) & (df['Level'] <= max_lvl)].copy()
            if not range_items.empty:
                range_items['Level Range'] = f'{min_lvl}-{max_lvl}'
                level_ranges.append(range_items.head(100))  # Top 100 items per range

        if level_ranges:
            level_range_df = pd.concat(level_ranges, ignore_index=True)
            level_range_df.to_excel(writer, sheet_name='Items by Level', index=False)

    print("Gear calculator Excel file created successfully!")

    # Also create a simplified CSV for direct Google Sheets import with instructions
    instructions = pd.DataFrame({
        'FALLEN SWORD GEAR CALCULATOR - INSTRUCTIONS': [
            'HOW TO USE THIS CALCULATOR:',
            '1. Import this file into Google Sheets (File > Import)',
            '2. The "Items Database" sheet contains all items',
            '3. In "Gear Calculator" sheet, type item names in column B to see stats',
            '4. Use "Item Comparison" to compare two items side by side',
            '5. "Best Items by Type" shows top items for each equipment type',
            '6. "Items by Level" helps find items in your level range',
            '',
            'FORMULAS INCLUDED:',
            '- Auto-lookup of item stats when you type item name',
            '- Total stats calculation for your full gear set',
            '- Difference calculation when comparing items',
            '- Automatic sorting by total stats for best items',
            '',
            'TIPS:',
            '- Use Data Validation to create dropdown lists of item names',
            '- Add conditional formatting to highlight stat improvements',
            '- Create additional sheets for different character builds',
            '- Use filters to find items with specific stats or enhancements'
        ]
    })

    with pd.ExcelWriter('fallensword_calculator_for_sheets.xlsx', engine='openpyxl') as writer:
        instructions.to_excel(writer, sheet_name='Instructions', index=False)
        df.to_excel(writer, sheet_name='Items Database', index=False)
        calc_df.to_excel(writer, sheet_name='Gear Calculator', index=False)
        comparison_df.to_excel(writer, sheet_name='Item Comparison', index=False)
        if best_items:
            best_items_df.to_excel(writer, sheet_name='Best Items', index=False)

def main():
    print("Creating Fallen Sword Gear Calculator for Google Sheets...")
    print("=" * 60)

    create_gear_calculator_sheets()

    print("\n✅ Files created:")
    print("  📊 fallensword_gear_calculator.xlsx - Full calculator with formulas")
    print("  📊 fallensword_calculator_for_sheets.xlsx - Optimized for Google Sheets")
    print("\n📝 To use in Google Sheets:")
    print("  1. Go to Google Sheets (sheets.google.com)")
    print("  2. Create a new spreadsheet")
    print("  3. Go to File > Import")
    print("  4. Upload 'fallensword_calculator_for_sheets.xlsx'")
    print("  5. Choose 'Replace spreadsheet' and click Import")
    print("\n💡 The calculator will automatically:")
    print("  - Look up item stats when you type item names")
    print("  - Calculate total stats for your equipment set")
    print("  - Compare items and show stat differences")
    print("  - Show best items organized by type and level")

if __name__ == "__main__":
    main()