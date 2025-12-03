# Monthly Promotion Review

A business intelligence tool for analyzing promotion performance and tracking new vs existing member participation in promotional campaigns.

## Features

- **Member Type Analysis**: Automatically identifies new members (blank VIP Code) vs existing members
- **Promotion Performance Metrics**: Track quantity sold, revenue, discounts, and discount percentages
- **Visual Analytics**: Doughnut chart showing new vs existing member distribution
- **Search & Filter**: Quickly find specific promotions
- **Excel Export**: Export analysis results for further reporting

## Data Requirements

### Required Columns

The tool accepts Excel (.xlsx, .xls) or CSV files with the following columns:

- **Tx Date**: Transaction date
- **Promotion ID**: Unique promotion identifier
- **Promotion Desci**: Promotion description
- **Store Code**: Store location code
- **VIP Code**: Customer membership ID (blank = new member)
- **Doc No**: Document/receipt number
- **PLU Style**: Product SKU/style code
- **Item Description**: Product name
- **Brand**: Product brand
- **Animal Type**: Target pet type (DOG, CAT, etc.)
- **Product Group**: Product group classification
- **Product Class**: Product class
- **Product Category**: Category
- **Product Sub-Category**: Sub-category
- **Qty Sold**: Quantity sold
- **Amt Sold**: Sales amount
- **Prom Less**: Promotion discount amount
- **Ttl Sell Price**: Total selling price
- **Ttl Org Price**: Total original price

### Column Mapping

The tool supports flexible column naming. Alternative column names are automatically recognized:
- VIP Code: `vip code`, `vipcode`, `customer id`
- Promotion ID: `promotion id`, `promo id`, `promoid`
- Qty Sold: `qty sold`, `quantity`, `qty`
- Amt Sold: `amt sold`, `amount`, `revenue`, `sales`

## How to Use

1. **Open the Application**
   - Open `index.html` in your web browser
   - No installation or server required

2. **Upload Data**
   - Click "Choose File" and select your promotion data file
   - Supported formats: Excel (.xlsx, .xls) or CSV
   - Click "Process Data"

3. **Review Results**
   - View summary statistics in the top cards
   - Check the member type distribution chart
   - Browse detailed promotion performance in the table
   - Use the search box to filter specific promotions

4. **Export Results**
   - Click "Export to Excel" to download the analysis
   - File includes all calculated metrics

## Key Metrics

### New Members
Customers with blank/empty VIP Code fields. Each blank entry is counted as a unique new member transaction.

### Existing Members
Customers with a valid VIP Code. Unique customers are counted per promotion.

### Discount Percentage
Calculated as: (Total Discount / Total Original Price) × 100

### Revenue
Total sales amount after discounts (Amt Sold)

## Example Data Format

```
Tx Date,Promotion ID,Promotion Desci,Store Code,VIP Code,Qty Sold,Amt Sold,Prom Less,Ttl Org Price
2025-09-25,PEWHOUSE00003030,OctTop10折扣伊豆國寵物車,SHOP89,96451082,1,499.00,500.00,999.00
2025-09-26,PEWHOUSE00003030,OctTop10折扣伊豆國寵物車,RC0009,,1,499.00,500.00,999.00
```

## Technical Details

- **Technology**: Pure HTML5, CSS3, and vanilla JavaScript
- **Libraries**: 
  - SheetJS (xlsx) for Excel processing
  - Chart.js for data visualization
- **Browser Requirements**: Chrome 80+, Firefox 75+, Safari 13+, Edge 80+
- **No Backend Required**: Runs entirely in the browser

## Tips

- Blank VIP Codes are automatically treated as new members
- Each promotion is analyzed separately
- Results are sorted by revenue (highest first)
- Large files may take a few seconds to process
- All processing happens locally in your browser

## Troubleshooting

**File won't process**
- Ensure the file contains the required columns
- Check that column names match the expected format
- Verify the file is not corrupted

**Incorrect member counts**
- Verify VIP Code column contains actual blanks for new members
- Check for spaces or special characters in "blank" cells

**Chart not displaying**
- Ensure JavaScript is enabled in your browser
- Try refreshing the page
- Check browser console for errors

## Support

For issues or questions, refer to the main project documentation or contact your system administrator.
"# Monthly-Promotion-Analyser" 
