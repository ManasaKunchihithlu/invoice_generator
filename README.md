# Invoice Generator 📄

A professional Python application that automatically generates beautiful PDF invoices from Excel data. Perfect for small businesses, freelancers, and anyone who needs to create multiple invoices quickly and efficiently.

## ✨ Features

- **Excel to PDF Conversion**: Reads invoice data from Excel and generates professional PDF invoices
- **Automatic Calculations**: Computes line totals, subtotals, taxes, discounts, and final amounts
- **Professional Design**: Clean, modern invoice layout with customizable branding
- **Batch Processing**: Generate multiple invoices from a single Excel file
- **Flexible Configuration**: Customize company details, logo, currency, and more
- **Multi-item Support**: Handle multiple line items per invoice
- **Tax & Discount Support**: Optional tax percentage and discount calculations

## 📋 Requirements

- Python 3.7 or higher
- openpyxl (for Excel file handling)
- reportlab (for PDF generation)

## 🚀 Installation

1. **Navigate to the project directory:**
   ```powershell
   cd C:\invoice_generator
   ```

2. **Install required packages:**
   ```powershell
   pip install -r requirements.txt
   ```

## 📊 Excel File Format

Your Excel file should have the following columns:

| Invoice Number | Customer Name | Address | Phone Number | Date | Item Name | Quantity | Price | Tax % | Discount % |
|----------------|---------------|---------|--------------|------|-----------|----------|-------|-------|------------|
| INV-001 | John Smith | 123 Main St | +1-555-0101 | 2025-11-27 | Laptop | 2 | 1200.00 | 8.5 | 5 |
| | | | | | Mouse | 2 | 25.00 | | |
| | | | | | Cable | 5 | 15.00 | | |
| INV-002 | Jane Doe | 456 Oak Ave | +1-555-0202 | 2025-11-26 | Desk | 1 | 450.00 | 7.5 | 0 |

**Important Notes:**
- The **first row with an Invoice Number** starts a new invoice
- **Subsequent rows** without an Invoice Number are treated as additional items for that invoice
- **Tax % and Discount %** only need to be specified once per invoice (first row)
- Empty cells in Tax % and Discount % columns default to 0

## 🎯 Quick Start

### 1. Create Sample Data (Optional)

Generate a sample Excel file to see the expected format:

```powershell
python create_sample_data.py
```

This creates `sample_invoices.xlsx` with 3 sample invoices.

### 2. Configure Your Company Details

Edit `config.json` to customize your company information:

```json
{
  "company_name": "Madalasa Enterprises",
  "company_address": "123 MG Road\nBengaluru, Karnataka 12345\nIndia",
  "company_phone": "+91 1234567890",
  "company_email": "info@madalasaenterprises.com",
  "company_website": "www.madalasaenterprises.com",
  "logo_path": "C:\\Users\\kmanas\\Python\\invoice_generator\\logo.jpg",
  "output_folder": "generated_invoices",
  "currency_symbol": "INR",
  "thank_you_note": "Thank you for your business! We appreciate your trust in us."
}
```

**Configuration Options:**
- `company_name`: Your business name (appears at top of invoice)
- `company_address`: Multi-line address (use `\n` for line breaks)
- `company_phone`: Contact phone number
- `company_email`: Contact email address
- `logo_path`: Path to your company logo image (PNG, JPG) - leave empty if no logo
- `output_folder`: Directory where PDFs will be saved
- `currency_symbol`: Currency symbol to display ($, €, £, etc.)
- `thank_you_note`: Custom message at the bottom of invoices

### 3. Add Your Logo (Optional)

1. Place your logo image (PNG or JPG) in the project folder
2. Update `logo_path` in `config.json`:
   ```json
   "logo_path": "logo.jpg"
   ```

### 4. Generate Invoices

Run the invoice generator with your Excel file:

```powershell
python invoice_generator.py sample_invoices.xlsx
```

Or with your own file:

```powershell
python invoice_generator.py your_invoice_data.xlsx
```

### 5. Find Your PDFs

Generated invoices will be saved in the `generated_invoices` folder:
```
generated_invoices/
  ├── Invoice_INV-001.pdf
  ├── Invoice_INV-002.pdf
  ├── Invoice_INV-002.pdf
  └── Invoice_INV-003.pdf

```

## 📖 Usage Examples

### Example 1: Basic Invoice

Excel data:
```
Invoice Number: INV-001
Customer: John Smith
Item: Laptop, Quantity: 1, Price: 1200
Tax: 8.5%
Discount: 0%
```

Result: `Invoice_INV-001.pdf` with calculated total of INR1,302.00

### Example 2: Multi-Item Invoice with Discount

Excel data:
```
Invoice Number: INV-002
Customer: Tech Corp
Item 1: Web Design, Quantity: 40 hours, Price: 150/hr
Item 2: Logo Design, Quantity: 1, Price: 500
Tax: 10%
Discount: 10%
```

Result: Invoice with subtotal INR6,500, discount -INR650, tax INR585, total INR6,435

### Example 3: Batch Processing

Create an Excel file with 10+ invoices and run:
```powershell
python invoice_generator.py bulk_invoices.xlsx
```

All invoices will be generated in seconds!

## 🎨 Invoice Design

Generated invoices include:

- **Header Section**: Company logo and details
- **Invoice Title**: Large "INVOICE" heading
- **Invoice Details**: Invoice number and date
- **Customer Info**: Bill-to information
- **Items Table**: Professional table with items, quantities, prices, and totals
- **Calculations Section**: 
  - Subtotal
  - Discount (if applicable)
  - Tax (if applicable)
  - Total Amount Due
- **Footer**: Thank you note and generation timestamp

## 🔧 Troubleshooting

### "File not found" Error
Make sure your Excel file exists and the path is correct:
```powershell
# Check if file exists
Test-Path your_file.xlsx
```

### Import Errors
Install the required packages:
```powershell
pip install openpyxl reportlab
```

### Logo Not Appearing
- Check that the logo file exists at the specified path
- Supported formats: PNG, JPG, JPEG
- Recommended size: 150x150 pixels or similar square aspect ratio

### Excel Format Issues
- Ensure the first row contains headers
- Invoice Number must be in the first column
- Numeric fields (Quantity, Price, Tax %, Discount %) should contain numbers only

## 📁 Project Structure

```
invoice_generator/
│
├── invoice_generator.py       # Main application
├── create_sample_data.py      # Sample data generator
├── config.json                # Configuration file
├── requirements.txt           # Python dependencies
├── README.md                  # This file
│
├── sample_invoices.xlsx       # Sample input (generated)
│
└── generated_invoices/        # Output folder (auto-created)
    ├── Invoice_INV-001.pdf
    ├── Invoice_INV-002.pdf
    └── ...
```

## 💡 Tips & Best Practices

1. **Consistent Formatting**: Keep your Excel data clean and consistent
2. **Unique Invoice Numbers**: Use unique invoice numbers for each invoice
3. **Backup**: Keep your Excel files as records
4. **Test First**: Use the sample data generator to test before processing real data
5. **Logo Size**: Use a square logo (150x150px recommended) for best results
6. **Currency**: Update the currency symbol in config.json for your region

## 🔄 Advanced Usage

### Custom Excel Layout

If your Excel file has a different structure, you can modify the `read_excel_data` method in `invoice_generator.py` to match your column names.

### Styling Customization

Modify the color scheme by changing the HexColor values in `invoice_generator.py`:
- Header color: `#34495E`
- Invoice title: `#E74C3C`
- Thank you note: `#27AE60`

### Add More Fields

You can extend the application to include:
- Payment terms
- Due date
- Payment methods
- Notes/comments
- PO numbers

Simply add columns to your Excel file and modify the code accordingly.

## 📝 Sample Invoice Preview

```
┌─────────────────────────────────────────────────────────┐
│                  YOUR COMPANY NAME                      │
│              123 Business Street                        │
│              Phone: +1 (555) 123-4567                  │
│              Email: info@yourcompany.com               │
│                                                         │
│                      INVOICE                            │
│                                                         │
│  Invoice Number: INV-001        Date: 2025-11-27      │
│                                                         │
│  Bill To:                                              │
│  John Smith                                            │
│  123 Main St, New York, NY 10001                       │
│  Phone: +1-555-0101                                    │
│                                                         │
│  ┌──────────────┬──────┬────────┬─────────────┐      │
│  │ Item         │ Qty  │ Price  │ Total       │      │
│  ├──────────────┼──────┼────────┼─────────────┤      │
│  │ Laptop       │ 2    │ INR1200  │ INR2,400.00   │      │
│  │ Mouse        │ 2    │ INR25    │ INR50.00      │      │
│  │ Cable        │ 5    │ INR15    │ INR75.00      │      │
│  └──────────────┴──────┴────────┴─────────────┘      │
│                                                         │
│                          Subtotal: INR2,525.00           │
│                     Discount (5%): -INR126.25            │
│              Subtotal after Discount: INR2,398.75        │
│                         Tax (8.5%): INR203.89            │
│                  Total Amount Due: INR2,602.64           │
│                                                         │
│          Thank you for your business!                  │
└─────────────────────────────────────────────────────────┘
```

## 📞 Support

If you encounter any issues:
1. Check that your Excel file matches the expected format
2. Verify all dependencies are installed
3. Review the configuration in `config.json`
4. Check the console output for specific error messages

## 🚀 Future Enhancements

Potential features for future versions:
- Web interface for uploading Excel files
- Email integration to send invoices automatically
- Multiple currency support
- Multiple tax rates per invoice
- Payment tracking
- Invoice templates selection
- QR code generation for payment

## 📄 License

This project is provided as-is for personal and commercial use.

---

**Happy Invoicing! 🎉**

For questions or feedback, please update the contact information in your config.json.
