# Japanese Invoice Template - Comprehensive Formatting Analysis

## Template File: /Users/fukku_maple/Downloads/invoice-template.xlsx

### Worksheet Overview
- **Name**: "請求書" (Invoice)
- **Dimensions**: 28 actual rows × 9 actual columns
- **Layout**: Traditional Japanese invoice format with sender/recipient sections and itemized table

---

## Column Specifications

| Column | Letter | Width (Excel units) | Purpose |
|--------|--------|---------------------|---------|
| 1 | A | 3.38 | Left margin/item content start |
| 2 | B | 12.63 | Main content area |
| 3 | C | 14.75 | Content continuation/client info |
| 4 | D | 7.5 | Quantity column |
| 5 | E | 2.25 | Spacer column |
| 6 | F | 9.88 | Unit price continuation |
| 7 | G | 12.63 | Amount/company info |
| 8 | H | 22.63 | Email/remarks |
| 9 | I | 3.38 | Right margin |

---

## Row Heights (Excel units)

| Row | Height | Content |
|-----|--------|---------|
| 1 | 17.25 | Issue date header |
| 2 | 40.5 | Main title "請求書" |
| 3 | 18.75 | Spacing |
| 4 | 18.75 | Client company name |
| 5 | 12 | Spacing |
| 6-9 | 14.25 | Address blocks |
| 10 | 6 | Small spacing |
| 11 | 14.25 | Email line |
| 12 | 18.75 | Company name |
| 13 | 24 | Large spacing |
| 14 | 18.75 | Request statement |
| 15 | 10.5 | Small spacing |
| 16 | 27 | Total amount (large) |
| 17 | 5.25 | Small spacing |
| 18 | 16.5 | Due date |
| 19 | 25.5 | Pre-table spacing |
| 20 | 22.5 | Table headers |
| 21-28 | 22.5 | Item rows |

---

## Key Design Elements

### 1. Header Section (Rows 1-2)
- **G1:I1**: Date field with custom format `yyyy" 年 "m" 月 "d" 日"`
- **A2:I2**: Merged title cell "請　　求　　書"
  - Font: Bold, 18pt
  - Alignment: Center, Middle
  - Character spacing gives formal appearance

### 2. Client Information (Left Side, Rows 4-8)
- **B4:C4**: Client company name with 様 (sama) honorific in D4
  - Bottom border: Thin black line
  - Font: 12pt
- **B6:C6**: Postal code "〒111-2222"
- **B7:C7**: Address line 1
- **B8:C8**: Address line 2 (apartment/building)

### 3. Sender Information (Right Side, Rows 7-12)
- **G7:H7**: Company postal code
- **G8:H8**: Company address line 1
- **G9:H9**: Company address line 2
- **H11**: Email address (right-aligned)
- **G12:H12**: Company name (right-aligned, 11pt)

### 4. Amount Summary (Rows 14-18)
- **B14:G14**: Request statement "下記の通りご請求申し上げます。"
- **C16:D16**: Total amount display
  - Font: Bold, 18pt
  - Format: ¥#,##0 (Japanese Yen)
  - Formula: =G29 (links to sum of items)
  - Alignment: Center, Middle
- **C18:E18**: Due date with custom format

### 5. Item Table (Rows 20-28)

#### Table Headers (Row 20)
All headers have:
- Background: Dark gray (#434343)
- Font: Bold, 11pt, White text
- Borders: Thin black lines
- Alignment: Center, Middle

| Column | Header Text | Merged Range |
|--------|-------------|--------------|
| A-C | "内　　　　　　容" (Content) | A20:C20 |
| D | "数量" (Quantity) | D20 |
| E-F | "単価（税込）" (Unit Price incl. Tax) | E20:F20 |
| G | "金　　額" (Amount) | G20 |
| H-I | "備　　　　　考" (Remarks) | H20:I20 |

#### Data Rows (Rows 21-28)
- **A21:C21**: Item description (merged, left-aligned)
- **D21**: Quantity (center-aligned, text format @)
- **E21:F21**: Unit price (¥#,##0 format)
- **G21**: Amount with formula: `=if(E21 = "" , "" , if(D21 = "" , 1*E21 , D21*E21))`
- **H21:I21**: Remarks (merged)

All data cells have:
- Font: 11pt
- Borders: Dotted bottom lines between rows
- Vertical alignment: Middle

### 6. Payment Information (Rows 31-33)
- **C32**: Bank account information
- **C33**: Account holder name with "名義：" prefix

---

## Critical Formatting Details

### Fonts
- **Main title**: Bold, 18pt
- **Headers**: Bold, 11pt, White color (#FFFFFF)
- **Data**: Regular, 11pt
- **Client name**: 12pt
- **Company info**: 10-11pt
- **Email**: 10pt, Google Sans Mono

### Colors
- **Header background**: #434343 (Dark gray)
- **Header text**: #FFFFFF (White)
- **Default text**: Black
- **Borders**: Black (#000000)

### Borders
- **Header cells**: Thin borders all around
- **Data cells**:
  - Dotted bottom borders between rows
  - Thin side borders
  - Medium left border on amount column (G) for emphasis
- **Client name**: Bottom border only

### Number Formats
- **Amounts**: `¥#,##0` (Japanese Yen with thousands separator)
- **Dates**: `yyyy" 年 "m" 月 "d" 日"` (Japanese date format)
- **Quantity**: `@` (Text format to preserve exact input)

### Merged Cell Ranges
Key merged ranges for layout:
- Title: A2:I2
- Client info: B4:C4, B6:C6, B7:C7, B8:C8
- Company info: G7:H7, G8:H8, G9:H9, G12:H12
- Total amount: C16:D16
- Item descriptions: A21:C21 through A28:C28
- Unit prices: E21:F21 through E28:F28
- Remarks: H21:I21 through H28:I28

### Formulas
- **Total display (C16)**: `=G29`
- **Item amounts (G21:G28)**: `=if(E21 = "" , "" , if(D21 = "" , 1*E21 , D21*E21))`
- **Grand total (G29)**: `=IF(SUM(G21:G28)=0,"",SUM(G21:G28))`

---

## Data Mapping for Template Filling

When preserving this template's exact formatting while filling with new data:

### Cell Addresses for Dynamic Content
- **G1**: Issue date (format as Japanese date)
- **B4:C4**: Client company name
- **B6:C6**: Client postal code
- **B7:C7**: Client address line 1
- **B8:C8**: Client address line 2
- **G7:H7**: Your company postal code
- **G8:H8**: Your company address line 1
- **G9:H9**: Your company address line 2
- **H11**: Your company email
- **G12:H12**: Your company name
- **C18:E18**: Payment due date
- **A21:C21**: Item 1 description
- **D21**: Item 1 quantity
- **E21:F21**: Item 1 unit price
- **A22:C22 through A28:C28**: Additional items
- **C32**: Bank account information
- **C33**: Account holder name

### Preservation Requirements
1. **Never modify**: Row heights, column widths, merged cell ranges
2. **Preserve exactly**: All font specifications, colors, borders, alignments
3. **Maintain**: All formulas in amount columns (G21:G29)
4. **Keep**: Japanese text formatting and spacing in headers
5. **Respect**: Number format codes for currency and dates

This analysis provides the complete blueprint for creating a system that preserves the exact visual appearance while only updating the data values.