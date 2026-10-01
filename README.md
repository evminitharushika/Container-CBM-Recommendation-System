# Container CBM Recommendation System

## Purpose
A simple web system that:
1. Uploads an Excel file.
2. Reads pack/box CBM values.
3. Calculates total CBM automatically.
4. Compares total CBM with standard container capacities.
5. Recommends the smallest suitable container.
6. Shows utilization and remaining theoretical capacity.

## Supported containers
- **20 FT**: 33 CBM
- **40 FT**: 67 CBM
- **40 FT High Cube**: 76 CBM

These values are configurable in `app.js`.

## Excel format
The system automatically looks for common column names. Recommended example:

| Pack | Quantity | CBM |
| :--- | :--- | :--- |
| Pack 001 | 1 | 5.2 |
| Pack 002 | 1 | 4.8 |
| Pack 003 | 1 | 6.1 |

> **Note:** The CBM value is always treated as the total CBM for that row. Quantity is not multiplied into CBM again.

## Run
Open `index.html` in a browser. The Excel reader uses SheetJS from a CDN, so internet access is required when opening the page unless the SheetJS library is downloaded locally.

## Important limitation
This is a CBM-based recommendation system. It does not calculate exact physical 3D loading. Exact physical fit depends on dimensions, stacking, orientation, weight limits and loading rules.
