# Advanced Sample Tracker

An automated multi-site biorepository sample tracking system built in Microsoft Excel.

## Overview
Designed to track biological samples distributed across three international 
research sites: CDI, NIH, and Max Planck Institute. Samples are assigned 
unique auto-generated Study Codes and tracked through collection, freezing, and storage.

## Workbook Structure

| Sheet | Purpose |
|---|---|
| MASTER | Central registry — Study Code, demographics, collection/freeze dates, storage locations |
| CDI | 96-well plate map for CDI site |
| NIH | 96-well plate map for NIH site |
| Max Planck | 96-well plate map for Max Planck Institute |
| CRYOVIAL | 10×10 cryovial box storage tracker |
| SETUP | Dropdown validation lists and pre-generated sample ID pool |

## Key Excel Features Used
- Auto-generated unique Study Codes using `COUNTIFS()` and `TEXT()`
- Cross-sheet plate position mapping using `LET()`, `XLOOKUP()`, `INDEX/MATCH()`
- Automatic sample labeling with collection date formatting
- Data validation dropdowns for Age Range, Gender, and Ancestry
- Supports up to 1,944 samples across multiple plates and cryovial boxes

## Tools Used
- Microsoft Excel (365 / with LET and XLOOKUP support)

## Files
- `Final_Sample_Tracker.xlsx` — Full sample tracking workbook
