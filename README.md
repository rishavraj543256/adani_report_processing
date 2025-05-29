# Adani Data Processing Suite

## Overview
The Adani Data Processing Suite is a comprehensive GUI application designed to process various types of data files for Adani operations. The application supports processing of master data, hygiene data, MB52 data, count sheets, stack data, and raw material data.

## Prerequisites
- Python 3.x
- Required Python packages (install using `pip install -r requirements.txt`):
  - tkinter
  - pandas
  - openpyxl
  - xlrd

## Installation
1. Clone or download this repository
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage Instructions

### Starting the Application
1. Run the main application:
   ```bash
   python gui.py
   ```

### Input Parameters
1. **Category**: Select either "Wheat" or "Paddy/Rice"
2. **S Loc Code**: Enter your location code

### File Processing Steps

#### 1. Master Data Processing
1. Click "Browse" to select your master file
2. Click "Process Data" to process the master data
3. The output will be saved in `output/format.xlsx`

#### 2. Hygiene Data Processing
1. Select your hygiene input file using the "Browse" button
2. Click "Process Hygiene" to process the hygiene data
3. Results will be updated in the format file

#### 3. MB52 Data Processing
1. Select your MB52 input file
2. Click "Process MB52" to process the MB52 data
3. Results will be updated in the format file

#### 4. Count Sheet Processing
1. Select your count sheet input file
2. Click "Process Count Sheet" to process the count sheet data
3. Results will be updated in the format file

#### 5. Stack Data Processing
1. Select your stack input file
2. Click "Process Stack" to process the stack data
3. Results will be updated in the format file

#### 6. Raw Material Processing
1. Click "Process Raw Material" to process raw material data
2. Results will be updated in the format file

### Output
- All processed data is saved in `output/format.xlsx`
- The console output area shows processing status and any errors
- Success/error messages are displayed in popup windows

### Troubleshooting
- Ensure all required input files are in the correct format
- Check that the S Loc Code is valid
- Verify that the selected category matches your data
- If errors occur, check the console output for detailed error messages

## Support
For any issues or questions, please contact the development team.

## License
This software is proprietary and confidential. Unauthorized copying, distribution, or use is strictly prohibited. 