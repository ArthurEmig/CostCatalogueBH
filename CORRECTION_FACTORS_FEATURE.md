# Correction Factors Feature Implementation

## Overview
This feature adds correction factors functionality to the CostsViewer application, allowing users to set year-based correction factors (1999 to current year) that are applied to project costs. The feature provides a comprehensive overview of original values, correction factors, and corrected values.

## New Components Added

### 1. Models
- **`Models/CorrectionFactor.cs`**: Individual correction factor model with Year and Factor properties
- **`CorrectionFactorSettings`**: Container for all correction factors with helper methods

### 2. Services
- **`Services/CorrectionFactorService.cs`**: Handles in-memory storage of correction factors (no external files needed)

### 3. Views & ViewModels
- **`Views/CorrectionFactorSettingsWindow.xaml`**: Settings dialog for managing correction factors
- **`Views/CorrectionFactorSettingsWindow.xaml.cs`**: Code-behind for the settings window
- **`ViewModels/CorrectionFactorSettingsViewModel.cs`**: ViewModel for the settings dialog with commands for:
  - Reset to Default (all factors = 1.0)
  - Apply Inflation (2% per year from 1999)
  - OK/Cancel operations

### 4. Updated Components

#### MainViewModel.cs
- Added correction factor settings loading and management
- Added corrected average properties for all cost groups (KG220-KG550)
- Added `CorrectionFactorSettingsCommand` to open settings dialog
- Updated `UpdateAverages()` to calculate both original and corrected values
- Added `CalculateCorrectedAverages()` method to apply correction factors

#### MainWindow.xaml
- Added "Correction Factors" button to toolbar
- Enhanced summary section to show:
  - Original averages
  - Corrected averages (highlighted in blue)
  - Difference indicators
- Added second tab "Corrected Averages" in right panel

#### Export Services
- **ExcelExporter.cs**: Enhanced to include:
  - Correction factors column in Projects worksheet
  - Original and corrected values for all cost groups
  - New "Correction Factors" worksheet with year-by-year factors
- **PdfExporter.cs**: Enhanced to include:
  - Correction factors section showing non-default factors
  - Project listings with correction factors
  - Year and factor information in project details

## Features

### 1. Settings Management
- **Access**: Click "Correction Factors" button in toolbar
- **Interface**: DataGrid showing Year, Correction Factor, and Percentage
- **Operations**:
  - Manual editing of individual factors
  - "Reset to Default" - sets all factors to 1.0
  - "Apply Inflation" - applies 2% yearly inflation from 1999

### 2. Data Persistence
- Correction factors are stored in memory during application runtime
- Settings are included in Excel exports for documentation
- Automatic initialization of all years from 1999 to current year
- Default factor of 1.0 for all years
- No external configuration files needed for single .exe deployment

### 3. Visual Indicators
- **Main UI**: 
  - Original values in standard black text
  - Corrected values in blue with light blue background
  - Separate tabs for "Original Averages" and "Corrected Averages"
- **Summary Section**: Shows both original and corrected averages side by side

### 4. Export Integration
- **Excel Export**: 
  - Projects sheet includes correction factors and corrected values
  - Dedicated "Correction Factors" worksheet
  - Clear separation between original and corrected data
- **PDF Export**: 
  - Correction factors section (shows only non-default factors)
  - Project details include year and correction factor
- **CSV Export**: Includes correction factors and corrected values

### 5. Calculations
- Correction factors are applied per project based on project year
- Corrected values = Original values × Correction factor for project year
- Averages are calculated from corrected individual project values
- All calculations maintain precision with proper rounding for display

## Usage Workflow

1. **Load Data**: Import project data as usual
2. **Set Correction Factors**: 
   - Click "Correction Factors" button
   - Adjust factors manually or use "Apply Inflation"
   - Click OK to save
3. **View Results**: 
   - Main summary shows both original and corrected averages
   - Switch between "Original Averages" and "Corrected Averages" tabs
4. **Export Data**: 
   - Excel/PDF/CSV exports include both original and corrected values
   - Correction factors are documented in exports

## Technical Implementation

### Data Flow
1. Correction factors initialized with default values on startup
2. Factors applied during average calculations in MainViewModel
3. UI updates automatically via data binding
4. Export services access the same in-memory factors for consistency

### Error Handling
- Graceful fallback to default factors if loading fails
- Validation ensures factors are positive numbers
- Console logging for debugging

### Performance
- Correction factors cached in memory
- Calculations performed only when data changes
- Efficient year-based lookup using Dictionary

## Single .exe Deployment

This implementation is designed for single .exe deployment:
- No external configuration files required
- Correction factors stored in memory during runtime
- Settings preserved in Excel exports for documentation
- All functionality works without any external dependencies

This provides a comprehensive correction factors system that integrates seamlessly with the existing CostsViewer functionality while maintaining data integrity and user experience, perfect for standalone executable distribution.
