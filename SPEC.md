# PowerPoint Translation App Specification

## Project Overview
- **Project Name**: PPT Translator
- **Type**: Windows Desktop Application (Standalone EXE)
- **Core Functionality**: Translate PowerPoint text from Japanese to English without content shortening, with optional glossary/vocabulary support

## Requirements
1. Single EXE file - no installation required
2. Translate Japanese text in PowerPoint to English
3. Preserve original content length (no shortening)
4. Support custom glossary/vocabulary for consistent translations

## UI/UX Specification

### Layout Structure
- **Window Size**: 800x600 pixels (resizable, min 600x450)
- **Layout**: Single window with vertical sections
  - Header: App title and brief instructions
  - Main area: File selection and translation controls
  - Glossary section: Vocabulary management
  - Footer: Status bar

### Visual Design
- **Color Palette**:
  - Primary: #2196F3 (Blue)
  - Secondary: #4CAF50 (Green)
  - Background: #FAFAFA
  - Text: #212121
  - Accent: #FF9800 (Orange)
- **Typography**:
  - Font: Segoe UI (Windows default)
  - Headings: 16px bold
  - Body: 14px regular
- **Spacing**: 10px padding, 8px gaps

### Components
1. **File Selection Area**
   - "Select PowerPoint File" button
   - File path display label
   - Selected file info (slides count)

2. **Translation Controls**
   - "Start Translation" button (primary action)
   - Progress indicator
   - Output file path display

3. **Glossary Panel**
   - Add new term (Japanese | English) input fields
   - Add button
   - Glossary table (scrollable)
   - Delete button per row
   - Import/Export glossary buttons (CSV)

4. **Status Bar**
   - Current operation status
   - Error messages

## Functionality Specification

### Core Features
1. **PowerPoint Loading**
   - Support .pptx files
   - Read all text from shapes and text boxes
   - Handle multiple slides

2. **Translation**
   - Use Google Translate API (free endpoint)
   - Preserve original text length
   - Apply glossary substitutions before translation
   - Apply glossary substitutions after translation

3. **Glossary Management**
   - Add/remove vocabulary pairs
   - Import from CSV
   - Export to CSV
   - Persist glossary between sessions (save to file)

4. **Output**
   - Save as new .pptx file
   - Original file remains unchanged

### User Flow
1. User selects a PowerPoint file
2. User optionally adds glossary terms
3. User clicks "Start Translation"
4. App translates all text
5. App saves translated file
6. User sees success message

### Edge Cases
- Empty PowerPoint file
- No text in PowerPoint
- Network failure during translation
- Invalid glossary format

## Technical Stack
- Python 3.x
- python-pptx (PowerPoint manipulation)
- googletrans (Google Translate)
- PyInstaller (EXE packaging)

## Acceptance Criteria
1. EXE runs without Python installation
2. Successfully translates Japanese to English
3. Glossary terms are applied correctly
4. Original PowerPoint is not modified
5. UI is responsive and clear