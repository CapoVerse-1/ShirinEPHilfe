# Promotion Planner

A React application for converting promotion schedule Excel files into a workable format for promotion companies.

## Features

- **Drag & Drop Interface**: Simply drag and drop your Excel files
- **Automatic Processing**: Converts complex promotion schedules into simplified format
- **Excel Export**: Export processed data as Excel files
- **Real-time Preview**: See processed data before exporting
- **German Localization**: Day names and interface in German

## How it Works

### Input Format
The application expects Excel files with the following structure:
- Column A: Market name (Point of Sales)
- Column B: District
- Column C: Argentur (ignored)
- Column D: AG (ignored)
- Columns E-AI: Days 01.Jul to 31.Jul

### Values in Date Columns
- `1` = 8-hour promotion
- `2` = Two 8-hour promotions (creates 2 separate entries)
- `0.75` = 6-hour promotion
- Empty/0 = No promotion

### Output Format
The processed Excel will contain:
- Column A: Day of week (Mo, Di, Mi, Do, Fr, Sa, So)
- Column B: Date (DD.MM.YYYY)
- Column C: Start time (Mo-Fr: 9:30, Sa-So: 9:00)
- Column D: End time (Mo-Fr: 18:30, Sa-So: 18:00)
- Column E: Total hours worked
- Column F: Point of sales
- Column G: District
- Column H: Market name (without MM prefix)
- Column I: Coffee advisor (empty)

## Getting Started

### Prerequisites
- Node.js 18+ 
- npm or yarn

### Installation

1. Clone the repository
2. Install dependencies:
```bash
npm install
```

3. Run the development server:
```bash
npm run dev
```

4. Open [http://localhost:3000](http://localhost:3000) in your browser

### Building for Production

```bash
npm run build
npm start
```

### Deploy to Vercel

1. Push your code to GitHub
2. Connect your repository to Vercel
3. Deploy automatically

Or use the Vercel CLI:
```bash
npx vercel
```

## Usage

1. Open the application in your browser
2. Drag and drop your Excel file into the upload area
3. Review the processed data in the preview table
4. Click "Export Excel" to download the converted file

## Technical Details

- **Framework**: Next.js 14 with TypeScript
- **Excel Processing**: SheetJS (xlsx library)
- **File Handling**: react-dropzone
- **Styling**: CSS with modern gradients and responsive design
- **Export**: file-saver for client-side downloads

## Support

For issues or questions, please check the code comments or create an issue in the repository. 