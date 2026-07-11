# 🎉 PowerPoint VSTO Add-in - System Project PowerPoint Extension

![Status](https://img.shields.io/badge/status-Active-brightgreen)
![License](https://img.shields.io/badge/license-Educational-blue)
![Framework](https://img.shields.io/badge/.NET-4.7.2-purple)

## 📋 Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Quick Start](#quick-start)
- [Installation & Setup](#installation--setup)
- [How to Use](#how-to-use)
- [Project Structure](#project-structure)
- [Architecture](#architecture)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Overview

A comprehensive **PowerPoint VSTO Add-in** that extends Microsoft PowerPoint with advanced presentation tools including image processing, interactive quizzes, custom controls, and presentation enhancement features. This add-in provides educators and presenters with powerful tools to create more engaging and interactive presentations.

## Features

### 🖼️ Image Processing & Effects
- **Magnify Feature**: Zoom and magnify specific regions of your slides with customizable settings
- **Blur Feature**: Apply blur effects to sensitive or highlighted areas with adjustable intensity
- **Spotlight Feature**: Highlight important content with spotlight effects
- **Resize Lab**: Advanced image resizing and positioning tools
- **QR Code Control**: Generate and embed QR codes directly into presentations
- **Zoom Feature**: Enhanced zoom capabilities for better visibility

### 📚 Interactive Tools
- **Quiz Feature**: Create and manage interactive quizzes within presentations
  - Quiz questions management
  - Quiz pane controls for interactive testing
  - AI-powered quiz service for backend operations
- **Navigation Bar**: Customizable navigation controls with color settings
- **Input Dialog**: Easy user input collection for interactive elements
- **Agenda Generator**: Automatically generate presentation agendas from slide content

### 🎨 Presentation Enhancement
- **Slide Text Service**: Extract and manage slide text content
- **Position Lab Service**: Control and manage element positioning
- **Resize Lab Service**: Handle image and object resizing
- **Custom Ribbon Tab**: "My Tools" tab in the PowerPoint Ribbon with all features
- **Settings Forms**: Customizable settings for each feature
- **Visual Controls**: Dedicated user controls for each feature

## Quick Start

### Prerequisites
Before you begin, ensure you have the following installed:

| Component | Version | Required |
|-----------|---------|----------|
| Visual Studio | 2022 or later | ✅ |
| .NET Framework | 4.7.2 | ✅ |
| Microsoft Office | PowerPoint (Office 2016+) | ✅ |
| Office Developer Tools | Latest for Visual Studio | ✅ |
| VSTO Runtime | Latest | ✅ |

### Installation & Setup

#### Step 1: Clone the Repository
```powershell
# Using PowerShell
git clone https://github.com/RajorshiDas/System_Project_PowerPoint_Extention.git
cd System_Project_PowerPoint_Extention
```

Or using Git Bash:
```bash
git clone https://github.com/RajorshiDas/System_Project_PowerPoint_Extention.git
cd System_Project_PowerPoint_Extention
```

#### Step 2: Open the Solution
1. Open **Visual Studio 2022** (or later)
2. Go to `File` → `Open` → `Project/Solution`
3. Navigate to the cloned repository and select `PowerPointAddIn1.sln`
4. Wait for Visual Studio to load the project and restore dependencies

#### Step 3: Verify References
1. Right-click the project in Solution Explorer
2. Select `Properties`
3. Verify that **Target Framework** is set to `.NET Framework 4.7.2`
4. Check the `References` folder:
   - Ensure `Microsoft.Office.Interop.PowerPoint` is present
   - Ensure `Microsoft.Office.Tools.PowerPoint` is present
   - If missing, right-click References → `Add Reference` → Search for Office components

#### Step 4: Build the Project
```powershell
# Option 1: Using Visual Studio
# Press Ctrl+Shift+B or go to Build → Build Solution

# Option 2: Using PowerShell
dotnet build PowerPointAddIn1.csproj
```

#### Step 5: Launch and Test
1. Press `F5` to start debugging
2. Visual Studio will:
   - Compile the project
   - Install the add-in
   - Launch PowerPoint automatically
3. PowerPoint will open with the add-in loaded
4. Look for the **"My Tools"** tab in the PowerPoint Ribbon

## How to Use

### Using Each Feature

#### 🔍 Magnify Feature
1. Open or create a PowerPoint presentation
2. Add or select an image/content to magnify
3. Go to **My Tools** → **Magnify**
4. The magnification settings form will appear
5. Configure:
   - Zoom level (1x to 10x)
   - Display style
6. Click **Apply** to see the magnified view

#### 🎭 Blur Feature
1. Select content on your slide that you want to blur
2. Navigate to **My Tools** → **Blur**
3. In the Blur Settings Form:
   - Adjust blur intensity (Low, Medium, High)
   - Choose blur style if available
4. Click **Apply** to apply the blur effect
5. Use **Reset** to remove the blur effect

#### 💡 Spotlight Feature
1. Click **My Tools** → **Spotlight**
2. In the Spotlight Settings Form:
   - Choose spotlight size and intensity
   - Select spotlight shape (circle, rectangle, etc.)
3. Click on the area of your slide you want to highlight
4. The spotlight will emphasize the selected area
5. Adjust position and size as needed

#### 📝 Quiz Feature
1. Click **My Tools** → **Quiz**
2. In the Quiz Pane:
   - Click **Create New Quiz**
   - Add questions and answer options
   - Set the correct answer for each question
3. Once quiz is created:
   - Run quiz in **Presentation Mode** (F5)
   - Students can answer questions
   - View quiz results and statistics

#### 🔗 QR Code Feature
1. Position cursor where you want the QR code
2. Go to **My Tools** → **QR Code**
3. Enter the URL or text to encode
4. Click **Generate**
5. The QR code will be embedded in the slide
6. Resize and reposition as needed

#### 📐 Resize Lab
1. Select one or more images on your slide
2. Click **My Tools** → **Resize Lab**
3. In the Resize Lab Panel:
   - Use the positioning grid
   - Adjust width and height values
   - Use preset sizes if available
4. Click **Apply Changes**
5. Images will be resized and repositioned

#### 📋 Agenda Generator
1. Create your presentation with multiple slides
2. Go to **My Tools** → **Agenda Generator**
3. Select the slides to include in the agenda
4. Choose the agenda format (numbered list, outline, etc.)
5. Click **Generate Agenda**
6. An agenda slide will be created at the beginning of your presentation

#### 🧭 Navigation Bar
1. Click **My Tools** → **Navigation Settings**
2. In the Navigation Bar Settings:
   - Enable/disable navigation
   - Choose navigation style
   - Configure colors (via **Color Settings**)
3. Click **Apply**
4. Navigation controls will appear during presentation

## Project Structure

```
PowerPointAddIn1/
├── Core Components
│   ├── ThisAddIn.cs                 # Main add-in initialization
│   ├── ThisAddIn.Designer.cs        # Auto-generated add-in designer
│   ├── MyRibbon.cs                  # Ribbon UI and button handlers
│   └── MyRibbon.Designer.cs         # Auto-generated ribbon designer
│
├── Feature Services
│   ├── QuizAiService.cs             # AI-powered quiz operations
│   ├── SlideTextService.cs          # Slide content extraction
│   ├── ResizeLabService.cs          # Image resizing functionality
│   └── PositionLabService.cs        # Element positioning operations
│
├── UI Features
│   ├── MagnifyFeature.cs            # Magnification tool implementation
│   ├── BlurFeature.cs               # Blur effect implementation
│   ├── SpotlightFeature.cs          # Spotlight effect implementation
│   ├── ZoomFeature.cs               # Zoom functionality
│   ├── QRCodeControl.cs             # QR code generation
│   ├── AgendaGenerator.cs           # Agenda creation utility
│   └── InputDialog.cs               # Input dialog functionality
│
├── UI Controls
│   ├── QuizPaneControl.cs           # Quiz interface control
│   ├── QuizPaneControl.Designer.cs  # Quiz pane designer
│   ├── ResizeLabControl.cs          # Resize lab interface
│   └── PositionsLabControl.cs       # Position lab interface
│
├── Settings & Configuration
│   ├── MagnifySettingsForm.cs       # Magnify feature settings
│   ├── BlurSettingsForm.cs          # Blur effect settings
│   ├── SpotlightSettingsForm.cs     # Spotlight settings
│   ├── ResizeLabSettingsForm.cs     # Resize lab configuration
│   ├── NavBarSettings.cs            # Navigation bar settings
│   └── NavBarColorSettings.cs       # Navigation bar color config
│
├── Data Models
│   ├── QuizQuestion.cs              # Quiz question structure
│   └── QuizSet.cs                   # Quiz collection management
│
└── Properties
    ├── AssemblyInfo.cs
    ├── Resources.Designer.cs
    └── Settings.Designer.cs
```

## Architecture

The add-in follows a **modular, layered architecture**:

```
┌─────────────────────────────────────────────────┐
│         UI Layer (Ribbon & Controls)            │
│  (MyRibbon.cs, Settings Forms, Controls)       │
└────────────────────┬────────────────────────────┘
                     │
┌────────────────────▼────────────────────────────┐
│     Feature Layer (Feature Classes)             │
│  (Magnify, Blur, Spotlight, Quiz, etc.)        │
└────────────────────┬────────────────────────────┘
                     │
┌────────────────────▼────────────────────────────┐
│      Service Layer (Business Logic)             │
│  (QuizAiService, SlideTextService, etc.)       │
└────────────────────┬────────────────────────────┘
                     │
┌────────────────────▼────────────────────────────┐
│  Data & Models Layer (Data Structures)          │
│      (QuizQuestion, QuizSet, etc.)             │
└────────────────────┬────────────────────────────┘
                     │
┌────────────────────▼────────────────────────────┐
│  PowerPoint Interop (Office API Layer)          │
└─────────────────────────────────────────────────┘
```

### Design Principles
- **Separation of Concerns**: Each feature is self-contained
- **Single Responsibility**: Classes have one reason to change
- **DRY (Don't Repeat Yourself)**: Common functionality in services
- **Modularity**: Features can be added/removed independently
- **Extensibility**: Easy to add new features following existing patterns

## Configuration

### Feature Settings
Each feature stores its configuration in:
1. **Settings Forms** (e.g., `MagnifySettingsForm.cs`)
2. **Application Properties**: Persistent storage across sessions
3. **Ribbon Controls**: Quick access to common settings

### Example: Accessing Quiz Settings
```csharp
// QuizSet contains all quiz configuration
QuizSet quizzes = new QuizSet();
quizzes.AddQuestion(new QuizQuestion { Text = "Question?", Answer = "Answer" });
quizzes.SaveSettings();
```

### Customizing Features
1. Open the corresponding settings form
2. Modify the default values
3. Save settings (automatically persisted)
4. Close and reopen PowerPoint to apply changes

## Troubleshooting

### Common Issues

| Issue | Symptoms | Solution |
|-------|----------|----------|
| **Add-in doesn't load** | No "My Tools" tab visible | Check PowerPoint **Trust Center** settings (File → Options → Trust Center → Trust Center Settings → Trusted Locations) |
| **Build fails with reference errors** | "Cannot find Microsoft.Office.Interop" | Right-click project → Manage NuGet Packages, reinstall Office interop packages |
| **"Object reference not set" error** | Add-in crashes on button click | Ensure a presentation is open before using features |
| **QR Code not embedding** | QR code generation fails | Check that the slide has available space and is not in read-only mode |
| **Quiz pane not appearing** | Quiz feature unavailable | Try restarting Visual Studio and PowerPoint |
| **Resize Lab not working** | Images don't resize | Select valid image objects; some shapes cannot be resized |
| **Blur/Magnify effects not visible** | Effects applied but not showing | Verify Office Interop references are correct version |
| **Project won't compile** | Multiple compilation errors | Ensure .NET Framework 4.7.2 is installed and selected as target |

### Debug Steps
1. Enable **Debug Logging** in ThisAddIn.cs
2. Check Visual Studio **Output Window** (Debug → Windows → Output)
3. Add breakpoints and step through code
4. Inspect variable values in the **Watch Window**
5. Review **Application Event Log** for VSTO errors

### Getting Help
- Check existing [Issues](https://github.com/RajorshiDas/System_Project_PowerPoint_Extention/issues)
- Review code comments in relevant files
- Consult Microsoft VSTO documentation
- Enable verbose logging for detailed error information

## Requirements Summary

```
✅ Visual Studio 2022 or later
✅ .NET Framework 4.7.2
✅ Microsoft Office PowerPoint 2016 or later
✅ Office Developer Tools for Visual Studio (latest)
✅ VSTO Runtime (v4.0 or later)
✅ Administrator access (for installation)
```

## Error Handling & Safety

The add-in includes:
- ✓ Try-catch blocks in all critical operations
- ✓ User-friendly error messages
- ✓ Graceful degradation if features unavailable
- ✓ Trust Center security compatibility
- ✓ Input validation before processing
- ✓ Automatic error recovery mechanisms

## Best Practices Implemented

✓ **VSTO Architecture** - Following Microsoft guidelines  
✓ **Separation of Concerns** - Services, UI, Data models clearly separated  
✓ **Error Handling** - Comprehensive exception handling  
✓ **User Experience** - Intuitive UI with settings customization  
✓ **PowerPoint API** - Proper use of Office Interop APIs  
✓ **Code Organization** - Modular, maintainable structure  
✓ **Configuration Management** - Feature-level settings and persistence  
✓ **Security** - Trust Center compatible, no unsafe operations  

## Future Enhancements

- [ ] Advanced text formatting options in quiz
- [ ] Custom animation support
- [ ] Cloud storage integration
- [ ] Real-time collaboration features
- [ ] Analytics dashboard for quiz data
- [ ] Dark mode support
- [ ] Keyboard shortcuts customization
- [ ] Export quiz results to Excel/PDF
- [ ] Multi-language support
- [ ] Undo/Redo for all operations

## Contributing

We welcome contributions! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Guidelines
- Follow existing code style and conventions
- Add comments for complex logic
- Test thoroughly before submitting PR
- Update documentation if adding new features
- Ensure all references are resolved

## License

This project is provided for **educational and development purposes**.

## 📞 Support & Contact

- **Repository**: [System_Project_PowerPoint_Extention](https://github.com/RajorshiDas/System_Project_PowerPoint_Extention)
- **Issues**: [GitHub Issues](https://github.com/RajorshiDas/System_Project_PowerPoint_Extention/issues)
- **Author**: Rajorshi Das

---

**Made with ❤️ for educators and presenters**
