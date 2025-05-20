# ChessExcel VBA Project

**ChessExcel** is a simple chess game built entirely in Excel using VBA macros. The board is drawn with Unicode chess symbols, pieces are moved by selecting cells, and turns are managed via an ActiveX button. This repository contains the macro-enabled workbook (`ChessExcel.xlsm`) and all exported VBA modules for version control.

---

## Prerequisites & Excel Settings

Before running the game, ensure macros and ActiveX controls are enabled in your Excel:

1. **Show the Developer Tab**  
   - Go to **File > Options > Customize Ribbon**.  
   - Under **Main Tabs**, check **Developer** and click **OK**.

2. **Enable Macros**
   - Navigate to **File > Options > Trust Center > Trust Center Settings…**.  
   - In **Macro Settings**, select **Enable all macros**.  
   - Check **Trust access to the VBA project object model**.

4. **ActiveX Settings**
   - Navigate to **File > Options > Trust Center > Trust Center Settings…**.  
   - In **Trust Center**, go to **ActiveX Settings**.  
   - Select **Enable all controls without restrictions and without prompting**.

---

## Installation & Version Control

1. Clone* this repository:  
   ```bash
   git clone https://github.com/MarcBasas/ChessExcel.git
   ```  
2. Open `ChessExcel.xlsm` in Excel. Ensure the VBA modules (in `Modules/`) remain alongside the workbook for versioning.

*You can also download the .zip of the project

---

## How to Play

1. Open the **ChessExcel.xlsm** file in Excel.  
2. On the **Developer** tab, click the button **Play** 
3. The board will initialize with all pieces in starting positions.  
4. Click on a piece you want to move ( yes, starts with the whites. select origin).  
5. Click on the target square (select destination).  
6. The piece will move if the move is valid.  
7. After each move, click **Next Turn** to switch players.  
8. Repeat moves until checkmate.

---

## Project Structure

```
ChessExcel/
├── ChessExcel.xlsm        # Macro-enabled workbook
└── Modules/
    ├── Module1.bas        # Draws the board and initializes pieces
    ├── Module2.bas        # Returns piece identifiers. Movement logic and turn 
    ├── Sheet1.cls         # Worksheet event handlers
    └── ThisWorkbook.cls   # Workbook open initialization
```

---

## Code Integrity & Audit

All code in this repository is open and can be freely reviewed. There are no hidden or malicious components—feel free to inspect every module for peace of mind.

---

## License

This project is licensed under the MIT License.

---

Enjoy your game of Chess in Excel! Contributions and improvements are welcome.
